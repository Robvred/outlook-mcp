"""Live mailbox queries through the local Outlook COM object model.

One engine for listing, searching and bulk selection. Nothing here goes through
the network or the Graph API: it only talks to the running Outlook client.

How it works
- Criteria are turned into one DASL filter evaluated by Outlook itself.
- Rows are read with ``Folder.GetTable`` (only the requested columns, no
  MailItem is opened), which is one to two orders of magnitude faster than
  walking ``Folder.Items``. If the table cannot be used on a store, the engine
  falls back to ``Items.Restrict`` with the same filter.
- Category criteria are always re-checked on the returned rows, so the result
  never depends on how a store evaluates multi-valued properties.
- Every row carries its EntryID and StoreID: an item can be reopened reliably
  later, whatever the store it lives in.
"""

from __future__ import annotations

import binascii
import re
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from .logging_config import get_logger

logger = get_logger(__name__)

# DASL property names (quoted, ready to embed in a filter)
P_SUBJECT = '"urn:schemas:httpmail:subject"'
P_FROM_NAME = '"urn:schemas:httpmail:fromname"'
P_FROM_EMAIL = '"urn:schemas:httpmail:fromemail"'
P_SENDER_SMTP = '"http://schemas.microsoft.com/mapi/proptag/0x5D01001F"'
P_RECEIVED = '"urn:schemas:httpmail:datereceived"'
P_KEYWORDS = '"urn:schemas-microsoft-com:office:office#Keywords"'
P_DISPLAY_TO = '"urn:schemas:httpmail:displayto"'
P_DISPLAY_CC = '"urn:schemas:httpmail:displaycc"'
P_BODY = '"urn:schemas:httpmail:textdescription"'
P_READ = '"urn:schemas:httpmail:read"'
P_TODO_FLAGS = '"http://schemas.microsoft.com/mapi/proptag/0x0E2B0003"'

# Table columns. Required ones are known to work on Exchange and PST stores;
# optional ones are added one by one and simply left out when a store refuses them.
REQUIRED_COLUMNS: Tuple[Tuple[str, str], ...] = (
    ("entry_id", "EntryID"),
    ("subject", "Subject"),
    ("sender", "SenderName"),
    ("received", "ReceivedTime"),
    ("categories", "Categories"),
    ("message_class", "MessageClass"),
    ("to", "urn:schemas:httpmail:displayto"),
    ("cc", "urn:schemas:httpmail:displaycc"),
)
OPTIONAL_COLUMNS: Tuple[Tuple[str, str], ...] = (
    ("sender_smtp", "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"),
    ("sender_address", "SenderEmailAddress"),
    ("message_flags", "http://schemas.microsoft.com/mapi/proptag/0x0E070003"),
    ("flag_status", "http://schemas.microsoft.com/mapi/proptag/0x10900003"),
    ("has_attach", "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"),
    ("size", "Size"),
    ("importance", "Importance"),
    ("modified", "LastModificationTime"),
)

MSGFLAG_READ = 0x1
OL_FLAG_MARKED = 2
_NO_DATE_YEAR = 4501
_CHUNK = 500
_HEX = re.compile(r"^[0-9A-Fa-f]{16,}$")


class QueryError(Exception):
    """The query cannot be honoured as asked (never silently relaxed)."""


def clean_terms(values: Optional[Iterable[Any]]) -> List[str]:
    """Strip and drop empty terms. An empty term would match every message."""
    out: List[str] = []
    for v in values or []:
        if isinstance(v, str) and v.strip() and v.strip() not in out:
            out.append(v.strip())
    return out


def split_categories(raw: Any) -> List[str]:
    """Categories as a list, whatever the list separator of the Windows locale."""
    if not raw:
        return []
    if isinstance(raw, (list, tuple)):
        return [str(c).strip() for c in raw if str(c).strip()]
    return [c.strip() for c in re.split(r"[;,]", str(raw)) if c.strip()]


def parse_search_terms(text: str) -> List[str]:
    """Split a free-text query into terms; "double quoted parts" stay one phrase."""
    terms: List[str] = []
    for quoted, bare in re.findall(r'"([^"]*)"|(\S+)', text or ""):
        term = (quoted or bare).strip()
        if term and term not in terms:
            terms.append(term)
    return terms


@dataclass
class MailCriteria:
    """Criteria are combined with AND. Terms of one list are OR unless match_all_terms."""

    days: Optional[int] = None
    older_than_days: Optional[int] = None
    received_from: Optional[datetime] = None  # local time, inclusive
    received_to: Optional[datetime] = None  # local time, exclusive
    subject: List[str] = field(default_factory=list)
    sender: List[str] = field(default_factory=list)
    recipient: List[str] = field(default_factory=list)
    body: List[str] = field(default_factory=list)
    match_all_terms: bool = False
    categories: List[str] = field(default_factory=list)
    categories_match_all: bool = False
    exclude_categories: List[str] = field(default_factory=list)
    unread_only: bool = False
    flagged_only: bool = False

    def __post_init__(self) -> None:
        for name in ("subject", "sender", "recipient", "body", "categories", "exclude_categories"):
            setattr(self, name, clean_terms(getattr(self, name)))
        for name in ("days", "older_than_days"):
            value = getattr(self, name)
            if value is not None:
                if isinstance(value, bool) or not isinstance(value, int) or value < 0:
                    raise ValueError(f"{name} must be a non-negative integer")
                if value == 0:
                    setattr(self, name, None)
        if self.days and self.older_than_days and self.older_than_days >= self.days:
            raise ValueError("older_than_days must be smaller than days (the window would be empty)")
        if self.received_from and self.received_to and self.received_to <= self.received_from:
            raise ValueError("The end date must be after the start date")

    def is_empty(self) -> bool:
        return not (
            self.days or self.older_than_days or self.received_from or self.received_to
            or self.subject or self.sender or self.recipient or self.body
            or self.categories or self.exclude_categories or self.unread_only or self.flagged_only
        )

    def describe(self) -> Dict[str, Any]:
        """Only the criteria that are set, for responses and the audit log."""
        out: Dict[str, Any] = {}
        for name in ("days", "older_than_days", "subject", "sender", "recipient", "body",
                     "categories", "exclude_categories"):
            if getattr(self, name):
                out[name] = getattr(self, name)
        if self.received_from:
            out["received_from"] = self.received_from.strftime("%Y-%m-%d %H:%M")
        if self.received_to:
            out["received_to"] = self.received_to.strftime("%Y-%m-%d %H:%M")
        if self.categories and self.categories_match_all:
            out["categories_match"] = "all"
        if self.match_all_terms:
            out["match_all_terms"] = True
        if self.unread_only:
            out["unread_only"] = True
        if self.flagged_only:
            out["flagged_only"] = True
        return out


@dataclass
class QueryResult:
    emails: List[Dict[str, Any]] = field(default_factory=list)
    total_matched: int = 0  # matches found, before max_results is applied
    truncated: bool = False  # more matches exist than returned
    complete: bool = True  # False when the scan stopped on the time or row limit
    method: str = "table"
    filter: str = ""
    folders: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    elapsed_s: float = 0.0


# --------------------------------------------------------------------------- filter

def _q(term: str) -> str:
    return term.replace("'", "''")


def _utc(value: datetime) -> str:
    """DASL compares dates in UTC."""
    if value.tzinfo is None:
        value = value.astimezone()  # naive = local time
    return value.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M")


def _like_group(props: Sequence[str], terms: Sequence[str], match_all: bool) -> str:
    per_term = ["(" + " OR ".join(f"{p} LIKE '%{_q(t)}%'" for p in props) + ")" for t in terms]
    return "(" + (" AND " if match_all else " OR ").join(per_term) + ")"


def build_filter(criteria: MailCriteria, now: Optional[datetime] = None,
                 dates_only: bool = False) -> str:
    """DASL filter for the criteria ('' when there is none).

    Excluded categories are deliberately not part of the filter: they are checked
    on the rows, which is exact and does not depend on how a store evaluates NOT
    on a property that most messages do not have.
    """
    now = now or datetime.now(timezone.utc)
    parts: List[str] = []
    if criteria.days:
        parts.append(f"{P_RECEIVED} >= '{_utc(now - timedelta(days=criteria.days))}'")
    if criteria.older_than_days:
        parts.append(f"{P_RECEIVED} < '{_utc(now - timedelta(days=criteria.older_than_days))}'")
    if criteria.received_from:
        parts.append(f"{P_RECEIVED} >= '{_utc(criteria.received_from)}'")
    if criteria.received_to:
        parts.append(f"{P_RECEIVED} < '{_utc(criteria.received_to)}'")
    if not dates_only:
        all_terms = criteria.match_all_terms
        if criteria.subject:
            parts.append(_like_group([P_SUBJECT], criteria.subject, all_terms))
        if criteria.sender:
            parts.append(_like_group([P_FROM_NAME, P_FROM_EMAIL, P_SENDER_SMTP], criteria.sender, all_terms))
        if criteria.recipient:
            parts.append(_like_group([P_DISPLAY_TO, P_DISPLAY_CC], criteria.recipient, all_terms))
        if criteria.body:
            parts.append(_like_group([P_BODY], criteria.body, all_terms))
        if criteria.categories:
            joiner = " AND " if criteria.categories_match_all else " OR "
            parts.append("(" + joiner.join(f"{P_KEYWORDS} = '{_q(c)}'" for c in criteria.categories) + ")")
        if criteria.unread_only:
            parts.append(f"{P_READ} = 0")
        if criteria.flagged_only:
            parts.append(f"{P_TODO_FLAGS} = 1")
    return "@SQL=" + " AND ".join(parts) if parts else ""


# --------------------------------------------------------------------------- rows

def _text(value: Any) -> str:
    return value if isinstance(value, str) else ""


def _small_int(value: Any, maximum: int) -> Optional[int]:
    """Integer column value; None for a missing property (COM returns None or an error code)."""
    if isinstance(value, bool) or not isinstance(value, int):
        return None
    return value if 0 <= value <= maximum else None


def _wall_clock(value: Any) -> Optional[datetime]:
    """COM date as a naive local datetime (None for Outlook's 'no date' sentinel)."""
    try:
        if value is None or value.year >= _NO_DATE_YEAR:
            return None
        return datetime(value.year, value.month, value.day, value.hour, value.minute, value.second)
    except Exception:
        return None


def _entry_id(value: Any) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, (bytes, bytearray, memoryview)):
        return binascii.hexlify(bytes(value)).decode("ascii").upper()
    return ""


def _names(display: str) -> List[Dict[str, str]]:
    return [{"name": n.strip(), "address": ""} for n in display.split(";") if n.strip()]


def folder_display_path(folder: Any) -> str:
    """'\\\\user@company.com\\Inbox\\ProjectA' -> 'user@company.com/Inbox/ProjectA'."""
    try:
        return str(folder.FolderPath).lstrip("\\").replace("\\", "/")
    except Exception:
        return _text(getattr(folder, "Name", "")) or "?"


def _make_row(values: Dict[str, Any], folder_path: str, store_id: str) -> Optional[Dict[str, Any]]:
    entry_id = _entry_id(values.get("entry_id"))
    if not entry_id:
        return None
    received = _wall_clock(values.get("received")) or _wall_clock(values.get("modified"))
    flags = _small_int(values.get("message_flags"), 0x7FFFFFFF)
    unread = values.get("unread")
    if unread is None and flags is not None:
        unread = not (flags & MSGFLAG_READ)
    sender_email = _text(values.get("sender_smtp"))
    if not sender_email and "@" in _text(values.get("sender_address")):
        sender_email = _text(values.get("sender_address"))
    has_attach = values.get("has_attach")
    to, cc = _text(values.get("to")), _text(values.get("cc"))
    return {
        "entry_id": entry_id,
        "store_id": store_id,
        "folder": folder_path,
        "subject": _text(values.get("subject")) or "(no subject)",
        "sender": _text(values.get("sender")) or "(unknown)",
        "sender_email": sender_email,
        "received_time": received.strftime("%Y-%m-%d %H:%M:%S") if received else "Unknown",
        "to_recipients": _names(to),
        "cc_recipients": _names(cc),
        "categories": ", ".join(split_categories(values.get("categories"))),
        "unread": bool(unread) if unread is not None else False,
        "flag_status": _small_int(values.get("flag_status"), 2) or 0,
        "importance": _small_int(values.get("importance"), 2),
        "size": _small_int(values.get("size"), 0x7FFFFFFF) or 0,
        "message_class": _text(values.get("message_class")),
        "has_attachments": bool(has_attach) if isinstance(has_attach, (bool, int)) and has_attach in (0, 1, True, False) else False,
        "attachments": [],
        "attachments_processed": False,
        "_known": {
            "unread": unread is not None,
            "flag_status": _small_int(values.get("flag_status"), 2) is not None,
        },
    }


def _contains(haystacks: Sequence[str], terms: Sequence[str], match_all: bool) -> bool:
    lowered = [h.lower() for h in haystacks if h]
    hits = [any(t.lower() in h for h in lowered) for t in terms]
    return all(hits) if match_all else any(hits)


def row_matches(row: Dict[str, Any], criteria: MailCriteria, check_text: bool = False) -> bool:
    """Exact check of a row. Text criteria are only checked when the filter could not carry them."""
    cats = {c.lower() for c in split_categories(row.get("categories"))}
    if criteria.categories:
        wanted = [c.lower() for c in criteria.categories]
        hit = all(c in cats for c in wanted) if criteria.categories_match_all else any(c in cats for c in wanted)
        if not hit:
            return False
    if criteria.exclude_categories and any(c.lower() in cats for c in criteria.exclude_categories):
        return False
    known = row.get("_known", {})
    if criteria.unread_only and known.get("unread") and not row.get("unread"):
        return False
    if criteria.flagged_only and known.get("flag_status") and row.get("flag_status") != OL_FLAG_MARKED:
        return False
    if check_text:
        all_terms = criteria.match_all_terms
        if criteria.subject and not _contains([row.get("subject", "")], criteria.subject, all_terms):
            return False
        if criteria.sender and not _contains(
                [row.get("sender", ""), row.get("sender_email", "")], criteria.sender, all_terms):
            return False
        if criteria.recipient:
            people = [r["name"] for r in row.get("to_recipients", []) + row.get("cc_recipients", [])]
            if not _contains(people, criteria.recipient, all_terms):
                return False
    return True


def _sort_key(row: Dict[str, Any]) -> str:
    received = row.get("received_time", "")
    return received if received != "Unknown" else ""


# --------------------------------------------------------------------------- table path

def _setup_columns(table: Any) -> List[str]:
    table.Columns.RemoveAll()
    keys: List[str] = []
    for key, name in REQUIRED_COLUMNS:
        table.Columns.Add(name)
        keys.append(key)
    for key, name in OPTIONAL_COLUMNS:
        try:
            table.Columns.Add(name)
            keys.append(key)
        except Exception as e:
            logger.debug(f"Table column {name} not available: {e}")
    return keys


def _looks_like_ids(values: Sequence[Any]) -> bool:
    return bool(values) and all(isinstance(v, str) and _HEX.match(v) for v in values)


def _read_chunk(table: Any, n_rows: int, n_cols: int, state: Dict[str, bool]) -> List[Sequence[Any]]:
    """Up to n_rows rows from the current position, in bulk when the store allows it."""
    if state.get("bulk", True) and n_rows > 0:
        try:
            block = table.GetArray(n_rows)
            rows = [tuple(r) for r in (block or ())]
            if rows and len(rows[0]) != n_cols and len(rows) == n_cols:
                rows = [tuple(r) for r in zip(*rows)]  # column-major array
            elif rows and len(rows) == n_cols and len(rows[0]) == n_cols and _looks_like_ids(rows[0]):
                rows = [tuple(r) for r in zip(*rows)]  # square block, first line holds only ids
            if all(len(r) == n_cols for r in rows) and len(rows) <= n_rows:
                return rows
            raise ValueError("unexpected array shape")
        except Exception as e:
            # Only safe to switch while nothing was consumed by the failed call.
            logger.debug(f"Table.GetArray unavailable, reading row by row: {e}")
            state["bulk"] = False
    rows = []
    while len(rows) < n_rows and not table.EndOfTable:
        row = table.GetNextRow()
        if row is None:
            break
        try:
            rows.append(tuple(row.GetValues()))
        except Exception as e:
            logger.debug(f"Skipping unreadable row: {e}")
    return rows


def _scan_table(folder: Any, flt: str, criteria: MailCriteria, check_text: bool, limit: Optional[int],
                deadline: float, max_scan: int, result: QueryResult) -> List[Dict[str, Any]]:
    table = folder.GetTable(flt) if flt else folder.GetTable()
    keys = _setup_columns(table)
    folder_path = folder_display_path(folder)
    store_id = _text(getattr(folder, "StoreID", ""))

    sorted_by_outlook = True
    try:
        table.Sort("[ReceivedTime]", True)
    except Exception as e:
        logger.debug(f"Table.Sort failed, sorting locally: {e}")
        sorted_by_outlook = False

    try:
        remaining = int(table.GetRowCount())
    except Exception:
        remaining = max_scan

    client_filtered = check_text or bool(
        criteria.categories or criteria.exclude_categories or criteria.unread_only or criteria.flagged_only)
    # Stopping early is only exact when Outlook delivers rows newest first.
    stop_at = limit if (limit and sorted_by_outlook) else None
    counted_by_outlook = not client_filtered
    if counted_by_outlook:
        result.total_matched += remaining

    rows: List[Dict[str, Any]] = []
    state: Dict[str, bool] = {}
    scanned = 0
    while remaining > 0:
        if time.monotonic() > deadline or scanned >= max_scan:
            result.complete = False
            break
        chunk = _read_chunk(table, min(_CHUNK, remaining), len(keys), state)
        if not chunk:
            break
        remaining -= len(chunk)
        scanned += len(chunk)
        for values in chunk:
            row = _make_row(dict(zip(keys, values)), folder_path, store_id)
            if row is None or not row_matches(row, criteria, check_text):
                continue
            if not counted_by_outlook:
                result.total_matched += 1
            if stop_at is None or len(rows) < stop_at:
                rows.append(row)
        if stop_at is not None and len(rows) >= stop_at and counted_by_outlook:
            break
    return rows


# --------------------------------------------------------------------------- items path

def _row_from_item(item: Any, folder_path: str, store_id: str) -> Optional[Dict[str, Any]]:
    def get(name: str) -> Any:
        try:
            return getattr(item, name)
        except Exception:
            return None

    attachments = get("Attachments")
    try:
        has_attach = bool(attachments is not None and attachments.Count > 0)
    except Exception:
        has_attach = False
    return _make_row({
        "entry_id": get("EntryID"), "subject": get("Subject"), "sender": get("SenderName"),
        "received": get("ReceivedTime"), "modified": get("LastModificationTime"),
        "categories": get("Categories"), "message_class": get("MessageClass"),
        "to": get("To"), "cc": get("CC"), "sender_address": get("SenderEmailAddress"),
        "unread": get("UnRead"), "flag_status": get("FlagStatus"), "has_attach": has_attach,
        "size": get("Size"), "importance": get("Importance"),
    }, folder_path, store_id)


def _scan_items(folder: Any, flt: str, criteria: MailCriteria, check_text: bool, limit: Optional[int],
                deadline: float, max_scan: int, result: QueryResult) -> List[Dict[str, Any]]:
    items = folder.Items
    if flt:
        items = items.Restrict(flt)
    sorted_by_outlook = True
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception as e:
        logger.debug(f"Items.Sort failed, sorting locally: {e}")
        sorted_by_outlook = False
    folder_path = folder_display_path(folder)
    store_id = _text(getattr(folder, "StoreID", ""))
    stop_at = limit if (limit and sorted_by_outlook) else None

    rows: List[Dict[str, Any]] = []
    scanned = 0
    item = items.GetFirst()
    while item is not None:
        if time.monotonic() > deadline or scanned >= max_scan:
            result.complete = False
            break
        scanned += 1
        try:
            row = _row_from_item(item, folder_path, store_id)
        except Exception as e:
            logger.debug(f"Skipping unreadable item: {e}")
            row = None
        if row is not None and row_matches(row, criteria, check_text):
            result.total_matched += 1
            if stop_at is None or len(rows) < stop_at:
                rows.append(row)
            elif time.monotonic() > deadline - 1:
                result.complete = False
                break
        item = items.GetNext()
    return rows


# --------------------------------------------------------------------------- entry point

def _subfolders(root: Any, max_folders: int, result: QueryResult) -> List[Any]:
    folders, stack = [root], [root]
    while stack:
        current = stack.pop()
        try:
            children = current.Folders
            for i in range(1, children.Count + 1):
                if len(folders) >= max_folders:
                    result.warnings.append(f"Only the first {max_folders} folders were searched")
                    return folders
                child = children.Item(i)
                folders.append(child)
                stack.append(child)
        except Exception as e:
            logger.debug(f"Subfolders unreadable: {e}")
    return folders


def canonical_categories(namespace: Any, names: Sequence[str]) -> Tuple[List[str], List[str]]:
    """Spell category names as the master list does; also return the ones it does not know."""
    try:
        master = namespace.Categories
        known = {}
        for i in range(1, master.Count + 1):
            name = master.Item(i).Name
            known[str(name).lower()] = str(name)
    except Exception as e:
        logger.debug(f"Master category list unreadable: {e}")
        return list(names), []
    canonical = [known.get(n.lower(), n) for n in names]
    unknown = [n for n in names if n.lower() not in known]
    return canonical, unknown


def query_emails(folder: Any, criteria: MailCriteria, max_results: Optional[int] = 1000,
                 include_subfolders: bool = False, time_budget_s: float = 40.0,
                 max_scan: int = 100000, max_folders: int = 300, namespace: Any = None) -> QueryResult:
    """Run the criteria against a folder (and optionally its subfolders), newest first.

    max_results=None returns every match (bounded by max_scan and the time budget;
    ``complete`` tells whether the scan went to the end).
    """
    started = time.monotonic()
    deadline = started + time_budget_s
    result = QueryResult()

    if namespace is not None and (criteria.categories or criteria.exclude_categories):
        criteria.categories, unknown = canonical_categories(namespace, criteria.categories)
        _, unknown_excluded = canonical_categories(namespace, criteria.exclude_categories)
        for name in unknown + unknown_excluded:
            result.warnings.append(f"Category '{name}' is not in the Outlook master category list (check spelling)")

    flt = build_filter(criteria)
    result.filter = flt
    folders = _subfolders(folder, max_folders, result) if include_subfolders else [folder]

    rows: List[Dict[str, Any]] = []
    for current in folders:
        if time.monotonic() > deadline:
            result.complete = False
            break
        path = folder_display_path(current)
        before = result.total_matched
        try:
            found = _scan_table(current, flt, criteria, False, max_results, deadline, max_scan, result)
        except Exception as table_error:
            logger.info(f"Table query failed on '{path}' ({table_error}); using Items.Restrict")
            result.total_matched = before
            try:
                found = _scan_items(current, flt, criteria, False, max_results, deadline, max_scan, result)
                result.method = "items"
            except Exception as items_error:
                if include_subfolders and current is not folder:
                    result.warnings.append(f"Folder '{path}' skipped: {items_error}")
                    continue
                if criteria.body:
                    raise QueryError(f"Outlook rejected the search filter on '{path}': {items_error}")
                # Last resort: dates in the filter, everything else checked on the rows.
                result.total_matched = before
                relaxed = build_filter(criteria, dates_only=True)
                try:
                    found = _scan_table(current, relaxed, criteria, True, max_results, deadline, max_scan, result)
                except Exception as e:
                    raise QueryError(f"Outlook rejected the search filter on '{path}': {e}")
                result.method = "table (filter relaxed, criteria checked locally)"
                result.filter = relaxed
        result.folders.append(path)
        rows.extend(found)

    rows.sort(key=_sort_key, reverse=True)
    if max_results is not None and len(rows) > max_results:
        rows = rows[:max_results]
    for row in rows:
        row.pop("_known", None)
    result.emails = rows
    result.truncated = result.total_matched > len(rows)
    if not result.complete:
        result.warnings.append("Scan stopped on the time or row limit: results may be partial")
    result.elapsed_s = round(time.monotonic() - started, 2)
    return result
