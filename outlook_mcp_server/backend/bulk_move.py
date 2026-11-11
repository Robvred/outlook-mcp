"""Bulk move of the emails matching criteria (period, subject, sender, categories).

Selection goes through the query engine (mail_query): the filter is evaluated by
Outlook on the live folder, category criteria are re-checked on every row, and the
email cache is left untouched. The whole selection is read before the first move,
so counts are exact and a move never disturbs the iteration.

Each call is time-boxed to stay under the MCP client timeout: run it again with the
same arguments until "remaining" is 0. A dry run returns the count and a sample.
"""

from __future__ import annotations

import time
from typing import Any, Dict, List, Optional

from .audit import log_event
from .com_helpers import get_item
from .logging_config import get_logger
from .mail_query import MailCriteria, build_filter, folder_display_path, query_emails
from .outlook_session.session_manager import OutlookSessionManager

logger = get_logger(__name__)

OL_FOLDER_INBOX = 6
_BUSY = (-2147418111, -2147417846)  # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER


def _criteria(days: Optional[int], subject_contains: Optional[List[str]], sender_contains: Optional[List[str]],
              older_than_days: Optional[int], exclude_categories: Optional[List[str]],
              categories: Optional[List[str]], categories_match: str) -> MailCriteria:
    if categories_match not in ("any", "all"):
        raise ValueError("categories_match must be 'any' or 'all'")
    criteria = MailCriteria(
        days=days, older_than_days=older_than_days, subject=subject_contains or [],
        sender=sender_contains or [], categories=categories or [],
        categories_match_all=categories_match == "all", exclude_categories=exclude_categories or [],
    )
    if criteria.is_empty():
        raise ValueError(
            "At least one criterion is required (days, older_than_days, subject_contains, "
            "sender_contains or categories)")
    clash = {c.lower() for c in criteria.categories} & {c.lower() for c in criteria.exclude_categories}
    if clash:
        raise ValueError(f"Category both required and excluded: {', '.join(sorted(clash))}")
    return criteria


def build_filter_for(days: Optional[int], subject_contains: List[str], sender_contains: List[str],
                     older_than_days: Optional[int] = None, exclude_categories: Optional[List[str]] = None,
                     categories: Optional[List[str]] = None, categories_match: str = "any") -> str:
    """DASL filter sent to Outlook for these criteria (exclusions are applied on the rows)."""
    return build_filter(_criteria(days, subject_contains, sender_contains, older_than_days,
                                  exclude_categories, categories, categories_match))


def _move_one(session: Any, row: Dict[str, Any], dst: Any) -> None:
    """Move one message, retrying briefly while Outlook reports it is busy."""
    for attempt in range(3):
        try:
            get_item(session, row["entry_id"], row.get("store_id")).Move(dst)
            return
        except Exception as e:
            if getattr(e, "hresult", None) in _BUSY and attempt < 2:
                time.sleep(0.5 * (attempt + 1))
                continue
            raise


def move_by_filter(target_folder: str, source_folder: Optional[str] = None, days: Optional[int] = 90,
                   subject_contains: Optional[List[str]] = None, sender_contains: Optional[List[str]] = None,
                   exclude_categories: Optional[List[str]] = None, older_than_days: Optional[int] = None,
                   dry_run: bool = False, time_budget_s: float = 45.0, sample_size: int = 10,
                   categories: Optional[List[str]] = None, categories_match: str = "any") -> Dict[str, Any]:
    """Move the matching emails of source_folder to target_folder.

    categories: keep only the emails carrying one of these categories ('any') or all of
    them ('all'). exclude_categories: skip the emails carrying one of these.
    """
    criteria = _criteria(days, subject_contains, sender_contains, older_than_days,
                         exclude_categories, categories, categories_match)
    start = time.monotonic()
    with OutlookSessionManager() as session:
        src = session.get_folder(source_folder) if source_folder else session.namespace.GetDefaultFolder(OL_FOLDER_INBOX)
        if src is None:
            raise ValueError(f"Source folder '{source_folder}' not found")
        dst = session.get_folder(target_folder)
        if dst is None:
            raise ValueError(f"Target folder '{target_folder}' not found")
        try:
            same = bool(src.EntryID) and src.EntryID == dst.EntryID
        except Exception:
            same = False
        if same:
            raise ValueError("Source and target are the same folder")

        selection = query_emails(src, criteria, max_results=None, time_budget_s=max(5.0, min(20.0, time_budget_s / 2)),
                                 namespace=session.namespace)
        rows = selection.emails
        result: Dict[str, Any] = {
            "matched": len(rows),
            "sample": [f"{r['received_time'][:10]} | {r['sender']} | {r['subject'][:90]}"
                       + (f" | [{r['categories']}]" if r["categories"] else "") for r in rows[:sample_size]],
            "source": folder_display_path(src),
            "target": folder_display_path(dst),
            "criteria": criteria.describe(),
            "filter": selection.filter,
            "selection_complete": selection.complete,
            "dry_run": dry_run,
        }
        if selection.warnings:
            result["warnings"] = selection.warnings
        if not selection.complete:
            result["note"] = ("The selection scan was cut short: more emails may match. "
                              "Run again until matched is 0.")
        if dry_run:
            result.update({"moved": 0, "remaining": len(rows)})
            return result

        moved, errors = 0, []
        for row in rows:
            if time.monotonic() - start > time_budget_s:
                break
            try:
                _move_one(session, row, dst)
                moved += 1
            except Exception as e:
                errors.append(f"{row['subject'][:60]}: {str(e)[:120]}")
        remaining = len(rows) - moved
        if not selection.complete:
            remaining = max(remaining, 1)  # unknown: never report 0 on a partial scan
        result.update({"moved": moved, "failed": len(errors), "errors": errors[:10], "remaining": remaining})
        if errors and remaining == len(errors):
            result["note"] = "Only failed emails remain: fix the reported errors before running again."

    log_event("bulk_move", target=result["target"], source=result["source"], moved=moved,
              failed=len(errors), criteria=result["criteria"])
    return result
