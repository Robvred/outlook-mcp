"""Bulk move of emails matching server-side criteria (subject / sender / period).

Filters run through Items.Restrict (DASL) so nothing is loaded in the email cache.
Each call is time-boxed to stay under the MCP device timeout; re-run it until
"remaining" is 0. A dry run returns the count and a sample without moving.
"""

from __future__ import annotations

import time
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

from .audit import log_event
from .com_helpers import split_categories
from .logging_config import get_logger
from .outlook_session.session_manager import OutlookSessionManager

logger = get_logger(__name__)

_SUBJECT = '"urn:schemas:httpmail:subject"'
_FROM_NAME = '"urn:schemas:httpmail:fromname"'
_FROM_EMAIL = '"urn:schemas:httpmail:fromemail"'
_RECEIVED = '"urn:schemas:httpmail:datereceived"'
_KEYWORDS = '"urn:schemas-microsoft-com:office:office#Keywords"'


def _q(term: str) -> str:
    return term.replace("'", "''")


def build_filter(days: Optional[int], subject_contains: List[str], sender_contains: List[str],
                 older_than_days: Optional[int] = None, exclude_categories: Optional[List[str]] = None) -> str:
    parts = []
    now = datetime.now(timezone.utc)
    if days:
        parts.append(f"{_RECEIVED} >= '{(now - timedelta(days=days)).strftime('%Y-%m-%d %H:%M')}'")
    if older_than_days:
        parts.append(f"{_RECEIVED} < '{(now - timedelta(days=older_than_days)).strftime('%Y-%m-%d %H:%M')}'")
    if subject_contains:
        parts.append("(" + " OR ".join(f"{_SUBJECT} LIKE '%{_q(t)}%'" for t in subject_contains) + ")")
    if sender_contains:
        parts.append("(" + " OR ".join(
            f"{_FROM_NAME} LIKE '%{_q(t)}%' OR {_FROM_EMAIL} LIKE '%{_q(t)}%'" for t in sender_contains
        ) + ")")
    if exclude_categories:
        parts.append("NOT (" + " OR ".join(f"{_KEYWORDS} = '{_q(c)}'" for c in exclude_categories) + ")")
    if not parts:
        raise ValueError("At least one criterion is required (days, subject_contains or sender_contains)")
    return "@SQL=" + " AND ".join(parts)


def move_by_filter(target_folder: str, source_folder: Optional[str] = None, days: Optional[int] = 90,
                   subject_contains: Optional[List[str]] = None, sender_contains: Optional[List[str]] = None,
                   exclude_categories: Optional[List[str]] = None, older_than_days: Optional[int] = None,
                   dry_run: bool = False, time_budget_s: float = 45.0, sample_size: int = 10,
                   batch_size: int = 250) -> Dict[str, Any]:
    flt = build_filter(days, subject_contains or [], sender_contains or [], older_than_days, exclude_categories)
    excluded = {c.lower() for c in (exclude_categories or [])}
    start = time.monotonic()
    with OutlookSessionManager() as session:
        src = session.get_folder(source_folder) if source_folder else session.namespace.GetDefaultFolder(6)
        if src is None:
            raise ValueError(f"Source folder '{source_folder}' not found")
        dst = session.get_folder(target_folder)
        if dst is None:
            raise ValueError(f"Target folder '{target_folder}' not found")

        items = src.Items.Restrict(flt)
        total = items.Count
        # Snapshot a batch of EntryIDs first: moving items mutates the restricted collection.
        ids, sample, skipped = [], [], 0
        it = items.GetFirst()
        while it is not None and len(ids) < batch_size:
            try:
                if excluded and any(c.lower() in excluded for c in split_categories(it.Categories)):
                    skipped += 1
                else:
                    ids.append(it.EntryID)
                    if len(sample) < sample_size:
                        sample.append(f"{it.ReceivedTime.strftime('%Y-%m-%d')} | {it.SenderName} | {it.Subject[:90]}")
            except Exception as e:
                logger.debug(f"Skipping item: {e}")
            it = items.GetNext()

        result: Dict[str, Any] = {
            "matched": total, "skipped_by_category_in_batch": skipped, "sample": sample,
            "target": target_folder, "filter": flt,
        }
        if dry_run:
            result["moved"] = 0
            result["remaining"] = total
            return result

        moved, errors = 0, []
        for entry_id in ids:
            if time.monotonic() - start > time_budget_s:
                break
            try:
                session.namespace.GetItemFromID(entry_id).Move(dst)
                moved += 1
            except Exception as e:
                errors.append(str(e)[:120])
        # Remaining = matches still in the source (excluded categories stay there by design).
        result.update({"moved": moved, "remaining": max(total - moved, 0), "errors": errors[:10],
                       })

    log_event("bulk_move", target=target_folder, source=source_folder or "Inbox", moved=moved,
              subject_contains=subject_contains, sender_contains=sender_contains, days=days)
    return result
