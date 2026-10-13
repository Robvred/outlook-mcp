"""Read-only calendar listing: meetings and appointments in a date window.

Recurring series are expanded into their occurrences (Items.IncludeRecurrences).
Nothing is created, accepted, declined or sent.
"""

from __future__ import annotations

import re
import time
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

from .com_helpers import fmt_dt, parse_dt, recipient_smtp, split_categories
from .logging_config import get_logger
from .outlook_session.session_manager import OutlookSessionManager

logger = get_logger(__name__)

OL_FOLDER_CALENDAR = 9

# OlMeetingRecipientType
OL_REQUIRED = 1
OL_OPTIONAL = 2
OL_RESOURCE = 3

# OlMeetingStatus
OL_MEETING_CANCELED = 5
OL_MEETING_RECEIVED_AND_CANCELED = 7

RESPONSE_STATUS = {0: "none", 1: "organizer", 2: "tentative", 3: "accepted", 4: "declined", 5: "not_responded"}
BUSY_STATUS = {0: "free", 1: "tentative", 2: "busy", 3: "out_of_office", 4: "working_elsewhere"}

# Lines that start the online-meeting join block appended by Teams / Zoom / Webex.
_JOIN_BLOCK = re.compile(r"^\s*(_{20,}|Microsoft Teams|Réunion Microsoft Teams|Join Zoom Meeting|Join Webex)",
                         re.IGNORECASE | re.MULTILINE)
_ONLINE_HINT = re.compile(r"teams\.microsoft\.com|zoom\.us/j|webex\.com", re.IGNORECASE)

# Jet date formats tried for Restrict; which one Outlook accepts depends on the Windows locale.
_JET_FORMATS = ("%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M")


def _naive(value: Any) -> Optional[datetime]:
    try:
        return datetime(value.year, value.month, value.day, value.hour, value.minute)
    except Exception:
        return None


def _clean_body(body: str, max_chars: int) -> Dict[str, Any]:
    body = (body or "").replace("\r\n", "\n")
    online = bool(_ONLINE_HINT.search(body))
    m = _JOIN_BLOCK.search(body)
    agenda = body[: m.start()] if m else body
    agenda = re.sub(r"\n{3,}", "\n\n", agenda).strip()
    truncated = len(agenda) > max_chars
    return {"agenda": agenda[:max_chars], "agenda_truncated": truncated, "online_link_in_body": online}


def _attendees(item: Any) -> Dict[str, List[str]]:
    out: Dict[str, List[str]] = {"required": [], "optional": [], "resources": []}
    try:
        recips = item.Recipients
        for i in range(1, recips.Count + 1):
            r = recips.Item(i)
            smtp = recipient_smtp(r)
            label = f"{r.Name} <{smtp}>" if smtp and smtp.lower() not in r.Name.lower() else r.Name
            key = {OL_REQUIRED: "required", OL_OPTIONAL: "optional", OL_RESOURCE: "resources"}.get(r.Type)
            if key:
                out[key].append(label)
    except Exception as e:
        logger.debug(f"Attendees unreadable: {e}")
    return out


def _restricted_occurrences(items: Any, start: datetime, end: datetime):
    """Try Restrict with each Jet date format; return the collection that yields an in-window item."""
    for fmt in _JET_FORMATS:
        flt = f"[Start] < '{end.strftime(fmt)}' AND [End] > '{start.strftime(fmt)}'"
        try:
            res = items.Restrict(flt)
            first = res.GetFirst()
            if first is None:
                continue
            s, e = _naive(first.Start), _naive(first.End)
            if s and e and s < end and e > start:
                return res, fmt
        except Exception as ex:
            logger.debug(f"Restrict failed with format {fmt}: {ex}")
    return None, None


def list_calendar_events(start_date: Optional[str] = None, end_date: Optional[str] = None,
                         days: int = 1, calendar_path: Optional[str] = None,
                         include_body: bool = True, max_body_chars: int = 1500,
                         include_canceled: bool = False, include_declined: bool = False,
                         max_events: int = 100, time_budget_s: float = 45.0) -> Dict[str, Any]:
    """List calendar occurrences overlapping [start, end) in local time, sorted by start."""
    start = parse_dt(start_date) if start_date else datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    if end_date:
        end = parse_dt(end_date)
        if end_date.strip().count(":") == 0:  # date only: include the whole day
            end = end + timedelta(days=1)
    else:
        end = start + timedelta(days=max(1, int(days)))
    if end <= start:
        raise ValueError("end_date must be after start_date")
    if (end - start).days > 31:
        raise ValueError("Window limited to 31 days")

    t0 = time.monotonic()
    diagnostics: Dict[str, Any] = {"method": None, "scanned": 0, "timed_out": False}
    events: List[Dict[str, Any]] = []

    with OutlookSessionManager() as session:
        folder = session.get_folder(calendar_path) if calendar_path else \
            session.namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)
        if folder is None:
            raise ValueError(f"Calendar folder '{calendar_path}' not found")

        items = folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        coll, fmt = _restricted_occurrences(items, start, end)
        if coll is not None:
            diagnostics["method"] = f"restrict ({fmt})"
        else:
            # Locale-independent fallback: walk the sorted collection and stop past the window.
            diagnostics["method"] = "scan"
            coll = items

        it = coll.GetFirst()
        while it is not None:
            if time.monotonic() - t0 > time_budget_s:
                diagnostics["timed_out"] = True
                break
            diagnostics["scanned"] += 1
            try:
                s, e = _naive(it.Start), _naive(it.End)
                if s is None or e is None:
                    it = coll.GetNext()
                    continue
                if s >= end:
                    break
                if e <= start or (e == s and s < start):
                    it = coll.GetNext()
                    continue

                status = getattr(it, "MeetingStatus", 0)
                canceled = status in (OL_MEETING_CANCELED, OL_MEETING_RECEIVED_AND_CANCELED) \
                    or str(it.Subject or "").lower().startswith(("canceled:", "annulé", "annulée"))
                response = RESPONSE_STATUS.get(getattr(it, "ResponseStatus", 0), "unknown")
                if (canceled and not include_canceled) or (response == "declined" and not include_declined):
                    it = coll.GetNext()
                    continue

                ev: Dict[str, Any] = {
                    "subject": it.Subject,
                    "start": fmt_dt(it.Start),
                    "end": fmt_dt(it.End),
                    "all_day": bool(getattr(it, "AllDayEvent", False)),
                    "location": getattr(it, "Location", "") or "",
                    "organizer": getattr(it, "Organizer", "") or "",
                    "my_response": response,
                    "busy_status": BUSY_STATUS.get(getattr(it, "BusyStatus", 2), "unknown"),
                    "is_meeting": status != 0,
                    "is_recurring": bool(getattr(it, "IsRecurring", False)),
                    "canceled": canceled,
                    "categories": split_categories(getattr(it, "Categories", "")),
                    "attendees": _attendees(it),
                    "attachments_count": getattr(getattr(it, "Attachments", None), "Count", 0),
                    "conversation_topic": getattr(it, "ConversationTopic", "") or "",
                    "entry_id": it.EntryID,
                }
                if include_body:
                    ev.update(_clean_body(getattr(it, "Body", ""), max_body_chars))
                events.append(ev)
                if len(events) >= max_events:
                    break
            except Exception as ex:
                logger.debug(f"Skipping calendar item: {ex}")
            it = coll.GetNext()

    events.sort(key=lambda x: (x["start"] or ""))
    diagnostics["elapsed_s"] = round(time.monotonic() - t0, 2)
    return {
        "window": {"start": start.strftime("%Y-%m-%d %H:%M"), "end": end.strftime("%Y-%m-%d %H:%M")},
        "count": len(events),
        "events": events,
        "diagnostics": diagnostics,
    }
