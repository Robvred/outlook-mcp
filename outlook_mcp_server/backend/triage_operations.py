"""Triage operations: categories, follow-up flags, tasks and To-Do listing.

These operations only change item metadata (categories, flags) or create tasks.
They never send anything. Every change is written to the local audit log.
"""

from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

from .audit import log_event
from .com_helpers import (
    fmt_dt,
    get_item_by_number,
    join_categories,
    parse_dt,
    split_categories,
)
from .logging_config import get_logger
from .outlook_session.session_manager import OutlookSessionManager

logger = get_logger(__name__)

# OlCategoryColor
CATEGORY_COLORS = {
    "none": 0, "red": 1, "orange": 2, "peach": 3, "yellow": 4, "green": 5,
    "teal": 6, "olive": 7, "blue": 8, "purple": 9, "maroon": 10, "steel": 11,
    "darksteel": 12, "gray": 13, "darkgray": 14, "black": 15,
}

# OlMarkInterval
OL_MARK_NO_DATE = 4
OL_MARK_COMPLETE = 5

OL_FOLDER_TASKS = 13
OL_TASK_ITEM = 3
OL_CLASS_MAIL = 43
OL_CLASS_TASK = 48
OL_TASK_COMPLETE = 2


def ensure_categories(categories: Dict[str, str]) -> List[str]:
    """Create master categories if missing. categories = {name: color_name}."""
    results = []
    with OutlookSessionManager() as session:
        master = session.namespace.Categories
        existing = {master.Item(i).Name.lower(): master.Item(i) for i in range(1, master.Count + 1)}
        for name, color in categories.items():
            color_id = CATEGORY_COLORS.get((color or "none").lower().replace(" ", ""), 0)
            if name.lower() in existing:
                cat = existing[name.lower()]
                if color and cat.Color != color_id:
                    cat.Color = color_id
                    results.append(f"{name}: exists, color updated to {color}")
                else:
                    results.append(f"{name}: already exists")
            else:
                master.Add(name, color_id)
                results.append(f"{name}: created ({color})")
    log_event("categories_ensured", categories=categories)
    return results


def set_categories(email_numbers: List[int], categories: List[str], mode: str = "replace",
                   exclusive_group: Optional[List[str]] = None) -> List[str]:
    """Apply categories to cached emails.

    mode: 'replace' (set exactly these), 'add', or 'remove'.
    exclusive_group: when adding, first remove any of these categories (e.g. the 4 triage statuses).
    """
    if mode not in ("replace", "add", "remove"):
        raise ValueError("mode must be 'replace', 'add' or 'remove'")
    results = []
    with OutlookSessionManager() as session:
        for n in email_numbers:
            try:
                item = get_item_by_number(session, n)
                current = split_categories(item.Categories)
                if mode == "replace":
                    new = list(categories)
                elif mode == "add":
                    group = {c.lower() for c in (exclusive_group or [])}
                    new = [c for c in current if c.lower() not in group] + list(categories)
                else:
                    drop = {c.lower() for c in categories}
                    new = [c for c in current if c.lower() not in drop]
                item.Categories = join_categories(new)
                item.Save()
                results.append(f"#{n} '{item.Subject[:60]}': {', '.join(dict.fromkeys(new)) or '(none)'}")
            except Exception as e:
                results.append(f"#{n}: ERROR {e}")
    log_event("categories_set", emails=email_numbers, categories=categories, mode=mode)
    return results


def flag_emails(email_numbers: List[int], action: str = "flag", due_date: Optional[str] = None,
                reminder: Optional[str] = None, flag_text: Optional[str] = None) -> List[str]:
    """Follow-up flag on cached emails. action: 'flag', 'complete' or 'clear'.

    Flagged emails appear in Outlook To-Do / Tasks and in Microsoft To Do.
    """
    if action not in ("flag", "complete", "clear"):
        raise ValueError("action must be 'flag', 'complete' or 'clear'")
    due = parse_dt(due_date)
    rem = parse_dt(reminder)
    results = []
    with OutlookSessionManager() as session:
        for n in email_numbers:
            try:
                item = get_item_by_number(session, n)
                if item.Class != OL_CLASS_MAIL:
                    results.append(f"#{n} '{item.Subject[:60]}': skipped (not a mail item, e.g. meeting invitation)")
                    continue
                if action == "clear":
                    item.ClearTaskFlag()
                elif action == "complete":
                    item.MarkAsTask(OL_MARK_COMPLETE)
                else:
                    item.MarkAsTask(OL_MARK_NO_DATE)
                    if flag_text:
                        item.FlagRequest = flag_text
                    if due:
                        start = min(datetime.now(), due)
                        item.TaskStartDate = start.replace(hour=0, minute=0, second=0, microsecond=0)
                        item.TaskDueDate = due.replace(hour=0, minute=0, second=0, microsecond=0)
                    if rem:
                        item.ReminderSet = True
                        item.ReminderTime = rem
                item.Save()
                detail = f" due {due_date}" if due_date and action == "flag" else ""
                detail += f", reminder {reminder}" if reminder and action == "flag" else ""
                results.append(f"#{n} '{item.Subject[:60]}': {action}{detail}")
            except Exception as e:
                results.append(f"#{n}: ERROR {e}")
    log_event("emails_flagged", emails=email_numbers, action=action, due=due_date, reminder=reminder)
    return results


def create_task(subject: str, body: str = "", due_date: Optional[str] = None,
                reminder: Optional[str] = None, categories: Optional[List[str]] = None,
                importance: int = 1, link_email_number: Optional[int] = None) -> str:
    """Create an Outlook task. Optionally attach the source email as a link (.msg item)."""
    due = parse_dt(due_date)
    rem = parse_dt(reminder)
    with OutlookSessionManager() as session:
        task = session.outlook.CreateItem(OL_TASK_ITEM)
        task.Subject = subject
        task.Body = body or ""
        task.Importance = int(importance)
        if due:
            task.StartDate = min(datetime.now(), due).replace(hour=0, minute=0, second=0, microsecond=0)
            task.DueDate = due.replace(hour=0, minute=0, second=0, microsecond=0)
        if rem:
            task.ReminderSet = True
            task.ReminderTime = rem
        if categories:
            task.Categories = join_categories(categories)
        if link_email_number:
            try:
                src = get_item_by_number(session, link_email_number)
                task.Attachments.Add(src, 1)  # olByValue: embeds the mail in the task
            except Exception as e:
                logger.warning(f"Could not attach email #{link_email_number}: {e}")
        task.Save()
        entry_id = task.EntryID
    log_event("task_created", subject=subject, due=due_date, reminder=reminder, entry_id=entry_id)
    return f"Task created: '{subject}' due {due_date or '-'} reminder {reminder or '-'}. EntryID={entry_id}"


def _iter_restricted(folder: Any, flt: str, max_items: int):
    """Iterate a folder's items through Restrict + GetFirst/GetNext (fast over COM)."""
    items = folder.Items.Restrict(flt)
    it = items.GetFirst()
    n = 0
    while it is not None and n < max_items:
        yield it
        n += 1
        it = items.GetNext()


def list_todo(days_ahead: Optional[int] = None, include_completed: bool = False,
              max_items: int = 200, flagged_folders: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    """List open tasks (Tasks folder) and flagged emails (Inbox + optional folders), sorted by due date.

    days_ahead: only items due within N days (overdue and undated items always included).
    flagged_folders: extra folder paths to scan for flagged emails (default: Inbox only).
    """
    limit = datetime.now() + timedelta(days=days_ahead) if days_ahead is not None else None
    out: List[Dict[str, Any]] = []

    def keep(due_raw: Any) -> Optional[str]:
        due_str = fmt_dt(due_raw)
        if limit and due_str and due_raw.replace(tzinfo=None) > limit:
            return "__skip__"
        return due_str

    with OutlookSessionManager() as session:
        # Tasks
        tasks = session.namespace.GetDefaultFolder(OL_FOLDER_TASKS)
        task_filter = "[Complete] = True OR [Complete] = False" if include_completed else "[Complete] = False"
        for it in _iter_restricted(tasks, task_filter, max_items):
            try:
                due_str = keep(it.DueDate)
                if due_str == "__skip__":
                    continue
                out.append({
                    "type": "task",
                    "subject": it.Subject,
                    "due": due_str,
                    "reminder": fmt_dt(it.ReminderTime) if it.ReminderSet else None,
                    "categories": split_categories(it.Categories),
                    "importance": {0: "Low", 1: "Normal", 2: "High"}.get(it.Importance, "Normal"),
                    "completed": bool(it.Complete),
                    "entry_id": it.EntryID,
                })
            except Exception as e:
                logger.debug(f"Skipping task: {e}")

        # Flagged emails (FlagStatus: 2 = flagged, 1 = complete)
        folders = [session.namespace.GetDefaultFolder(6)]  # olFolderInbox
        for path in flagged_folders or []:
            try:
                f = session.get_folder(path)
                if f:
                    folders.append(f)
            except Exception as e:
                logger.debug(f"Folder {path} not found: {e}")
        mail_filter = "[FlagStatus] = 2 OR [FlagStatus] = 1" if include_completed else "[FlagStatus] = 2"
        for folder in folders:
            for it in _iter_restricted(folder, mail_filter, max_items):
                try:
                    if it.Class != OL_CLASS_MAIL:
                        continue
                    due_str = keep(it.TaskDueDate)
                    if due_str == "__skip__":
                        continue
                    out.append({
                        "type": "email",
                        "subject": it.Subject,
                        "from": getattr(it, "SenderName", ""),
                        "flag": getattr(it, "FlagRequest", ""),
                        "due": due_str,
                        "reminder": fmt_dt(it.ReminderTime) if it.ReminderSet else None,
                        "categories": split_categories(it.Categories),
                        "importance": {0: "Low", 1: "Normal", 2: "High"}.get(it.Importance, "Normal"),
                        "completed": it.FlagStatus == 1,
                        "entry_id": it.EntryID,
                    })
                except Exception as e:
                    logger.debug(f"Skipping flagged mail: {e}")

    out.sort(key=lambda e: (e["due"] is None, e["due"] or ""))
    return out


def complete_todo(entry_id: str) -> str:
    """Mark a task complete, or a flagged email's flag complete, by EntryID (from list_todo)."""
    with OutlookSessionManager() as session:
        it = session.namespace.GetItemFromID(entry_id)
        if it.Class == OL_CLASS_TASK:
            it.MarkComplete()
        elif it.Class == OL_CLASS_MAIL:
            it.MarkAsTask(OL_MARK_COMPLETE)
        else:
            raise ValueError("Item is neither a task nor an email")
        it.Save()
        subject = it.Subject
    log_event("todo_completed", entry_id=entry_id)
    return f"Completed: '{subject}'"
