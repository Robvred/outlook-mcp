"""Triage tools: categories, follow-up flags, tasks and To-Do list. Never send anything."""

from typing import Any, Dict, List, Optional, Union

from ..backend import triage_operations as ops


def _as_list(value: Union[int, str, List[Any], None]) -> List[Any]:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    if isinstance(value, str):
        return [v.strip() for v in value.replace(",", ";").split(";") if v.strip()]
    return [value]


def _err(e: Exception) -> Dict[str, Any]:
    return {"type": "text", "text": f"Error: {e}"}


def ensure_categories_tool(categories: Dict[str, str]) -> Dict[str, Any]:
    """Create Outlook master categories if they do not exist (idempotent).

    Args:
        categories: {name: color}. Colors: red, orange, peach, yellow, green, teal, olive,
                    blue, purple, maroon, steel, darksteel, gray, darkgray, black, none.
                    Example: {"Action": "red", "Attente": "yellow", "Traité": "green", "Info": "gray"}
    """
    try:
        return {"type": "text", "text": "\n".join(ops.ensure_categories(categories))}
    except Exception as e:
        return _err(e)


def set_email_categories_tool(
    email_numbers: Union[int, List[int]],
    categories: Union[str, List[str]],
    mode: str = "add",
    exclusive_group: Union[str, List[str], None] = None,
) -> Dict[str, Any]:
    """Set categories on emails from the current cache.

    Args:
        email_numbers: One number or a list (cache positions, 1-based)
        categories: Category name(s) to apply
        mode: 'add' (default), 'replace' (exactly these) or 'remove'
        exclusive_group: With mode='add', categories removed first so only one status remains.
                         Example: ["Action", "Attente", "Traité", "Info"]
    """
    try:
        res = ops.set_categories(
            [int(n) for n in _as_list(email_numbers)], _as_list(categories), mode, _as_list(exclusive_group)
        )
        return {"type": "text", "text": "\n".join(res)}
    except Exception as e:
        return _err(e)


def flag_email_tool(
    email_numbers: Union[int, List[int]],
    action: str = "flag",
    due_date: Optional[str] = None,
    reminder: Optional[str] = None,
    flag_text: Optional[str] = None,
) -> Dict[str, Any]:
    """Follow-up flag on emails from the current cache. Flagged emails show in Outlook To-Do and Microsoft To Do.

    Args:
        email_numbers: One number or a list (cache positions)
        action: 'flag' (default), 'complete' or 'clear'
        due_date: 'YYYY-MM-DD' (optional)
        reminder: 'YYYY-MM-DD HH:MM' local time (optional, triggers an Outlook reminder popup)
        flag_text: Flag label, e.g. 'Relancer Infosys' (optional, default 'Follow up')
    """
    try:
        res = ops.flag_emails([int(n) for n in _as_list(email_numbers)], action, due_date, reminder, flag_text)
        return {"type": "text", "text": "\n".join(res)}
    except Exception as e:
        return _err(e)


def create_task_tool(
    subject: str,
    body: str = "",
    due_date: Optional[str] = None,
    reminder: Optional[str] = None,
    categories: Union[str, List[str], None] = None,
    importance: int = 1,
    link_email_number: Optional[int] = None,
) -> Dict[str, Any]:
    """Create an Outlook task (with optional reminder). Does not assign or send anything.

    Args:
        subject: Task title
        body: Details: context, expected action, source thread
        due_date: 'YYYY-MM-DD'
        reminder: 'YYYY-MM-DD HH:MM' local time
        categories: Category name(s)
        importance: 0=Low, 1=Normal, 2=High
        link_email_number: Cache number of the source email to embed in the task (optional)
    """
    try:
        return {"type": "text", "text": ops.create_task(
            subject, body, due_date, reminder, _as_list(categories), importance, link_email_number
        )}
    except Exception as e:
        return _err(e)


def list_todo_tool(days_ahead: Optional[int] = None, include_completed: bool = False) -> Dict[str, Any]:
    """List open Outlook To-Do items (tasks + flagged emails) sorted by due date. Use it to build the daily brief.

    Args:
        days_ahead: Only items due within N days (overdue and undated always included)
        include_completed: Include completed items (default False)
    """
    try:
        items = ops.list_todo(days_ahead, include_completed)
        return {"type": "json", "data": {"count": len(items), "items": items}}
    except Exception as e:
        return _err(e)


def complete_todo_tool(entry_id: str) -> Dict[str, Any]:
    """Mark a task or flagged email as complete, by EntryID returned by list_todo_tool."""
    try:
        return {"type": "text", "text": ops.complete_todo(entry_id)}
    except Exception as e:
        return _err(e)
