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
        flag_text: Flag label, e.g. 'Follow up vendor' (optional, default 'Follow up')
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


def list_todo_tool(
    days_ahead: Optional[int] = None,
    include_completed: bool = False,
    flagged_folders: Union[str, List[str], None] = None,
) -> Dict[str, Any]:
    """List open Outlook tasks and flagged emails sorted by due date. Use it to build the daily brief.

    Args:
        days_ahead: Only items due within N days (overdue and undated always included)
        include_completed: Include completed items (default False)
        flagged_folders: Extra folder paths to scan for flagged emails (Inbox is always scanned),
                         e.g. "user@company.com/Inbox/Critique"
    """
    try:
        items, diagnostics = ops.list_todo(days_ahead, include_completed, flagged_folders=_as_list(flagged_folders))
        return {"type": "json", "data": {"count": len(items), "items": items, "diagnostics": diagnostics}}
    except Exception as e:
        return _err(e)


def complete_todo_tool(entry_id: str) -> Dict[str, Any]:
    """Mark a task or flagged email as complete, by EntryID returned by list_todo_tool."""
    try:
        return {"type": "text", "text": ops.complete_todo(entry_id)}
    except Exception as e:
        return _err(e)


def mailbox_stats_tool(
    folder_path: Optional[str] = None,
    days: int = 90,
    top_topics: int = 150,
    top_senders: int = 60,
    include_subfolders: bool = False,
) -> Dict[str, Any]:
    """Read-only statistics on a folder over a long period (up to 730 days), without loading the email cache.

    Returns volume per month, To/Cc/other addressing share, automated-sender share, top senders and
    top conversation topics (normalized subject, count, mails where I am in To, last date, main senders).
    Use it to understand the mailbox before designing folders or rules.

    Args:
        folder_path: Folder path (default Inbox), e.g. "user@company.com/Éléments envoyés"
        days: Look-back window in days (default 90, max 730)
        top_topics: Number of conversation topics returned (default 150)
        top_senders: Number of senders returned (default 60)
        include_subfolders: Also scan subfolders (default False)
    """
    try:
        from ..backend.mailbox_stats import mailbox_stats
        return {"type": "json", "data": mailbox_stats(folder_path, days, top_topics, top_senders, include_subfolders)}
    except Exception as e:
        return _err(e)


def move_emails_by_filter_tool(
    target_folder: str,
    source_folder: Optional[str] = None,
    days: Optional[int] = 90,
    subject_contains: Union[str, List[str], None] = None,
    sender_contains: Union[str, List[str], None] = None,
    exclude_categories: Union[str, List[str], None] = None,
    older_than_days: Optional[int] = None,
    dry_run: bool = True,
) -> Dict[str, Any]:
    """Move emails matching criteria from a folder to another, in bulk (server-side filter, cache untouched).

    Criteria are combined with AND; values inside a list are OR. Matching is "contains", case-insensitive.
    Each call is time-boxed (~45s): re-run with the same arguments until "remaining" is 0.
    Default is dry_run=True: returns the match count and a sample without moving anything.

    Args:
        target_folder: Full path of the destination, e.g. "user@company.com/Projects/ProjectA"
        source_folder: Full path of the source (default Inbox)
        days: Only emails received in the last N days (None = no lower bound)
        subject_contains: Subject must contain one of these terms
        sender_contains: Sender name or address must contain one of these terms
        exclude_categories: Skip emails carrying one of these categories, e.g. ["Action", "Attente"]
        older_than_days: Only emails older than N days
        dry_run: True (default) to preview, False to move
    """
    try:
        from ..backend.bulk_move import move_by_filter
        return {"type": "json", "data": move_by_filter(
            target_folder, source_folder, days, _as_list(subject_contains), _as_list(sender_contains),
            _as_list(exclude_categories), older_than_days, dry_run,
        )}
    except Exception as e:
        return _err(e)


def save_attachments_tool(
    email_number: int,
    target_dir: str,
    name_prefix: str = "",
    include_inline: bool = False,
    create_link_shortcuts: bool = True,
    only_names: Union[str, List[str], None] = None,
    dry_run: bool = False,
) -> Dict[str, Any]:
    """Save an email's attachments to a local folder and create .url shortcuts for online file links.

    Real attachments are saved (inline signature images and hidden parts are skipped). Existing files are
    never overwritten: an identical file (same name and size) is skipped, otherwise " (2)" is appended.
    Links to SharePoint / OneDrive / Teams / Google Drive / Box / Dropbox files found in the body become
    Windows .url shortcuts in the same folder. The folder is created if missing.

    Args:
        email_number: Cache number of the email (1-based)
        target_dir: Absolute Windows path of the destination folder
        name_prefix: Optional prefix for saved names, e.g. "2026-10-13 - "
        include_inline: Also save inline images (default False)
        create_link_shortcuts: Create .url shortcuts for online file links (default True)
        only_names: Restrict to these attachment file names
        dry_run: Preview paths without writing (default False)
    """
    try:
        from ..backend.attachments import save_attachments
        return {"type": "json", "data": save_attachments(
            email_number, target_dir, name_prefix, include_inline, create_link_shortcuts,
            _as_list(only_names), dry_run,
        )}
    except Exception as e:
        return _err(e)


def list_local_directory_tool(path: str, depth: int = 2) -> Dict[str, Any]:
    """List the sub-folders of a local directory (with file counts) to decide where to file attachments.

    Args:
        path: Absolute Windows path
        depth: How many levels to list (default 2)
    """
    try:
        from ..backend.attachments import list_directory
        return {"type": "json", "data": list_directory(path, depth)}
    except Exception as e:
        return _err(e)
