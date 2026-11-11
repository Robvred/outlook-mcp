"""Single-email operations on the numbered listing: lookup, move, delete (to Deleted Items)."""

from typing import Any, Dict

from ..audit import log_event
from ..logging_config import get_logger
from ..outlook_session.session_manager import OutlookSessionManager
from ..shared import update_cache_entry

logger = get_logger(__name__)

OL_FOLDER_DELETED_ITEMS = 3


def _folder_path(folder: Any) -> str:
    try:
        return str(folder.FolderPath).lstrip("\\").replace("\\", "/")
    except Exception:
        return str(getattr(folder, "Name", "?"))


def _same_folder(a: Any, b: Any) -> bool:
    try:
        return bool(a.EntryID) and a.EntryID == b.EntryID
    except Exception:
        return False


class EmailOperations:
    """Email operations bound to an open Outlook session."""

    def __init__(self, session_manager):
        self.session_manager = session_manager

    def get_email_by_number(self, email_number: int) -> Dict[str, Any]:
        """Cached row of an email by its number in the listing (1-based)."""
        from ..com_helpers import cached_entry

        return cached_entry(email_number)

    def _move(self, email_number: int, target: Any, event: str) -> str:
        from ..com_helpers import cached_entry, get_item_by_number

        entry = cached_entry(email_number)
        item = get_item_by_number(self.session_manager, email_number)
        subject = str(getattr(item, "Subject", "") or "")[:80]
        target_path = _folder_path(target)
        try:
            if _same_folder(item.Parent, target):
                return f"Email #{email_number} '{subject}' is already in '{target_path}'"
        except Exception:
            pass
        moved = item.Move(target)
        # A move changes the EntryID: keep the same number pointing at the moved message.
        new_id = getattr(moved, "EntryID", None) if moved is not None else None
        update_cache_entry(
            entry["entry_id"], new_id if isinstance(new_id, str) and new_id else None,
            folder=target_path, store_id=getattr(target, "StoreID", None) or entry.get("store_id"),
        )
        log_event(event, email_number=email_number, target=target_path, entry_id=new_id or entry["entry_id"])
        logger.info(f"Moved email #{email_number} to '{target_path}'")
        return f"Email #{email_number} '{subject}' moved to '{target_path}'. Listing numbers are unchanged."

    def move_email_to_folder(self, email_number: int, target_folder_name: str) -> str:
        """Move one email of the listing to a folder. Returns a message (starting with 'Error' on failure)."""
        if not target_folder_name or not str(target_folder_name).strip():
            return f"Error: Invalid folder name: {target_folder_name}"
        try:
            target = self.session_manager.get_folder(target_folder_name)
            if not target:
                return f"Error: Target folder '{target_folder_name}' not found"
            return self._move(email_number, target, "email_moved")
        except Exception as e:
            logger.error(f"Error moving email: {e}")
            return f"Error: {e}"

    def delete_email_by_number(self, email_number: int) -> str:
        """Move one email to the Deleted Items folder of its own mailbox (recoverable)."""
        from ..com_helpers import get_item_by_number

        try:
            item = get_item_by_number(self.session_manager, email_number)
            try:
                target = item.Parent.Store.GetDefaultFolder(OL_FOLDER_DELETED_ITEMS)
            except Exception:
                target = self.session_manager.namespace.GetDefaultFolder(OL_FOLDER_DELETED_ITEMS)
            return self._move(email_number, target, "email_deleted")
        except Exception as e:
            logger.error(f"Error deleting email: {e}")
            return f"Error: {e}"


def get_email_by_number(email_number: int) -> Dict[str, Any]:
    """Cached row of an email by its number in the listing."""
    from ..com_helpers import cached_entry

    return cached_entry(email_number)


def move_email_to_folder(email_number: int, target_folder_name: str) -> str:
    """Move an email to a different folder."""
    with OutlookSessionManager() as session_manager:
        return EmailOperations(session_manager).move_email_to_folder(email_number, target_folder_name)


def delete_email_by_number(email_number: int) -> str:
    """Delete an email by moving it to the Deleted Items folder."""
    with OutlookSessionManager() as session_manager:
        return EmailOperations(session_manager).delete_email_by_number(email_number)
