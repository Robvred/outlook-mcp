"""
Folder operations for Outlook session management.

This module provides folder-related operations such as creation, deletion, moving, and retrieval.
"""

# Standard library imports
import time
from typing import Any, Dict, List, Optional, Tuple

# Local application imports
from ..audit import log_event
from ..logging_config import get_logger
from ..utils import OutlookFolderType, retry_on_com_error
from .exceptions import FolderNotFoundError, InvalidParameterError, OperationFailedError

logger = get_logger(__name__)


class FolderOperations:
    """Handles all folder-related operations for Outlook."""

    def __init__(self, session_manager):
        """Initialize with a session manager instance."""
        self.session_manager = session_manager
        self._folder_cache = {}
        self._cache_timestamp = 0
        self._cache_ttl = 300  # 5 minutes TTL for folder cache

    def _is_cache_valid(self):
        """Check if folder cache is still valid."""
        return time.time() - self._cache_timestamp < self._cache_ttl

    def _get_cached_folder(self, folder_name: str):
        """Get folder from cache if available and valid."""
        if self._is_cache_valid() and folder_name in self._folder_cache:
            logger.debug(f"Folder cache hit for: {folder_name}")
            return self._folder_cache[folder_name]
        return None

    def _cache_folder(self, folder_name: str, folder):
        """Cache folder for future use."""
        self._folder_cache[folder_name] = folder
        self._cache_timestamp = time.time()
        logger.debug(f"Folder cached: {folder_name}")

    def clear_folder_cache(self):
        """Clear the folder cache."""
        self._folder_cache.clear()
        self._cache_timestamp = 0
        logger.info("Folder cache cleared")

    def get_folder(self, folder_name: Optional[str] = None):
        """Get specified folder or default inbox with caching."""
        # Normalize folder name for caching
        cache_key = folder_name.lower() if folder_name else "inbox"
        
        # Try cache first
        cached_folder = self._get_cached_folder(cache_key)
        if cached_folder:
            return cached_folder
        
        # Get folder and cache it
        folder = self._get_folder_internal(folder_name)
        if folder:
            self._cache_folder(cache_key, folder)
        return folder

    _WELL_KNOWN = {
        "inbox": OutlookFolderType.INBOX,
        "sent items": OutlookFolderType.SENT_MAIL, "sent": OutlookFolderType.SENT_MAIL,
        "deleted items": OutlookFolderType.DELETED_ITEMS, "trash": OutlookFolderType.DELETED_ITEMS,
        "drafts": OutlookFolderType.DRAFTS,
        "outbox": OutlookFolderType.OUTBOX,
        "calendar": OutlookFolderType.CALENDAR,
        "contacts": OutlookFolderType.CONTACTS,
        "tasks": OutlookFolderType.TASKS,
    }

    def _well_known(self, name: str):
        """Default-store folder for an English well-known name (works on any Outlook language)."""
        folder_type = self._WELL_KNOWN.get(name.strip().lower())
        if folder_type is None:
            return None
        try:
            return self.session_manager.outlook_namespace.GetDefaultFolder(folder_type)
        except Exception:
            return None

    def _get_folder_internal(self, folder_name: Optional[str] = None):
        """Internal method to get folder without caching."""
        if not folder_name or not folder_name.strip() or folder_name.strip().lower() == "null":
            return self.session_manager.outlook_namespace.GetDefaultFolder(OutlookFolderType.INBOX)
        return self._well_known(folder_name) or self._get_folder_by_name(folder_name)

    @staticmethod
    def _child(parent, name: str):
        """Subfolder by name: direct access first, then a case-insensitive scan."""
        try:
            return parent.Folders[name]
        except Exception:
            pass
        wanted = name.lower()
        try:
            for subfolder in parent.Folders:
                if str(subfolder.Name).lower() == wanted:
                    return subfolder
        except Exception:
            pass
        return None

    def _stores(self):
        """Store root folders, the default store first."""
        namespace = self.session_manager.outlook_namespace
        roots = []
        try:
            roots.append(namespace.GetDefaultFolder(OutlookFolderType.INBOX).Parent)
        except Exception:
            pass
        default_id = getattr(roots[0], "EntryID", None) if roots else None
        try:
            for root in namespace.Folders:
                if default_id is None or getattr(root, "EntryID", None) != default_id:
                    roots.append(root)
        except Exception:
            pass
        return roots

    def _walk(self, start, parts):
        current = start
        for part in parts:
            current = self._child(current, part)
            if current is None:
                return None
        return current

    def _get_folder_by_name(self, folder_name: str):
        """Resolve a folder path.

        Accepted forms: "user@company.com/Inbox/ProjectA" (store first), "Inbox/ProjectA"
        (relative to a store, the default one first), a single folder name, and Outlook's
        own "\\\\user@company.com\\Inbox\\ProjectA". "%2F" stands for a slash inside a name.
        """
        try:
            parts = [p.strip().replace("%2F", "/").replace("%2f", "/")
                     for p in folder_name.replace("\\", "/").split("/") if p.strip()]
            if not parts:
                raise FolderNotFoundError(f"Folder '{folder_name}' not found")
            stores = self._stores()

            # 1. First part names a store.
            for root in stores:
                if str(root.Name).lower() == parts[0].lower():
                    found = self._walk(root, parts[1:])
                    if found is not None:
                        return found
                    raise FolderNotFoundError(
                        f"Folder '{'/'.join(parts[1:])}' not found in mailbox '{root.Name}'")

            # 2. Path relative to a well-known folder of the default store ("Inbox/ProjectA").
            head = self._well_known(parts[0])
            if head is not None:
                found = self._walk(head, parts[1:])
                if found is not None:
                    return found

            # 3. Path relative to a store root, default store first.
            for root in stores:
                found = self._walk(root, parts)
                if found is not None:
                    return found

            # 4. Single name one level below the top-level folders (e.g. under Inbox).
            if len(parts) == 1:
                for root in stores:
                    try:
                        for top in root.Folders:
                            found = self._child(top, parts[0])
                            if found is not None:
                                return found
                    except Exception:
                        continue
            raise FolderNotFoundError(f"Folder '{folder_name}' not found")
        except FolderNotFoundError:
            raise
        except Exception as e:
            logger.error(f"Error finding folder: {str(e)}")
            raise OperationFailedError(f"Error finding folder '{folder_name}': {str(e)}")

    @retry_on_com_error(max_attempts=3, initial_delay=1.0)
    def create_folder(self, folder_name: str, parent_folder_name: Optional[str] = None) -> str:
        """Create a new folder in the specified parent folder.
        
        Args:
            folder_name: Name of the folder to create
            parent_folder_name: Name of the parent folder (optional, defaults to Inbox)
            
        Returns:
            Success message
        """
        if not folder_name or not isinstance(folder_name, str):
            raise InvalidParameterError("Folder name must be a non-empty string")
            
        try:
            parent_folder = self.get_folder(parent_folder_name)
            
            # Check if folder already exists
            for folder in parent_folder.Folders:
                if folder.Name == folder_name:
                    return f"Folder '{folder_name}' already exists in '{parent_folder.Name}'"
            
            # Create the new folder
            new_folder = parent_folder.Folders.Add(folder_name)
            log_event("folder_created", folder=folder_name, parent=parent_folder.Name)
            logger.info(f"Created folder '{folder_name}' in '{parent_folder.Name}'")
            return f"Folder '{folder_name}' created successfully in '{parent_folder.Name}'"
            
        except Exception as e:
            error_msg = f"Error creating folder '{folder_name}': {str(e)}"
            logger.error(error_msg)
            raise OperationFailedError(error_msg)

    @retry_on_com_error(max_attempts=3, initial_delay=1.0)
    def remove_folder(self, folder_name: str) -> str:
        """Remove an existing folder.
        
        Args:
            folder_name: Name or path of the folder to remove
            
        Returns:
            Success message
        """
        if not folder_name or not isinstance(folder_name, str):
            raise InvalidParameterError("Folder name must be a non-empty string")
            
        try:
            folder = self._get_folder_by_name(folder_name)
            
            # Check if it's a default folder
            if self._is_default_folder(folder):
                raise OperationFailedError(f"Cannot remove default folder '{folder_name}'")
            
            # Get parent folder for the success message
            parent_folder = folder.Parent
            folder_name_only = folder.Name
            
            # Delete the folder
            folder.Delete()
            
            self.clear_folder_cache()
            log_event("folder_removed", folder=folder_name_only, parent=parent_folder.Name)
            logger.info(f"Removed folder '{folder_name_only}' from '{parent_folder.Name}'")
            return f"Folder '{folder_name_only}' removed successfully from '{parent_folder.Name}'"
            
        except Exception as e:
            error_msg = f"Error removing folder '{folder_name}': {str(e)}"
            logger.error(error_msg)
            if isinstance(e, (FolderNotFoundError, OperationFailedError)):
                raise
            raise OperationFailedError(error_msg)

    @retry_on_com_error(max_attempts=3, initial_delay=1.0)
    def move_folder(self, source_folder_path: str, target_parent_path: str) -> str:
        """Move a folder to a different parent folder.
        
        Args:
            source_folder_path: Path to the source folder
            target_parent_path: Path to the target parent folder
            
        Returns:
            Success message
        """
        if not source_folder_path or not isinstance(source_folder_path, str):
            raise InvalidParameterError("Source folder path must be a non-empty string")
        if not target_parent_path or not isinstance(target_parent_path, str):
            raise InvalidParameterError("Target parent path must be a non-empty string")
            
        try:
            # Get source folder
            source_folder = self._get_folder_by_name(source_folder_path)
            
            # Get target parent folder
            target_parent = self._get_folder_by_name(target_parent_path)
            
            # Check if it's a default folder (cannot move default folders)
            if self._is_default_folder(source_folder):
                raise OperationFailedError(f"Cannot move default folder '{source_folder_path}'")
            
            # Move the folder
            source_folder.MoveTo(target_parent)
            
            self.clear_folder_cache()
            log_event("folder_moved", folder=source_folder.Name, target=target_parent.Name)
            logger.info(f"Moved folder '{source_folder.Name}' to '{target_parent.Name}'")
            return f"Folder '{source_folder.Name}' moved successfully to '{target_parent.Name}'"
            
        except Exception as e:
            error_msg = f"Error moving folder '{source_folder_path}' to '{target_parent_path}': {str(e)}"
            logger.error(error_msg)
            if isinstance(e, (FolderNotFoundError, OperationFailedError)):
                raise
            raise OperationFailedError(error_msg)

    def get_folder_list(self):
        """Get list of all folders."""
        try:
            folders = []
            for folder in self.session_manager.outlook_namespace.Folders:
                folders.append(folder)
                # Also add subfolders
                self._add_subfolders(folder, folders)
            return folders
        except Exception as e:
            logger.error(f"Error getting folder list: {str(e)}")
            raise OperationFailedError(f"Error getting folder list: {str(e)}")

    def _add_subfolders(self, folder, folders_list):
        """Recursively add subfolders to the list."""
        try:
            for subfolder in folder.Folders:
                folders_list.append(subfolder)
                self._add_subfolders(subfolder, folders_list)
        except Exception as e:
            logger.warning(f"Error accessing subfolders of {folder.Name}: {str(e)}")

    def _is_default_folder(self, folder) -> bool:
        """True for a store's special folders, in any Outlook language (EntryID comparison)."""
        if isinstance(folder, str):
            path, entry_id, store = folder, None, None
        else:
            path = getattr(folder, "FolderPath", "") or getattr(folder, "Name", "")
            entry_id = getattr(folder, "EntryID", None)
            store = getattr(folder, "Store", None)
        if entry_id:
            owners = [store, self.session_manager.outlook_namespace]
            for folder_type in (3, 4, 5, 6, 9, 10, 11, 12, 13, 16, 23):
                for owner in owners:
                    try:
                        if owner is not None and owner.GetDefaultFolder(folder_type).EntryID == entry_id:
                            return True
                    except Exception:
                        continue
        default_folders = [
            "Inbox", "Sent Items", "Deleted Items", "Drafts",
            "Outbox", "Junk Email", "Archive", "Conversation History"
        ]
        return str(path).replace("/", "\\").split("\\")[-1] in default_folders

    def get_folder_emails(self, folder_name: str = "Inbox", max_emails: int = 100, fast_mode: bool = True, days_filter: int = None) -> Tuple[List[Dict[str, Any]], str]:
        """Load a folder into the numbered listing, newest first.

        Rows come from the query engine: they carry categories, flag status, unread state,
        EntryID and StoreID. ``fast_mode`` is ignored (kept for compatibility).

        Args:
            folder_name: Folder name or path
            max_emails: Maximum number of emails to load
            days_filter: Only the last N days (None = the most recent ``max_emails``)
        """
        from ..mail_query import MailCriteria, query_emails
        from ..shared import replace_email_cache

        if not folder_name or not folder_name.strip():
            return [], f"Error: Invalid folder name: {folder_name}"
        if isinstance(max_emails, bool) or not isinstance(max_emails, int) or max_emails < 1:
            return [], f"Error: Invalid max_emails value: {max_emails}"
        try:
            folder = self.get_folder(folder_name)
            if not folder:
                return [], f"Error: Folder '{folder_name}' not found"
            result = query_emails(folder, MailCriteria(days=days_filter), max_results=max_emails,
                                  namespace=self.session_manager.outlook_namespace)
            scope = f"last {days_filter} days of {folder_name}" if days_filter else f"latest {max_emails} of {folder_name}"
            kept = replace_email_cache(result.emails, source=scope)
            if not kept:
                return [], f"No emails found in '{folder_name}'"
            if result.truncated and days_filter:
                message = (f"Loaded the {kept} most recent of {result.total_matched} emails "
                           f"in '{folder_name}' (limit reached)")
            else:
                message = f"Found {kept} emails in '{folder_name}'"
            if result.warnings:
                message += " [" + "; ".join(result.warnings) + "]"
            return result.emails[:kept], message
        except Exception as e:
            error_msg = f"Error getting emails from folder: {e}"
            logger.error(error_msg)
            return [], f"Error: {error_msg}"


def list_folders():
    """Get list of all folder names.
    
    Returns:
        List of folder names
    """
    from ..outlook_session.session_manager import OutlookSessionManager
    
    try:
        with OutlookSessionManager() as session_manager:
            folder_ops = FolderOperations(session_manager)
            folders = folder_ops.get_folder_list()
            return [folder.Name for folder in folders]
    except Exception as e:
         logger.error(f"Error getting folder list: {str(e)}")
         return []


def create_folder(folder_name: str, parent_folder_name: Optional[str] = None) -> str:
    """Create a new folder in the specified parent folder.
    
    Args:
        folder_name: Name of the folder to create
        parent_folder_name: Name of the parent folder (optional, defaults to Inbox)
        
    Returns:
        Success message
    """
    from ..outlook_session.session_manager import OutlookSessionManager
    
    try:
        with OutlookSessionManager() as session_manager:
            folder_ops = FolderOperations(session_manager)
            return folder_ops.create_folder(folder_name, parent_folder_name)
    except Exception as e:
        logger.error(f"Error creating folder: {str(e)}")
        return f"Error: {str(e)}"


def remove_folder(folder_name: str) -> str:
    """Remove an existing folder.
    
    Args:
        folder_name: Name or path of the folder to remove
        
    Returns:
        Success message
    """
    from ..outlook_session.session_manager import OutlookSessionManager
    
    try:
        with OutlookSessionManager() as session_manager:
            folder_ops = FolderOperations(session_manager)
            return folder_ops.remove_folder(folder_name)
    except Exception as e:
        logger.error(f"Error removing folder: {str(e)}")
        return f"Error: {str(e)}"


def move_folder(source_folder_path: str, target_parent_path: str) -> str:
    """Move a folder to a different parent folder.
    
    Args:
        source_folder_path: Path to the source folder
        target_parent_path: Path to the target parent folder
        
    Returns:
        Success message
    """
    from ..outlook_session.session_manager import OutlookSessionManager
    
    try:
        with OutlookSessionManager() as session_manager:
            folder_ops = FolderOperations(session_manager)
            return folder_ops.move_folder(source_folder_path, target_parent_path)
    except Exception as e:
        logger.error(f"Error moving folder: {str(e)}")
        return f"Error: {str(e)}"


def get_folder_emails(folder_name: str = "Inbox", max_emails: int = 100, days_filter: int = None) -> Tuple[List[Dict[str, Any]], str]:
    """Get emails from a folder with pagination support.
    
    Args:
        folder_name: Name of the folder to get emails from
        max_emails: Maximum number of emails to return
        days_filter: Number of days to filter by (None for number-based loading)
        
    Returns:
        Tuple of (list of email dictionaries, status message)
    """
    from ..outlook_session.session_manager import OutlookSessionManager
    
    try:
        with OutlookSessionManager() as session_manager:
            folder_ops = FolderOperations(session_manager)
            return folder_ops.get_folder_emails(folder_name, max_emails, True, days_filter)
    except Exception as e:
        logger.error(f"Error getting folder emails: {str(e)}")
        return [], f"Error: {str(e)}"