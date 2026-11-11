"""Email search and listing for the Outlook MCP server (backed by backend.mail_query)."""

from .subject_search import search_email_by_subject
from .sender_search import search_email_by_sender
from .recipient_search import search_email_by_recipient
from .body_search import search_email_by_body
from .email_listing import list_recent_emails, get_emails_from_folder
from ..outlook_session.folder_operations import get_folder_emails, list_folders
from ..outlook_session.email_operations import (
    get_email_by_number,
    move_email_to_folder,
    delete_email_by_number,
)
from .sender_search import search_email_by_sender as search_email_by_from
from .recipient_search import search_email_by_recipient as search_email_by_to
from .search_common import get_folder_path_safe, get_date_limit, unified_cache_load_workflow
from .unified_search import unified_search

__all__ = [
    "search_email_by_subject",
    "search_email_by_sender",
    "search_email_by_recipient",
    "search_email_by_body",
    "search_email_by_from",
    "search_email_by_to",
    "list_recent_emails",
    "get_emails_from_folder",
    "get_folder_emails",
    "list_folders",
    "get_email_by_number",
    "move_email_to_folder",
    "delete_email_by_number",
    "get_folder_path_safe",
    "get_date_limit",
    "unified_cache_load_workflow",
    "unified_search",
]
