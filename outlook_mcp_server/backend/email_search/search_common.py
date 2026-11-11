"""Shared helpers for the search and listing functions (all backed by mail_query)."""

from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional, Tuple

from ..logging_config import get_logger
from ..mail_query import MailCriteria, QueryResult, query_emails
from ..outlook_session.session_manager import OutlookSessionManager
from ..shared import MAX_CACHE_SIZE, replace_email_cache

logger = get_logger(__name__)


def get_folder_path_safe(folder_name: Optional[str] = None) -> str:
    """Folder path, defaulting to Inbox."""
    return folder_name if folder_name else "Inbox"


def get_date_limit(days: int) -> datetime:
    return datetime.now(timezone.utc) - timedelta(days=days)


def unified_cache_load_workflow(emails_data: List[Dict[str, Any]], operation_name: str = "cache_operation") -> bool:
    """Replace the numbered listing with these emails (kept in the given order)."""
    try:
        return replace_email_cache(emails_data, source=operation_name) > 0
    except Exception as e:
        logger.error(f"Failed to load the cache for {operation_name}: {e}")
        return False


def query_to_cache(folder_name: Optional[str], criteria: MailCriteria, source: str,
                   max_results: Optional[int] = None, include_subfolders: bool = False,
                   time_budget_s: float = 40.0) -> QueryResult:
    """Run a live query and make its result the numbered listing (also when it is empty)."""
    limit = min(max_results or MAX_CACHE_SIZE, MAX_CACHE_SIZE)
    with OutlookSessionManager() as session:
        folder = session.get_folder(folder_name)
        if not folder:
            raise ValueError(f"Folder '{get_folder_path_safe(folder_name)}' not found")
        result = query_emails(folder, criteria, max_results=limit, include_subfolders=include_subfolders,
                              time_budget_s=time_budget_s, namespace=session.namespace)
    replace_email_cache(result.emails, source=source)
    return result


def describe_result(result: QueryResult, folder_label: str, what: str = "emails") -> str:
    """One-line status for a load or search."""
    count = len(result.emails)
    if result.truncated:
        text = f"Loaded the {count} most recent of {result.total_matched} {what} in '{folder_label}'"
    else:
        text = f"Found {count} {what} in '{folder_label}'"
    if result.warnings:
        text += " [" + "; ".join(result.warnings) + "]"
    return text


def search_field(field: str, search_term: str, days: int, folder_name: Optional[str],
                 match_all: bool) -> Tuple[List[Dict[str, Any]], str]:
    """Search one field (subject, sender, recipient or body) and load the matches."""
    from ..mail_query import parse_search_terms

    if not search_term or not isinstance(search_term, str) or not search_term.strip():
        return [], "Search term must be a non-empty string"
    if isinstance(days, bool) or not isinstance(days, int) or days < 1 or days > 30:
        return [], "Days must be between 1 and 30"
    folder_path = get_folder_path_safe(folder_name)
    try:
        criteria = MailCriteria(days=days, match_all_terms=match_all,
                                **{field: parse_search_terms(search_term)})
        result = query_to_cache(folder_name, criteria, f"search {field} in {folder_path}")
    except Exception as e:
        logger.error(f"Search failed: {e}")
        return [], f"Error searching emails: {e}"
    if not result.emails:
        return [], f"No emails found in '{folder_path}' matching '{search_term}'"
    return result.emails, describe_result(result, folder_path, "matching emails")
