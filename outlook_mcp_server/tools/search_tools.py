"""Email search tools for Outlook MCP Server."""

# Type imports
from typing import Any, Callable, Dict, List, Optional, Union

# Local application imports
from ..backend import email_search
from ..backend.logging_config import get_logger
from ..backend.validation import (
    ValidationError,
    execute_cache_loading_operation,
    get_folder_path_safe,
    validate_days_parameter,
    validate_folder_name,
    validate_search_term
)

logger = get_logger(__name__)

# Import specific functions from the email_search module
list_recent_emails = email_search.list_recent_emails
search_email_by_subject = email_search.search_email_by_subject
search_email_by_sender = email_search.search_email_by_sender
search_email_by_recipient = email_search.search_email_by_recipient
search_email_by_body = email_search.search_email_by_body


def list_recent_emails_tool(days: int = 7, folder_name: Optional[str] = None) -> Dict[str, Any]:
    """Load emails into cache and return count message.

    Args:
        days: Days to look back (1-30, default:7, max:30)
        folder_name: Folder to search (default:Inbox, or use full path like "user@company.com/Inbox")

    Returns:
        dict: Response containing email count message:
        {
            "type": "text",
            "text": "Found X emails from last N days (max 30 days). Use 'view_email_cache_tool' to view them."
        }
        
    Note:
        For nested folders, use full path format: "user@company.com/Inbox/SubFolder"
        For top-level folders, you can use just the folder name or full path: "Inbox" or "user@company.com/Inbox"
    """
    def validate_params() -> Optional[str]:
        validate_days_parameter(days)
        return validate_folder_name(folder_name)
    
    validated_folder = validate_params()
    folder_path = get_folder_path_safe(validated_folder)
    
    return execute_cache_loading_operation(
        operation_func=list_recent_emails,
        operation_name="list_recent_emails_tool",
        message_suffix=" (max 30 days)",
        folder_name=folder_path,
        days=days
    )


def search_email_by_subject_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> Dict[str, Any]:
    """Search email subjects and load matching emails into cache.

    This function only searches the email subject field. It does not search in the email body,
    sender name, recipients, or other fields.

    Args:
        search_term: Plain text search term (colons are allowed as part of regular text)
        days: Number of days to look back (1-30, default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox, or use full path like "user@company.com/Inbox/SubFolder")
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        {
            "type": "text",
            "text": "Found X matching emails from last N days (max 30 days). Use 'view_email_cache_tool' to view them."
        }
        
    Note:
        For nested folders, use full path format: "user@company.com/Inbox/SubFolder"
        For top-level folders, you can use just the folder name or full path: "Inbox" or "user@company.com/Inbox"

    """
    def validate_params() -> Optional[str]:
        validate_search_term(search_term)
        validate_days_parameter(days)
        return validate_folder_name(folder_name)
    
    validated_folder = validate_params()
    
    return execute_cache_loading_operation(
        operation_func=search_email_by_subject,
        operation_name="search_email_by_subject_tool",
        message_suffix=f" from last {days} days (max 30 days)",
        search_term=search_term,
        days=days,
        folder_name=validated_folder,
        match_all=match_all
    )


def search_email_by_sender_name_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> Dict[str, Any]:
    """Search emails by sender name and load matching emails into cache.

    This function only searches the sender name field. It does not search in the email body,
    subject, recipients, or other fields.

    Search by name only, not email address.

    Args:
        search_term: Plain text search term for sender name (colons are allowed as part of regular text)
        days: Number of days to look back (1-30, default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox, or use full path like "user@company.com/Inbox/SubFolder")
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        {
            "type": "text",
            "text": "Found X matching emails from last N days (max 30 days). Use 'view_email_cache_tool' to view them."
        }
        
    Note:
        For nested folders, use full path format: "user@company.com/Inbox/SubFolder"
        For top-level folders, you can use just the folder name or full path: "Inbox" or "user@company.com/Inbox"

    """
    def validate_params() -> Optional[str]:
        validate_search_term(search_term)
        validate_days_parameter(days)
        return validate_folder_name(folder_name)
    
    validated_folder = validate_params()
    
    return execute_cache_loading_operation(
        operation_func=search_email_by_sender,
        operation_name="search_email_by_sender_name_tool",
        message_suffix=f" from last {days} days (max 30 days)",
        search_term=search_term,
        days=days,
        folder_name=validated_folder,
        match_all=match_all
    )


def search_email_by_recipient_name_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> Dict[str, Any]:
    """Search emails by recipient name and load matching emails into cache.

    This function only searches the recipient (To) field. It does not search in the email body,
    subject, sender, or other fields.

    Search by name only, not email address.

    Args:
        search_term: Plain text search term for recipient name (colons are allowed as part of regular text)
        days: Number of days to look back (1-30, default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox, or use full path like "user@company.com/Inbox/SubFolder")
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        {
            "type": "text",
            "text": "Found X matching emails from last N days (max 30 days). Use 'view_email_cache_tool' to view them."
        }
        
    Note:
        For nested folders, use full path format: "user@company.com/Inbox/SubFolder"
        For top-level folders, you can use just the folder name or full path: "Inbox" or "user@company.com/Inbox"

    """
    def validate_params() -> Optional[str]:
        validate_search_term(search_term)
        validate_days_parameter(days)
        return validate_folder_name(folder_name)
    
    validated_folder = validate_params()
    
    return execute_cache_loading_operation(
        operation_func=search_email_by_recipient,
        operation_name="search_email_by_recipient_name_tool",
        message_suffix=f" from last {days} days (max 30 days)",
        search_term=search_term,
        days=days,
        folder_name=validated_folder,
        match_all=match_all
    )


def search_email_by_body_tool(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True
) -> Dict[str, Any]:
    """Search emails by body content and load matching emails into cache.

    This function searches the email body content. It does not search in the subject,
    sender name, recipients, or other fields.

    Note: Searching email body is slower than searching other fields: keep the period short.

    Args:
        search_term: Plain text search term (colons are allowed as part of regular text)
                    For exact phrase matching, enclose the term in quotes (e.g., "project a status")
                    For word-based matching, use the term without quotes (e.g., project a status)
        days: Number of days to look back (1-30, default: 7, max: 30)
        folder_name: Optional folder name to search (default: Inbox, or use full path like "user@company.com/Inbox/SubFolder")
        match_all: If True, requires ALL search terms to match (AND logic, default).
                  If False, matches ANY search term (OR logic)

    Returns:
        dict: Response containing email count message
        
    Note:
        For nested folders, use full path format: "user@company.com/Inbox/SubFolder"
        For top-level folders, you can use just the folder name or full path: "Inbox" or "user@company.com/Inbox"
        {
            "type": "text",
            "text": "Found X matching emails from last N days (max 30 days). Use 'view_email_cache_tool' to view them."
        }

    """
    def validate_params() -> Optional[str]:
        validate_search_term(search_term)
        validate_days_parameter(days)
        return validate_folder_name(folder_name)
    
    validated_folder = validate_params()
    
    return execute_cache_loading_operation(
        operation_func=search_email_by_body,
        operation_name="search_email_by_body_tool",
        message_suffix=f" from last {days} days (max 30 days)",
        search_term=search_term,
        days=days,
        folder_name=validated_folder,
        match_all=match_all
    )


def _terms(value: Union[str, List[str], None]) -> List[str]:
    if value is None:
        return []
    if isinstance(value, str):
        value = value.replace(",", ";").split(";")
    return [str(v).strip() for v in value if str(v).strip()]


def search_emails_tool(
    subject_contains: Union[str, List[str], None] = None,
    sender_contains: Union[str, List[str], None] = None,
    recipient_contains: Union[str, List[str], None] = None,
    body_contains: Union[str, List[str], None] = None,
    categories: Union[str, List[str], None] = None,
    exclude_categories: Union[str, List[str], None] = None,
    unread_only: bool = False,
    flagged_only: bool = False,
    days: Optional[int] = 30,
    older_than_days: Optional[int] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    folder_path: Optional[str] = None,
    include_subfolders: bool = False,
    require_all_terms: bool = False,
    categories_match: str = "any",
    max_results: int = 200,
    limit: int = 25,
) -> Dict[str, Any]:
    """Search emails live in Outlook with combined criteria and return the matches directly.

    The search runs on the mailbox itself at call time (not on a previous listing). Criteria are
    combined with AND; values inside a list are OR (AND with require_all_terms). Matching is
    "contains", case-insensitive. Results are newest first and become the numbered listing used by
    the other tools (read, reply, move, categories, flag): numbers stay valid until the next load or search.
    Same criteria vocabulary as move_emails_by_filter_tool.

    Args:
        subject_contains: Subject must contain one of these terms
        sender_contains: Sender name or address must contain one of these terms
        recipient_contains: A To/Cc display name must contain one of these terms
        body_contains: Body must contain one of these terms (slower: keep the period short)
        categories: Only emails carrying one of these categories (all of them with categories_match="all")
        exclude_categories: Skip emails carrying one of these categories
        unread_only: Only unread emails
        flagged_only: Only emails flagged for follow-up (not completed)
        days: Only emails received in the last N days (default 30; None = no lower bound)
        older_than_days: Only emails older than N days
        date_from: 'YYYY-MM-DD' or 'YYYY-MM-DD HH:MM' local time, inclusive (alternative to days)
        date_to: Same format; a date alone includes that whole day
        folder_path: Folder to search (default Inbox), e.g. "user@company.com/Inbox/ProjectA"
        include_subfolders: Also search the subfolders
        require_all_terms: Every term of a list must match instead of any
        categories_match: "any" (default) or "all"
        max_results: Matches kept in the numbered listing (default 200, max 1000)
        limit: Matches returned in this response (default 25, max 100); page the rest with view_email_cache_tool

    Returns:
        dict: {"type": "json", "data": {"total_matched", "listed", "returned", "truncated", "emails": [...]}}
    """
    from datetime import timedelta

    from ..backend.com_helpers import parse_dt
    from ..backend.email_search.search_common import query_to_cache
    from ..backend.mail_query import MailCriteria
    from .viewing_tools import listing_rows

    try:
        if categories_match not in ("any", "all"):
            raise ValueError("categories_match must be 'any' or 'all'")
        start = parse_dt(date_from)
        end = parse_dt(date_to)
        if end is not None and ":" not in date_to:
            end += timedelta(days=1)  # date only: include the whole day
        if start is not None or end is not None:
            days = None  # explicit dates replace the default look-back
        criteria = MailCriteria(
            days=days, older_than_days=older_than_days, received_from=start, received_to=end,
            subject=_terms(subject_contains), sender=_terms(sender_contains),
            recipient=_terms(recipient_contains), body=_terms(body_contains),
            match_all_terms=bool(require_all_terms), categories=_terms(categories),
            categories_match_all=categories_match == "all", exclude_categories=_terms(exclude_categories),
            unread_only=bool(unread_only), flagged_only=bool(flagged_only),
        )
        max_results = max(1, min(int(max_results), 1000))
        limit = max(1, min(int(limit), 100))
        folder = validate_folder_name(folder_path)
        label = get_folder_path_safe(folder)
        result = query_to_cache(folder, criteria, f"search in {label}", max_results=max_results,
                                include_subfolders=bool(include_subfolders))
        data: Dict[str, Any] = {
            "folder": label,
            "criteria": criteria.describe(),
            "total_matched": result.total_matched,
            "listed": len(result.emails),
            "returned": min(limit, len(result.emails)),
            "truncated": result.truncated,
            "emails": listing_rows(0, min(limit, len(result.emails))),
        }
        if include_subfolders:
            data["folders_searched"] = len(result.folders)
        if result.warnings:
            data["warnings"] = result.warnings
        if len(result.emails) > limit:
            data["next"] = "More matches are in the listing: use view_email_cache_tool (page_size up to 50)."
        return {"type": "json", "data": data}
    except Exception as e:
        logger.error(f"Error in search_emails_tool: {e}")
        return {"type": "text", "text": f"Error searching emails: {e}"}
