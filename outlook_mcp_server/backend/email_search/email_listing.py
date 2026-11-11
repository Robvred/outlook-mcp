"""Listing of the recent emails of a folder."""

from typing import Any, Dict, List, Optional, Tuple

from ..logging_config import get_logger
from ..mail_query import MailCriteria
from ..validators import EmailListParams
from .search_common import describe_result, get_folder_path_safe, query_to_cache

logger = get_logger(__name__)


def get_emails_from_folder(folder_name: str = "Inbox", days: int = 7) -> Tuple[List[Dict[str, Any]], str]:
    """Load the emails of the last N days of a folder into the numbered listing."""
    try:
        params = EmailListParams(folder_name=folder_name, days=days)
    except Exception as e:
        return [], f"Error: Invalid parameters: {e}"
    label = get_folder_path_safe(params.folder_name)
    try:
        result = query_to_cache(params.folder_name, MailCriteria(days=params.days),
                                f"last {params.days} days of {label}")
    except Exception as e:
        logger.error(f"Error getting emails from folder: {e}")
        return [], f"Error: Failed to get emails from folder '{label}': {e}"
    return result.emails, describe_result(result, label) + f" from last {params.days} days"


get_emails_from_folder_optimized = get_emails_from_folder


def list_recent_emails(folder_name: str = "Inbox", days: Optional[int] = None) -> Tuple[List[Dict[str, Any]], str]:
    """Load recent emails (default 30 days) and return (emails, message)."""
    try:
        EmailListParams(days=days or 30, folder_name=folder_name)
    except Exception as e:
        raise ValueError(f"Invalid parameters: {e}")
    emails, message = get_emails_from_folder(folder_name, days or 30)
    if message.startswith("Error:"):
        raise RuntimeError(message[len("Error: "):])
    return emails, message
