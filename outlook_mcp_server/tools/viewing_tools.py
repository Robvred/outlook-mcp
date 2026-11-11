"""Email viewing tools for Outlook MCP Server."""

# Type imports
from typing import Any, Dict, List, Optional, Union

# Local application imports
from ..backend.email_data_extractor import format_email_with_media, get_email_by_number_unified
from ..backend.outlook_session import OutlookSessionManager
from ..backend.shared import MAX_CACHE_SIZE, cache_meta, clear_email_cache, email_cache, email_cache_order
from ..backend.validation import (
    ValidationError,
    validate_cache_available,
    validate_days_parameter,
    validate_email_number,
    validate_folder_name,
    validate_page_parameter
)


_FLAG_LABELS = {0: "", 1: "complete", 2: "flagged"}


def _people(recipients: List[Dict[str, Any]]) -> str:
    if not recipients:
        return "N/A"
    shown = ", ".join(r.get("name") or r.get("address") or "Unknown" for r in recipients[:3])
    return shown + (f" and {len(recipients) - 3} more" if len(recipients) > 3 else "")


def listing_row(number: int, email_data: Dict[str, Any]) -> Dict[str, Any]:
    """One email of the numbered listing, as shown by the view and search tools."""
    sender = email_data.get("sender", "Unknown")
    row = {
        "number": number,
        "subject": email_data.get("subject", "No Subject"),
        "from": sender.get("name", "Unknown") if isinstance(sender, dict) else str(sender),
        "to": _people(email_data.get("to_recipients", [])),
        "cc": _people(email_data.get("cc_recipients", [])),
        "received": email_data.get("received_time", "Unknown"),
        "status": "Unread" if email_data.get("unread", False) else "Read",
        "attachments_count": len(email_data.get("attachments", [])),
        "embedded_images_count": email_data.get("embedded_images_count", 0),
        "categories": email_data.get("categories", "") or "",
        "flag": _FLAG_LABELS.get(email_data.get("flag_status", 0), ""),
    }
    if email_data.get("sender_email"):
        row["from_email"] = email_data["sender_email"]
    if email_data.get("folder"):
        row["folder"] = email_data["folder"]
    message_class = email_data.get("message_class") or ""
    if message_class and not message_class.upper().startswith("IPM.NOTE"):
        row["type"] = message_class
    return row


def listing_rows(start: int, end: int, with_attachments: bool = True) -> List[Dict[str, Any]]:
    """Rows [start, end) of the listing (0-based), attachment details resolved for these rows only."""
    ids = email_cache_order[start:end]
    entries = [email_cache[i] for i in ids if i in email_cache]
    if with_attachments and any(not e.get("attachments_processed") for e in entries):
        try:
            from ..backend.attachments import enrich_attachments
            with OutlookSessionManager() as session:
                enrich_attachments(session, entries)
        except Exception:
            pass  # the listing is still valid without attachment details
    return [listing_row(start + offset + 1, email_cache[i]) for offset, i in enumerate(ids) if i in email_cache]


def view_email_cache_tool(page: int = 1, page_size: int = 5) -> Dict[str, Any]:
    """View the emails of the last load or search (numbered listing), page by page.

    Each email shows number, subject, from, to, cc, received (local time), status (Read/Unread),
    attachments_count (real attachments), embedded_images_count, categories and flag.
    Numbers are the handles used by the other tools; they stay valid until the next load or search.

    Args:
        page: Page number (1-based)
        page_size: Emails per page (default 5, max 50). Use a larger page to review a listing in one call.

    Returns:
        dict: {"type": "json", "data": {"page", "total_pages", "total_emails", "listing", "emails": [...]}}
    """
    try:
        validate_page_parameter(page, 0)
    except ValidationError as e:
        raise ValidationError(str(e))
    if isinstance(page_size, bool) or not isinstance(page_size, int) or not 1 <= page_size <= 50:
        raise ValidationError("page_size must be between 1 and 50")

    try:
        total = len(email_cache_order)
        if not total:
            return {
                "type": "json",
                "data": {
                    "error": "No emails in cache",
                    "message": "Please load emails first using list_recent_emails or search functions."
                }
            }
        total_pages = (total + page_size - 1) // page_size
        start_idx = (page - 1) * page_size
        if start_idx >= total:
            return {
                "type": "json",
                "data": {
                    "error": "Page out of range",
                    "message": f"Page {page} is out of range. Available range: 1-{total_pages}"
                }
            }
        return {
            "type": "json",
            "data": {
                "page": page,
                "total_pages": total_pages,
                "total_emails": total,
                "listing": dict(cache_meta),
                "emails": listing_rows(start_idx, min(start_idx + page_size, total)),
            }
        }
    except Exception as e:
        return {
            "type": "json",
            "data": {
                "error": "Error viewing email cache",
                "message": str(e)
            }
        }


def get_email_by_number_tool(email_number: int, mode: str = "basic", include_attachments: bool = True, embed_images: bool = True, max_body_chars: Optional[int] = None) -> Dict[str, Any]:
    """Get email content by cache number with 3 retrieval modes.
    
    Mode Selection Guide:
    - "basic": Plain-text body with headers, categories and attachment names (no HTML) - default, cheapest
    - "enhanced": Same plus the raw HTML body - only when the markup itself is needed
    - "lazy": Auto-adapts cached vs live data - use when unsure
    
    Email Thread Handling:
    - "basic": No conversation threads (focus on individual email content)
    - "enhanced": Shows complete conversation thread
    - "lazy": Auto-adaptive thread handling
    
    Requires emails to be loaded first via list_recent_emails or search_emails.
    
    Args:
        email_number: Position in cache (1-based)
        mode: "basic" (text-only), "enhanced" (complete), "lazy" (adaptive)
        include_attachments: Include file content (enhanced mode only)
        embed_images: Embed inline images as data URIs (enhanced mode only)
        max_body_chars: Truncate the body to this many characters (default: no limit)

    Returns:
        dict: Response containing email details based on requested mode
        {
            "type": "text", 
            "text": "Formatted email content"
        }

    Raises:
        ValueError: If email number is invalid or no emails are loaded
        RuntimeError: If cache contains invalid data
    """
    try:
        validate_email_number(email_number, len(email_cache_order))
    except ValidationError as e:
        raise ValidationError(str(e))
    
    if mode not in ["basic", "enhanced", "lazy"]:
        raise ValidationError("Mode must be one of: basic, enhanced, lazy")
    
    try:
        email_data = get_email_by_number_unified(
            email_number, 
            mode=mode, 
            include_attachments=include_attachments, 
            embed_images=embed_images
        )
        
        if email_data is None:
            return {
                "type": "text", 
                "text": f"No email found at position {email_number}. Please load emails first using list_recent_emails or search_emails."
            }
        
        # Format the email with media
        formatted_email = format_email_with_media(email_data, max_body_chars=max_body_chars)
        return {"type": "text", "text": formatted_email}
        
    except ValidationError as e:
        return {"type": "text", "text": f"Validation error: {str(e)}"}
    except RuntimeError as e:
        return {"type": "text", "text": f"Runtime error: {str(e)}"}
    except Exception as e:
        return {"type": "text", "text": f"Error retrieving email: {str(e)}"}


def load_emails_by_folder_tool(folder_path: str, days: int = None, max_emails: int = None) -> Dict[str, Any]:
    """Load emails from a specific folder into cache.

    **LLM Note**: This function enforces strict mutual exclusion between 'days' and 'max_emails' parameters.
    You CANNOT use both parameters together. Choose either time-based loading (days) or number-based loading (max_emails).
    Attempting to use both parameters will raise a ValidationError.

    Args:
        folder_path: Path to the folder (supports nested paths like "user@company.com/Inbox/SubFolder1")
        days: Number of days to look back (max: 30) - mutually exclusive with max_emails
        max_emails: Maximum number of emails to load (mutually exclusive with days) - when specified, loads the most recent emails up to this count

    Returns:
        dict: Response containing email count message

    Note:
        Loaded emails carry their categories, flag status and unread state (shown by view_email_cache_tool).
        At most 1000 emails are loaded, newest first; the message says when more exist.
        Folder paths: "user@company.com/Inbox/SubFolder1" or, for the default mailbox, "Inbox/SubFolder1".
        
        Usage examples:
        - Time-based: load_emails_by_folder_tool("Inbox", days=7)
        - Number-based: load_emails_by_folder_tool("Inbox", max_emails=50)
        - Cannot use both: load_emails_by_folder_tool("Inbox", days=7, max_emails=50) - this will raise an error
    """
    try:
        validate_folder_name(folder_path)
    except ValidationError as e:
        return {"type": "text", "text": f"Validation error: {str(e)}"}
    
    validated_folder = folder_path
    
    # Enforce mutual exclusion: cannot use both days and max_emails together
    if days is not None and max_emails is not None:
        return {"type": "text", "text": "Cannot specify both 'days' and 'max_emails' parameters. Use either time-based (days) or number-based (max_emails) loading, not both."}
    
    # Set default behavior if neither parameter is specified
    if days is None and max_emails is None:
        days = 7  # Default to 7 days if neither parameter is specified
    
    # Validate parameters
    try:
        if days is not None:
            validate_days_parameter(days)
        
        if max_emails is not None and (not isinstance(max_emails, int) or max_emails < 1):
            raise ValidationError("max_emails must be a positive integer when specified")
    except ValidationError as e:
        return {"type": "text", "text": f"Validation error: {str(e)}"}
    
    try:
        # Number-based: the N most recent. Time-based: everything in the window, up to the listing size.
        actual_max_emails = min(max_emails, MAX_CACHE_SIZE) if max_emails is not None else MAX_CACHE_SIZE
        with OutlookSessionManager() as outlook_session:
            email_list, message = outlook_session.get_folder_emails(
                validated_folder, actual_max_emails, fast_mode=True,
                days_filter=days if max_emails is None else None)
        if message.startswith("Error"):
            return {"type": "text", "text": message}
        return {"type": "text", "text": message + ". Use 'view_email_cache_tool' to view them."}
    except Exception as e:
        return {"type": "text", "text": f"Error loading emails from folder: {str(e)}"}


def clear_email_cache_tool() -> dict:
    """Clear the email cache both in memory and on disk.

    This tool removes all cached emails from memory and deletes the persistent
    cache file from disk. Use this when you want to free up memory or ensure
    fresh data is loaded from Outlook.

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Email cache cleared successfully"
        }
    """
    try:
        # Get current cache size for confirmation
        cache_size = len(email_cache_order)
        
        # Clear the cache
        clear_email_cache()
        
        return {
            "type": "text", 
            "text": f"Email cache cleared successfully. Removed {cache_size} cached emails."
        }
        
    except Exception as e:
        return {
            "type": "text", 
            "text": f"Error clearing email cache: {str(e)}"
        }