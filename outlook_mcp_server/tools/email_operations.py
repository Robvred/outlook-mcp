"""Email operations tools for Outlook MCP Server."""

from typing import Dict, Any, Union, List, Optional
from ..backend.email_composition import reply_to_email_by_number, compose_email
from ..backend.outlook_session import OutlookSessionManager
from ..backend.validation import ValidationError


def reply_to_email_by_number_tool(
    email_number: int, 
    reply_text: str, 
    to_recipients: Union[str, List[str], None] = None, 
    cc_recipients: Union[str, List[str], None] = None,
    extra_cc: Union[str, List[str], None] = None,
    importance: Optional[int] = None,
    include_signature: bool = True,
) -> Dict[str, Any]:
    """Create a DRAFT reply-all to an email. Does NOT send — the user must review and send manually from Outlook Drafts.

    Args:
        email_number: Email's position in the last listing
        reply_text: Plain text of the reply (blank lines = paragraphs). Converted to HTML and placed above the signature and quoted thread.
        to_recipients: Replaces the To line (addresses or GAL display names). None keeps reply-all To.
        cc_recipients: Replaces the Cc line. None keeps reply-all Cc.
        extra_cc: Adds people in Cc on top of reply-all (addresses or GAL display names, e.g. "FLORES Ivan Anthony").
        importance: 0=Low, 1=Normal, 2=High. None keeps Normal.
        include_signature: Insert the user's default Outlook signature.

    Behavior:
        - Always creates a draft; never sends.
        - Uses Outlook's native ReplyAll(): HTML format, quoted thread, conversation threading and SMTP recipients are preserved.
        - Response lists the final To/Cc and any unresolved recipient: show them to the user for checking.

    Returns:
        dict: {"type": "text", "text": "Draft reply ... saved to Drafts folder ..."}
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")
    if not reply_text or not isinstance(reply_text, str):
        raise ValidationError("Reply text must be a non-empty string")
    
    try:
        if isinstance(extra_cc, str):
            extra_cc = [a.strip() for a in extra_cc.split(';') if a.strip()]
        result = reply_to_email_by_number(
            email_number, reply_text, to_recipients, cc_recipients,
            extra_cc=extra_cc, importance=importance, include_signature=include_signature,
        )
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error replying to email: {str(e)}"}


def compose_email_tool(recipient_email: str, subject: str, body: str, cc_email: Optional[str] = None) -> Dict[str, Any]:
    """Compose a new email as a DRAFT. Does NOT send — the user must review and send manually from Outlook Drafts.

    Args:
        recipient_email: Email address(es) of the recipient(s) - single email or semicolon-separated list
        subject: Subject line
        body: Main content of the email
        cc_email: Optional CC address(es) - single email or semicolon-separated list

    Returns:
        dict: {"type": "text", "text": "Draft saved to Drafts folder ... EntryID=..."}
    """
    if not recipient_email or not isinstance(recipient_email, str):
        raise ValidationError("Recipient email must be a non-empty string")
    if not subject or not isinstance(subject, str):
        raise ValidationError("Subject must be a non-empty string")
    if not body or not isinstance(body, str):
        raise ValidationError("Body must be a non-empty string")
    
    try:
        # Parse semicolon-separated email addresses into lists
        to_recipients = [email.strip() for email in recipient_email.split(';') if email.strip()]
        cc_recipients = None
        if cc_email:
            cc_recipients = [email.strip() for email in cc_email.split(';') if email.strip()]
        
        result = compose_email(to_recipients, subject, body, cc_recipients)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error composing email: {str(e)}"}


def move_email_tool(email_number: int, target_folder_name: str) -> Dict[str, Any]:
    """Move an email to the specified folder.

    Args:
        email_number: The number of the email in the cache to move (1-based)
        target_folder_name: Name or path of the target folder (supports nested paths like "user@company.com/Inbox/SubFolder1/SubFolder2")

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Email moved successfully to target_folder"
        }

    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
        After moving, the cache will be cleared to reflect the new email positions.
        
        IMPORTANT: Target folder paths must include the email address as the root folder.
        Use format: "user@company.com/Inbox/SubFolder" not just "Inbox/SubFolder"
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")
    if not target_folder_name or not isinstance(target_folder_name, str):
        raise ValidationError("Target folder name must be a non-empty string")

    try:
        # Use direct email operations instead of session manager wrapper
        from ..backend.outlook_session.email_operations import move_email_to_folder
        result = move_email_to_folder(email_number, target_folder_name)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error moving email: {str(e)}"}


def delete_email_by_number_tool(email_number: int) -> Dict[str, Any]:
    """Move an email to the Deleted Items folder.

    Args:
        email_number: The number of the email in the cache to delete (1-based)

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Email moved to Deleted Items successfully"
        }

    Note:
        Requires emails to be loaded first via list_recent_emails or search_emails.
        This tool moves the email to the Deleted Items folder instead of permanently deleting it.
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")

    try:
        # Use direct email operations instead of session manager wrapper
        from ..backend.outlook_session.email_operations import delete_email_by_number
        result = delete_email_by_number(email_number)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error deleting email: {str(e)}"}