"""Email composition and reply functions with improved encoding handling"""

# Type imports
from typing import Any, Callable, Dict, List, Optional, Union

# Local application imports
from .audit import log_event
from .logging_config import get_logger
from .outlook_session.session_manager import OutlookSessionManager
from .shared import email_cache, email_cache_order
from .utils import safe_encode_text, normalize_email_address
from .com_helpers import (
    OL_CC,
    OL_TO,
    describe_recipients,
    insert_html_at_top,
    text_to_html,
)
from .validation import (
    DisplayConstants,
    OutlookConstants,
    ValidationError,
    validate_cache_available,
    validate_email_number
)
from .validators import EmailComposeParams, EmailForwardParams, EmailReplyParams

logger = get_logger(__name__)


def reply_to_email_by_number(
    email_number: int,
    reply_text: str,
    to_recipients: Optional[Union[str, List[str]]] = None,
    cc_recipients: Optional[Union[str, List[str]]] = None,
    extra_cc: Optional[List[str]] = None,
    importance: Optional[int] = None,
    include_signature: bool = True,
) -> str:
    """
    Reply to an email with custom recipients if provided.

    Args:
        email_number: Email's position in the last listing
        reply_text: Text to prepend to the reply
        to_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
        cc_recipients: Either a single email string OR a list of email strings (None preserves original recipients)
        extra_cc: Addresses or GAL display names added in CC on top of the reply-all recipients
        importance: 0=Low, 1=Normal, 2=High (None keeps default)
        include_signature: Insert the default Outlook signature (default True)

    Returns:
        str: Success or error message, including the final To/Cc lists
    """
    # Validate inputs using Pydantic
    try:
        params = EmailReplyParams(
            email_number=email_number,
            reply_text=reply_text,
            to_recipients=to_recipients,
            cc_recipients=cc_recipients,
        )
    except Exception as e:
        logger.error(f"Validation error in reply_to_email_by_number: {e}")
        raise ValueError(f"Invalid parameters: {e}")

    # Convert to list if needed (validator already did this)
    to_recipients = params.to_recipients
    cc_recipients = params.cc_recipients
    reply_text = params.reply_text

    try:
        validate_cache_available(len(email_cache_order))
        validate_email_number(email_number, len(email_cache_order))
    except ValidationError as e:
        logger.error(f"Validation error in reply_to_email_by_number: {e}")
        raise ValueError(f"Invalid parameters: {e}")

    # Get the entry_id from the cache order
    entry_id = email_cache_order[email_number - 1]
    if not entry_id:
        raise ValueError(f"Email #{email_number} has no entry ID")

    # Get the cached email data
    cached_email = email_cache.get(entry_id)
    if not cached_email:
        raise ValueError(f"Email #{email_number} data not found in cache")

    with OutlookSessionManager() as session:
        try:
            email_id = cached_email.get("id") or cached_email.get("entry_id") or entry_id
            email = session.namespace.GetItemFromID(email_id)
            if not email:
                raise RuntimeError("Could not retrieve the email from Outlook.")

            # Native Outlook reply: keeps HTML format, quoted thread, conversation
            # threading and the correct SMTP recipients (Exchange users included).
            custom = to_recipients is not None or cc_recipients is not None
            reply = email.ReplyAll()

            # Accessing the inspector makes Outlook insert the default signature.
            if include_signature:
                try:
                    _ = reply.GetInspector
                except Exception as e:
                    logger.debug(f"Signature insertion skipped: {e}")

            if custom:
                # Replace only the fields that were explicitly provided.
                for i in range(reply.Recipients.Count, 0, -1):
                    rtype = reply.Recipients.Item(i).Type
                    if (rtype == OL_TO and to_recipients is not None) or (
                        rtype == OL_CC and cc_recipients is not None
                    ):
                        reply.Recipients.Remove(i)
                for addr in to_recipients or []:
                    reply.Recipients.Add(addr).Type = OL_TO
                for addr in cc_recipients or []:
                    reply.Recipients.Add(addr).Type = OL_CC
            for addr in extra_cc or []:
                reply.Recipients.Add(addr).Type = OL_CC
            reply.Recipients.ResolveAll()

            if importance is not None:
                reply.Importance = int(importance)

            reply.HTMLBody = insert_html_at_top(reply.HTMLBody, text_to_html(reply_text))

            # Draft-first guardrail: save as draft instead of sending.
            reply.Save()
            draft_entry_id = getattr(reply, "EntryID", "")
            recipients = describe_recipients(reply)
            log_event(
                "reply_draft_created",
                email_number=email_number,
                draft_entry_id=draft_entry_id,
                to=recipients["to"],
                cc=recipients["cc"],
            )
            logger.info(f"Draft reply created for email #{email_number} (EntryID={draft_entry_id})")
            warning = (
                f" WARNING unresolved recipients: {', '.join(recipients['unresolved'])}."
                if recipients["unresolved"]
                else ""
            )
            return (
                f"Draft reply to email #{email_number} saved to Drafts folder (HTML, threaded). "
                f"To: {'; '.join(recipients['to']) or '-'} | Cc: {'; '.join(recipients['cc']) or '-'}.{warning} "
                f"Review and send it manually in Outlook. EntryID={draft_entry_id}"
            )

        except Exception as e:
            logger.error(f"Error replying to email #{email_number}: {e}")
            return f"Error replying to email: {str(e)}"


def forward_email_by_number(
    email_number: int,
    to_recipients: Union[str, List[str]],
    cc_recipients: Optional[Union[str, List[str]]] = None,
    forward_text: str = "",
    include_attachments: bool = True,
    importance: Optional[int] = None,
    include_signature: bool = True,
) -> str:
    """
    Forward one cached email to explicit recipients, as a draft.

    Uses Outlook's native Forward(): the original message (header, body, thread) and its
    attachments are carried over, unlike a reply. Nothing is sent.

    Args:
        email_number: Email's position in the last listing
        to_recipients: One or more addresses or GAL display names (string with ";" or list)
        cc_recipients: Optional Cc recipients (string with ";" or list)
        forward_text: Plain text placed above the forwarded message
        include_attachments: Keep the original attachments (default True)
        importance: 0=Low, 1=Normal, 2=High (None keeps default)
        include_signature: Insert the default Outlook signature (default True)

    Returns:
        str: Success or error message, including the final To/Cc lists and attachment count
    """
    try:
        params = EmailForwardParams(
            email_number=email_number,
            to_recipients=to_recipients,
            cc_recipients=cc_recipients,
            forward_text=forward_text or "",
        )
    except Exception as e:
        logger.error(f"Validation error in forward_email_by_number: {e}")
        raise ValueError(f"Invalid parameters: {e}")

    try:
        validate_cache_available(len(email_cache_order))
        validate_email_number(email_number, len(email_cache_order))
    except ValidationError as e:
        logger.error(f"Validation error in forward_email_by_number: {e}")
        raise ValueError(f"Invalid parameters: {e}")

    entry_id = email_cache_order[email_number - 1]
    if not entry_id:
        raise ValueError(f"Email #{email_number} has no entry ID")
    cached_email = email_cache.get(entry_id)
    if not cached_email:
        raise ValueError(f"Email #{email_number} data not found in cache")

    with OutlookSessionManager() as session:
        try:
            email_id = cached_email.get("id") or cached_email.get("entry_id") or entry_id
            email = session.namespace.GetItemFromID(email_id)
            if not email:
                raise RuntimeError("Could not retrieve the email from Outlook.")

            fwd = email.Forward()

            # Accessing the inspector makes Outlook insert the default signature.
            if include_signature:
                try:
                    _ = fwd.GetInspector
                except Exception as e:
                    logger.debug(f"Signature insertion skipped: {e}")

            # A forward starts with no recipients; clear defensively anyway.
            for i in range(fwd.Recipients.Count, 0, -1):
                fwd.Recipients.Remove(i)
            for addr in params.to_recipients:
                fwd.Recipients.Add(addr).Type = OL_TO
            for addr in params.cc_recipients:
                fwd.Recipients.Add(addr).Type = OL_CC
            fwd.Recipients.ResolveAll()

            if not include_attachments:
                for i in range(fwd.Attachments.Count, 0, -1):
                    fwd.Attachments.Remove(i)

            if importance is not None:
                fwd.Importance = int(importance)

            if params.forward_text.strip():
                fwd.HTMLBody = insert_html_at_top(fwd.HTMLBody, text_to_html(params.forward_text))

            # Draft-first guardrail: save as draft instead of sending.
            fwd.Save()
            draft_entry_id = getattr(fwd, "EntryID", "")
            recipients = describe_recipients(fwd)
            try:
                attachment_count = int(fwd.Attachments.Count)
            except Exception:
                attachment_count = 0
            log_event(
                "forward_draft_created",
                email_number=email_number,
                draft_entry_id=draft_entry_id,
                to=recipients["to"],
                cc=recipients["cc"],
                attachments=attachment_count,
            )
            logger.info(f"Draft forward created for email #{email_number} (EntryID={draft_entry_id})")
            warning = (
                f" WARNING unresolved recipients: {', '.join(recipients['unresolved'])}."
                if recipients["unresolved"]
                else ""
            )
            return (
                f"Draft forward of email #{email_number} saved to Drafts folder (HTML, original "
                f"message included, {attachment_count} attachment(s)). "
                f"To: {'; '.join(recipients['to']) or '-'} | Cc: {'; '.join(recipients['cc']) or '-'}.{warning} "
                f"Review and send it manually in Outlook. EntryID={draft_entry_id}"
            )

        except Exception as e:
            logger.error(f"Error forwarding email #{email_number}: {e}")
            return f"Error forwarding email: {str(e)}"


def compose_email(
    to_recipients: List[str],
    subject: str,
    body: str,
    cc_recipients: Optional[List[str]] = None,
    html: bool = False,
) -> str:
    """
    Compose and send a new email using Outlook COM API.

    Args:
        to_recipients: List of recipient email addresses
        subject: Email subject line
        body: Email body content
        cc_recipients: Optional list of CC email addresses
        html: If True, body is treated as HTML (default: False)

    Returns:
        str: Success/error message
    """
    # Validate inputs using Pydantic
    try:
        params = EmailComposeParams(
            recipient_email=to_recipients[0] if to_recipients else "",
            subject=subject,
            body=body,
            cc_email=cc_recipients[0] if cc_recipients else None,
        )
    except Exception as e:
        logger.error(f"Validation error in compose_email: {e}")
        raise ValueError(f"Invalid parameters: {e}")

    # Additional validation for list
    if not to_recipients or not isinstance(to_recipients, list):
        raise ValueError("To recipients must be a non-empty list")

    if not all(isinstance(email, str) and email.strip() for email in to_recipients):
        raise ValueError("All recipient email addresses must be non-empty strings")

    if cc_recipients is not None:
        if not isinstance(cc_recipients, list):
            raise ValueError("CC recipients must be a list or None")
        if not all(isinstance(email, str) and email.strip() for email in cc_recipients):
            raise ValueError("All CC email addresses must be non-empty strings")

    with OutlookSessionManager() as session:
        try:
            # Encode all components safely
            encoded_to = [
                safe_encode_text(recipient, "to_recipient").strip() for recipient in to_recipients
            ]
            subject_safe = safe_encode_text(subject, "subject")
            body_safe = safe_encode_text(body, "body")

            encoded_cc = []
            if cc_recipients:
                encoded_cc = [
                    safe_encode_text(recipient, "cc_recipient").strip()
                    for recipient in cc_recipients
                ]

            # Create and send the email
            mail = session.outlook.CreateItem(OutlookConstants.OL_MAIL_ITEM)
            mail.To = "; ".join(encoded_to)
            mail.Subject = subject_safe

            if cc_recipients:
                mail.CC = "; ".join(encoded_cc)

            try:
                # Accessing the inspector makes Outlook insert the default signature.
                try:
                    _ = mail.GetInspector
                except Exception as e:
                    logger.debug(f"Signature insertion skipped: {e}")
                snippet = body_safe if html else text_to_html(body_safe)
                mail.HTMLBody = insert_html_at_top(mail.HTMLBody, snippet)
            except Exception as e:
                logger.warning(f"Failed to set HTML body, using plain text: {e}")
                mail.Body = body_safe
            mail.Recipients.ResolveAll()

            # Draft-first guardrail: save as draft, do not send.
            mail.Save()
            draft_entry_id = getattr(mail, "EntryID", "")
            log_event(
                "compose_draft_created",
                to=encoded_to,
                cc=encoded_cc or None,
                subject=subject_safe,
                draft_entry_id=draft_entry_id,
            )
            logger.info(
                f"Draft created for {len(to_recipients)} recipients (EntryID={draft_entry_id})"
            )
            return (
                f"Draft saved to Drafts folder for {len(to_recipients)} recipient(s). "
                f"Review and send it manually in Outlook. EntryID={draft_entry_id}"
            )

        except Exception as e:
            logger.error(f"Error composing email: {e}")
            return f"Error composing email: {str(e)}"
