"""Full content of one email of the numbered listing, read live from Outlook."""

from typing import Any, Dict, Optional

from .attachments import summarize_attachments
from .com_helpers import get_item
from .email_utils import _format_recipient_for_display
from .logging_config import get_logger
from .outlook_session.session_manager import OutlookSessionManager
from .shared import email_cache, email_cache_order
from .utils import safe_encode_text

logger = get_logger(__name__)


def _recipients(email: Dict[str, Any], key: str) -> str:
    return ", ".join(_format_recipient_for_display(r) for r in email.get(key) or [])


def create_basic_email_response(email: Dict[str, Any]) -> Dict[str, Any]:
    """Email data from the cached row only (no Outlook access)."""
    sender = email.get("sender", "Unknown Sender")
    sender_name = sender.get("name", "Unknown Sender") if isinstance(sender, dict) else str(sender)
    entry_id = email.get("entry_id") or email.get("id", "")
    return {
        "id": entry_id,
        "entry_id": entry_id,
        "subject": email.get("subject", "No Subject"),
        "sender": sender_name,
        "from": sender_name,
        "sender_email": email.get("sender_email", ""),
        "received_time": email.get("received_time", ""),
        "received": email.get("received_time", ""),
        "unread": email.get("unread", False),
        "has_attachments": email.get("has_attachments", False),
        "size": email.get("size", 0),
        "to": _recipients(email, "to_recipients"),
        "cc": _recipients(email, "cc_recipients"),
        "body": email.get("body", ""),
        "attachments": email.get("attachments", []),
        "attachments_count": len(email.get("attachments", [])),
        "categories": email.get("categories", ""),
        "flag_status": email.get("flag_status", 0),
    }


def extract_comprehensive_email_data(email: Dict[str, Any], include_html: bool = True) -> Dict[str, Any]:
    """Cached row completed with the live item: body, recipients, attachments, metadata."""
    result = create_basic_email_response(email)
    try:
        with OutlookSessionManager() as session:
            item = get_item(session, result["entry_id"], email.get("store_id"))
            if not item:
                result["warning"] = "Email not found in Outlook (moved or deleted?)"
                return result

            result["body"] = safe_encode_text(getattr(item, "Body", ""), "body")
            if include_html:
                result["html_body"] = safe_encode_text(getattr(item, "HTMLBody", ""), "html_body")
            result["body_format"] = getattr(item, "BodyFormat", 1)  # 1=Plain, 2=HTML, 3=RichText

            # Live values win over the listing: they may have changed since it was loaded.
            for key, attr in (("to", "To"), ("cc", "CC")):
                value = getattr(item, attr, None)
                if isinstance(value, str) and value:
                    result[key] = value
            summary = summarize_attachments(item)
            result["attachments"] = summary["attachments"]
            result["attachments_count"] = len(summary["attachments"])
            result["embedded_images_count"] = summary["embedded_images_count"]
            result["has_attachments"] = bool(summary["attachments"])

            result["importance"] = getattr(item, "Importance", 1)  # 0=Low, 1=Normal, 2=High
            result["sensitivity"] = getattr(item, "Sensitivity", 0)
            result["conversation_topic"] = safe_encode_text(getattr(item, "ConversationTopic", ""), "conversation_topic")
            result["conversation_id"] = getattr(item, "ConversationID", "")
            result["categories"] = getattr(item, "Categories", "") or ""
            result["flag_status"] = getattr(item, "FlagStatus", 0)  # 0=none, 1=complete, 2=flagged
            result["unread"] = bool(getattr(item, "UnRead", result["unread"]))
    except Exception as e:
        logger.error(f"Error loading email details: {e}")
        result["warning"] = f"Live content unavailable: {e}"
    return result


def extract_basic_email_data(email: Dict[str, Any]) -> Dict[str, Any]:
    """Text-only view: plain body and metadata, without the HTML body."""
    return extract_comprehensive_email_data(email, include_html=False)


def get_email_by_number_unified(email_number: int, mode: str = "basic", include_attachments: bool = True,
                                embed_images: bool = True) -> Optional[Dict[str, Any]]:
    """Email data by listing number (1-based). None when the number is not in the listing.

    mode "basic" returns the plain-text body; any other mode also returns the HTML body.
    """
    if isinstance(email_number, bool) or not isinstance(email_number, int) or email_number < 1:
        return None
    if email_number > len(email_cache_order):
        return None
    email_data = email_cache.get(email_cache_order[email_number - 1])
    if not email_data:
        return None
    if mode == "basic":
        return extract_basic_email_data(email_data)
    return extract_comprehensive_email_data(email_data)


def format_email_with_media(email_data: Dict[str, Any], max_body_chars: Optional[int] = None) -> str:
    """Readable text block for one email."""
    sender = email_data.get("from", "N/A")
    if email_data.get("sender_email"):
        sender += f" <{email_data['sender_email']}>"
    lines = [
        f"Subject: {email_data.get('subject', 'N/A')}",
        f"From: {sender}",
        f"To: {email_data.get('to') or 'N/A'}",
    ]
    if email_data.get("cc"):
        lines.append(f"Cc: {email_data['cc']}")
    lines.append(f"Date: {email_data.get('received', 'N/A')}")
    if email_data.get("conversation_topic"):
        lines.append(f"Conversation: {email_data['conversation_topic']}")
    if email_data.get("warning"):
        lines.append(f"Warning: {email_data['warning']}")

    body = email_data.get("body") or ""
    if max_body_chars and max_body_chars > 0 and len(body) > max_body_chars:
        body = body[:max_body_chars] + f"\n[... truncated, {len(email_data['body'])} characters in total]"
    lines.append(f"Body: {body or 'N/A'}")
    if email_data.get("html_body") and email_data.get("html_body") != email_data.get("body"):
        lines.append(f"HTML Body: {email_data['html_body']}")

    attachments = email_data.get("attachments") or []
    if attachments:
        lines.append(f"\nAttachments: {len(attachments)}")
        for attachment in attachments:
            size = f" ({attachment['size']} bytes)" if attachment.get("size") else ""
            lines.append(f"  - {attachment.get('name', 'Unknown')}{size}")

    if email_data.get("importance") is not None:
        lines.append("Importance: " + {0: "Low", 1: "Normal", 2: "High"}.get(email_data["importance"], "Normal"))
    if email_data.get("categories"):
        lines.append(f"Categories: {email_data['categories']}")
    flag = {1: "complete", 2: "flagged"}.get(email_data.get("flag_status"))
    if flag:
        lines.append(f"Flag: {flag}")
    return "\n".join(lines) + "\n"
