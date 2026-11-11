"""Batch email operations for forwarding and replying to multiple recipients."""

# Standard library imports
import csv

# Local application imports
from .audit import log_event
from .logging_config import get_logger
from .outlook_session.session_manager import OutlookSessionManager
from .com_helpers import cached_entry, get_item
from .utils import safe_encode_text
from .validation import (
    BatchLimits,
    BodyFormat,
    DisplayConstants,
    OutlookConstants,
    ValidationError,
    validate_email_address
)

logger = get_logger(__name__)


def batch_forward_emails(email_number: int, csv_path: str, custom_text: str = "") -> str:
    """Forward email to recipients in batches of 500 (Outlook BCC limit)"""
    # Input validation
    if not isinstance(email_number, int) or email_number < 1:
        raise ValidationError("Email number must be a positive integer")

    if not csv_path or not isinstance(csv_path, str):
        raise ValidationError("CSV path must be a non-empty string")

    if not isinstance(custom_text, str):
        raise ValidationError("Custom text must be a string")

    try:
        email_data = cached_entry(email_number)
    except ValueError as e:
        raise ValidationError(str(e))

    try:
        # Clean and validate CSV path
        clean_path = csv_path.strip("\"'")

        # Read recipients from CSV
        with open(clean_path, "r", newline="", encoding="utf-8-sig") as csvfile:
            reader = csv.DictReader(csvfile)
            if "email" not in reader.fieldnames:
                raise ValidationError("CSV must contain an 'email' column")

            recipients = []
            invalid_emails = []

            for row in reader:
                email = row.get("email", "").strip()
                if email:
                    try:
                        validated_email = validate_email_address(email)
                        recipients.append(validated_email)
                    except ValidationError:
                        invalid_emails.append(email)
                        logger.warning(f"Invalid email address found: {email}")

        if invalid_emails:
            raise ValidationError(
                f"Invalid email addresses found: {', '.join(invalid_emails[:5])}{'...' if len(invalid_emails) > 5 else ''}"
            )

        if not recipients:
            raise ValidationError("No valid email addresses found in CSV")

        # Process in batches of 500 (Outlook BCC limit)
        batch_size = BatchLimits.OUTLOOK_BCC_LIMIT
        batches = [recipients[i : i + batch_size] for i in range(0, len(recipients), batch_size)]
        total_recipients = len(recipients)
        results = []

        with OutlookSessionManager() as session:
            template = get_item(session, email_data["entry_id"], email_data.get("store_id"))

            for i, batch in enumerate(batches, 1):
                try:
                    # Create a regular mail item instead of using Forward()
                    mail = session.outlook.CreateItem(OutlookConstants.OL_MAIL_ITEM)

                    # Copy relevant properties from template with encoding handling
                    try:
                        # Handle subject encoding using safe utility
                        subject = safe_encode_text(template.Subject, "batch_subject")
                        mail.Subject = f"FW: {subject}"
                    except Exception as e:
                        logger.error(f"Encoding error in batch subject: {e}")
                        mail.Subject = "FW: [Subject encoding error]"

                    mail.BCC = "; ".join(batch)

                    # Copy body content from template with proper encoding and email headers
                    try:
                        # Extract email metadata for headers
                        sender_name = safe_encode_text(
                            getattr(template, "SenderName", "Unknown Sender"), "sender_name"
                        )
                        sent_on = safe_encode_text(
                            str(getattr(template, "SentOn", "Unknown")), "sent_on"
                        )
                        to_field = safe_encode_text(getattr(template, "To", "Unknown"), "to_field")
                        subject = safe_encode_text(
                            getattr(template, "Subject", "No Subject"), "subject"
                        )

                        if hasattr(template, "HTMLBody") and template.HTMLBody:
                            mail.BodyFormat = BodyFormat.OL_FORMAT_HTML
                            html_body = safe_encode_text(template.HTMLBody, "batch_html_body")

                            # Build HTML email headers
                            header_html = f"""
<div>
{'' if not custom_text else f'<div>{safe_encode_text(custom_text, "batch_custom_text")}</div><br>'}
<div style="margin-bottom: 10px;">__________________________________________________</div>
<div><strong>From:</strong> {sender_name}</div>
<div><strong>Sent:</strong> {sent_on}</div>
<div><strong>To:</strong> {to_field}</div>
<div><strong>Subject:</strong> {subject}</div>
<div style="margin-top: 10px; margin-bottom: 10px;">__________________________________________________</div>
</div>
<br><br>"""

                            mail.HTMLBody = header_html + html_body
                        else:
                            mail.BodyFormat = BodyFormat.OL_FORMAT_PLAIN
                            plain_body = safe_encode_text(
                                getattr(template, "Body", ""), "batch_plain_body"
                            )

                            # Build plain text email headers
                            header_lines = []
                            if custom_text:
                                header_lines.append(
                                    safe_encode_text(custom_text, "batch_custom_text")
                                )
                            header_lines.extend(
                                [
                                    "",
                                    "_" * DisplayConstants.SEPARATOR_LINE_LENGTH,
                                    f"From: {sender_name}",
                                    f"Sent: {sent_on}",
                                    f"To: {to_field}",
                                    f"Subject: {subject}",
                                    "_" * DisplayConstants.SEPARATOR_LINE_LENGTH,
                                    "",
                                ]
                            )

                            mail.Body = "\n".join(header_lines) + plain_body
                    except Exception as e:
                        logger.error(f"Error processing batch body: {e}")
                        mail.Body = "[Content processing error - please view original email]"

                    # Draft-first guardrail: save as draft, do not send.
                    mail.Save()
                    draft_entry_id = getattr(mail, "EntryID", "")
                    log_event(
                        "batch_forward_draft_created",
                        batch=i,
                        recipients=len(batch),
                        draft_entry_id=draft_entry_id,
                    )
                    logger.info(
                        f"Batch {i} draft created for {len(batch)} recipients "
                        f"(EntryID={draft_entry_id})"
                    )
                    results.append(
                        f"Batch {i}: draft saved for {len(batch)} recipients "
                        f"(EntryID={draft_entry_id})"
                    )
                except Exception as e:
                    logger.error(f"Error sending batch {i}: {e}")
                    results.append(f"Error sending batch {i}: {str(e)}")

        return "\n".join(
            [
                f"Batch drafts created for {total_recipients} recipients across {len(batches)} batches. "
                f"Review and send manually from Outlook Drafts folder:",
                *results,
            ]
        )

    except Exception as e:
        logger.error(f"Error in batch sending process: {e}")
        return f"Error in batch sending process: {str(e)}"
