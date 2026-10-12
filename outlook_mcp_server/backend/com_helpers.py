"""Shared COM helpers: item lookup, dates, HTML insertion, recipient resolution."""

from __future__ import annotations

import html
import re
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional

from .shared import email_cache_order

# Outlook "no date" sentinel (4501-01-01)
_NO_DATE_YEAR = 4501

OL_TO = 1
OL_CC = 2


def get_item_by_number(session: Any, email_number: int) -> Any:
    """Return the live Outlook item for a cache number (1-based)."""
    if not email_cache_order:
        raise ValueError("No emails loaded in cache. Load a folder or search first.")
    if email_number < 1 or email_number > len(email_cache_order):
        raise ValueError(f"Email #{email_number} not in cache (1..{len(email_cache_order)})")
    entry_id = email_cache_order[email_number - 1]
    item = session.namespace.GetItemFromID(entry_id)
    if not item:
        raise ValueError(f"Email #{email_number} could not be retrieved from Outlook")
    return item


def parse_dt(value: Optional[str]) -> Optional[datetime]:
    """Parse 'YYYY-MM-DD' or 'YYYY-MM-DD HH:MM' (local time)."""
    if not value:
        return None
    value = value.strip().replace("T", " ")
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    raise ValueError(f"Invalid date '{value}'. Use YYYY-MM-DD or YYYY-MM-DD HH:MM")


def to_com(value: datetime) -> datetime:
    """Wall-clock datetime for Outlook COM.

    pywin32 converts naive datetimes from local time to UTC before handing them to
    Outlook, which then stores the UTC value as local time (a -2h shift in Paris summer).
    Tagging the wall-clock value as UTC makes pywin32 pass it through unchanged.
    """
    return value.replace(tzinfo=timezone.utc)


def fmt_dt(value: Any) -> Optional[str]:
    """Format a COM date, returning None for Outlook's 'no date' sentinel."""
    try:
        if value is None or getattr(value, "year", _NO_DATE_YEAR) >= _NO_DATE_YEAR:
            return None
        return value.strftime("%Y-%m-%d %H:%M") if (value.hour or value.minute) else value.strftime("%Y-%m-%d")
    except Exception:
        return None


def list_separator() -> str:
    """Windows list separator, used by Outlook to split Categories (';' in fr-FR, ',' in en-US)."""
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\International") as key:
            sep, _ = winreg.QueryValueEx(key, "sList")
            return sep or ";"
    except Exception:
        return ";"


def split_categories(raw: Optional[str]) -> List[str]:
    if not raw:
        return []
    return [c.strip() for c in re.split(r"[;,]", str(raw)) if c.strip()]


def join_categories(cats: List[str]) -> str:
    sep = list_separator()
    return f"{sep} ".join(dict.fromkeys(cats))


def text_to_html(text: str) -> str:
    """Plain text -> simple HTML paragraphs, preserving single line breaks."""
    paragraphs = re.split(r"\n\s*\n", text.strip())
    parts = []
    for p in paragraphs:
        escaped = html.escape(p).replace("\n", "<br>")
        parts.append(f"<p style=\"margin:0 0 10pt 0\">{escaped}</p>")
    return (
        "<div style=\"font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt\">"
        + "".join(parts)
        + "</div>"
    )


def insert_html_at_top(html_body: str, snippet: str) -> str:
    """Insert snippet right after the <body> tag (above signature and quoted thread)."""
    match = re.search(r"<body[^>]*>", html_body or "", flags=re.IGNORECASE)
    if not match:
        return snippet + (html_body or "")
    pos = match.end()
    return html_body[:pos] + snippet + html_body[pos:]


def recipient_smtp(recipient: Any) -> str:
    """Best-effort SMTP address for a Recipient (Exchange users return X.500 by default)."""
    try:
        entry = recipient.AddressEntry
        exch = entry.GetExchangeUser()
        if exch is not None and exch.PrimarySmtpAddress:
            return exch.PrimarySmtpAddress
    except Exception:
        pass
    try:
        addr = recipient.Address or ""
        if "@" in addr:
            return addr
    except Exception:
        pass
    return ""


def describe_recipients(mail: Any) -> Dict[str, List[str]]:
    """Return {'to': [...], 'cc': [...], 'unresolved': [...]} with 'Name <smtp>' strings."""
    out: Dict[str, List[str]] = {"to": [], "cc": [], "unresolved": []}
    for i in range(1, mail.Recipients.Count + 1):
        r = mail.Recipients.Item(i)
        label = f"{r.Name} <{recipient_smtp(r)}>" if recipient_smtp(r) else r.Name
        if not r.Resolved:
            out["unresolved"].append(r.Name)
        if r.Type == OL_TO:
            out["to"].append(label)
        elif r.Type == OL_CC:
            out["cc"].append(label)
    return out
