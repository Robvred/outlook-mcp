"""Save email attachments to disk and create shortcuts for online file links.

- Real attachments (not inline signature images, not hidden parts) are saved with
  Attachment.SaveAsFile into a target directory. Existing files are never
  overwritten: identical name+size is skipped, otherwise a numeric suffix is added.
- Links to online files (SharePoint, OneDrive, Teams, Google Drive, Box, Dropbox)
  found in the HTML body become Windows .url shortcuts in the same directory.
"""

from __future__ import annotations

import html
import os
import re
from typing import Any, Dict, List, Optional
from urllib.parse import unquote, urlparse

from .audit import log_event
from .com_helpers import get_item_by_number
from .logging_config import get_logger
from .outlook_session.session_manager import OutlookSessionManager

logger = get_logger(__name__)

PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
IMAGE_EXT = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg", ".ico", ".emz", ".wmz")

ONLINE_HOSTS = re.compile(
    r"(sharepoint\.com|onedrive\.live\.com|1drv\.ms|my\.sharepoint|teams\.microsoft\.com/l/file|"
    r"docs\.google\.com|drive\.google\.com|box\.com|dropbox\.com)",
    re.IGNORECASE,
)
ANCHOR = re.compile(r'<a\b[^>]*href\s*=\s*"([^"]+)"[^>]*>(.*?)</a>', re.IGNORECASE | re.DOTALL)
BARE_URL = re.compile(r"https?://[^\s<>\"']+", re.IGNORECASE)
_BAD = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def safe_name(name: str, max_len: int = 150) -> str:
    name = _BAD.sub("_", html.unescape(name or "")).strip().strip(".")
    name = re.sub(r"\s+", " ", name)
    return (name or "fichier")[:max_len]


def unique_path(directory: str, filename: str, size: Optional[int] = None) -> Optional[str]:
    """Return a free path; None if an identical file (same name and size) already exists."""
    path = os.path.join(directory, filename)
    if os.path.exists(path):
        if size is not None and os.path.getsize(path) == size:
            return None
        stem, ext = os.path.splitext(filename)
        i = 2
        while os.path.exists(os.path.join(directory, f"{stem} ({i}){ext}")):
            i += 1
        path = os.path.join(directory, f"{stem} ({i}){ext}")
    return path


def _is_inline(att: Any) -> bool:
    try:
        pa = att.PropertyAccessor
        try:
            if pa.GetProperty(PR_ATTACHMENT_HIDDEN):
                return True
        except Exception:
            pass
        try:
            cid = pa.GetProperty(PR_ATTACH_CONTENT_ID)
        except Exception:
            cid = None
        name = (getattr(att, "FileName", "") or "").lower()
        return bool(cid) and name.endswith(IMAGE_EXT)
    except Exception:
        return False


def extract_online_links(html_body: str, text_body: str) -> List[Dict[str, str]]:
    links: Dict[str, str] = {}
    for href, label in ANCHOR.findall(html_body or ""):
        url = html.unescape(href).strip()
        if ONLINE_HOSTS.search(url):
            text = re.sub(r"<[^>]+>", "", label).strip()
            links.setdefault(url, text)
    for url in BARE_URL.findall(text_body or ""):
        url = url.rstrip(").,;>")
        if ONLINE_HOSTS.search(url):
            links.setdefault(url, "")
    out = []
    for url, label in links.items():
        if not label or label.startswith("http"):
            path = unquote(urlparse(url).path).rstrip("/")
            label = path.rsplit("/", 1)[-1] or urlparse(url).netloc
        out.append({"url": url, "label": label})
    return out


def write_url_shortcut(directory: str, label: str, url: str, prefix: str = "") -> Optional[str]:
    filename = safe_name(f"{prefix}{label}") + ".url"
    path = os.path.join(directory, filename)
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8", errors="ignore") as f:
                if url in f.read():
                    return None  # same shortcut already there
        except Exception:
            pass
        path = unique_path(directory, filename) or path
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"[InternetShortcut]\nURL={url}\n")
    return path


def save_attachments(email_number: int, target_dir: str, name_prefix: str = "",
                     include_inline: bool = False, create_link_shortcuts: bool = True,
                     only_names: Optional[List[str]] = None, dry_run: bool = False) -> Dict[str, Any]:
    if not os.path.isabs(target_dir):
        raise ValueError("target_dir must be an absolute Windows path")
    if not dry_run:
        os.makedirs(target_dir, exist_ok=True)
    wanted = {n.lower() for n in (only_names or [])}

    saved, skipped, links_out = [], [], []
    with OutlookSessionManager() as session:
        item = get_item_by_number(session, email_number)
        subject = getattr(item, "Subject", "")
        atts = item.Attachments
        for i in range(1, atts.Count + 1):
            att = atts.Item(i)
            fname = getattr(att, "FileName", "") or getattr(att, "DisplayName", "") or f"piece_{i}"
            if wanted and fname.lower() not in wanted:
                continue
            if not include_inline and _is_inline(att):
                skipped.append({"name": fname, "reason": "inline image"})
                continue
            if getattr(att, "Type", 1) == 6:  # olOLE / embedded object without file
                skipped.append({"name": fname, "reason": "embedded object"})
                continue
            size = getattr(att, "Size", None)
            target_name = safe_name(f"{name_prefix}{fname}")
            path = unique_path(target_dir, target_name, size) if os.path.isdir(target_dir) else os.path.join(target_dir, target_name)
            if path is None:
                skipped.append({"name": fname, "reason": "identical file already present"})
                continue
            if dry_run:
                saved.append({"name": fname, "path": path, "size": size})
                continue
            try:
                att.SaveAsFile(path)
                saved.append({"name": fname, "path": path, "size": size})
            except Exception as e:
                skipped.append({"name": fname, "reason": f"save failed: {e}"})

        if create_link_shortcuts:
            for link in extract_online_links(getattr(item, "HTMLBody", ""), getattr(item, "Body", "")):
                if dry_run:
                    links_out.append({**link, "path": os.path.join(target_dir, safe_name(name_prefix + link["label"]) + ".url")})
                    continue
                try:
                    p = write_url_shortcut(target_dir, link["label"], link["url"], name_prefix)
                    if p:
                        links_out.append({**link, "path": p})
                except Exception as e:
                    skipped.append({"name": link["label"], "reason": f"shortcut failed: {e}"})

    if not dry_run:
        log_event("attachments_saved", email_number=email_number, target=target_dir,
                  files=[s["path"] for s in saved], shortcuts=[l.get("path") for l in links_out])
    return {"subject": subject, "target_dir": target_dir, "dry_run": dry_run,
            "saved": saved, "shortcuts": links_out, "skipped": skipped}


def list_directory(path: str, depth: int = 2) -> Dict[str, Any]:
    """List sub-directories (and file counts) to choose where to file attachments."""
    if not os.path.isdir(path):
        raise ValueError(f"Not a directory: {path}")
    out = []
    base_depth = path.rstrip("\\/").count(os.sep)
    for root, dirs, files in os.walk(path):
        level = root.rstrip("\\/").count(os.sep) - base_depth
        dirs[:] = sorted(d for d in dirs if not d.startswith("."))
        if level >= depth:
            dirs[:] = []
        out.append({"dir": os.path.relpath(root, path), "files": len(files)})
        if len(out) > 500:
            break
    return {"root": path, "tree": out}
