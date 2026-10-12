"""Read-only mailbox statistics over long periods, using the Outlook Table API.

The Table API reads only the requested columns, so thousands of items are
scanned in seconds without loading each MailItem (and without touching the
email cache used by the other tools).
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

from .logging_config import get_logger
from .outlook_session.session_manager import OutlookSessionManager

logger = get_logger(__name__)

_PREFIX = re.compile(r"^\s*((re|tr|fw|fwd|réf|ref|aw|wg|rv|sv)\s*:\s*|\[ext(ernal)?\]\s*)+", re.IGNORECASE)
_AUTO = re.compile(
    r"no-?reply|noreply|notification|do-not-reply|donotreply|mailer|automat|sharepoint|"
    r"servicenow|service-now|jira|confluence|workday|sailpoint|ariba|teams|yammer|newsletter|"
    r"microsoft|copilot|seatable|airtable",
    re.IGNORECASE,
)

COLUMNS = [
    "Subject",
    "SenderName",
    "ConversationTopic",
    "ReceivedTime",
    "Categories",
    "MessageClass",
    "urn:schemas:httpmail:displayto",
    "urn:schemas:httpmail:displaycc",
]


def _norm_topic(topic: str) -> str:
    t = _PREFIX.sub("", topic or "").strip()
    return re.sub(r"\s+", " ", t)[:120] or "(no subject)"


def _contains_me(field: str, me: str) -> bool:
    return bool(field and me and me.lower() in field.lower())


def mailbox_stats(folder_path: Optional[str] = None, days: int = 90, top_topics: int = 150,
                  top_senders: int = 60, include_subfolders: bool = False) -> Dict[str, Any]:
    """Aggregate sender / conversation statistics for a folder over N days."""
    days = max(1, min(int(days), 730))
    since = datetime.now(timezone.utc) - timedelta(days=days)
    dasl = f"@SQL=\"urn:schemas:httpmail:datereceived\" >= '{since.strftime('%Y-%m-%d %H:%M')}'"

    total = 0
    by_month: Counter = Counter()
    addressing = Counter()
    senders: Dict[str, Counter] = defaultdict(Counter)
    topics: Dict[str, Dict[str, Any]] = {}
    categories: Counter = Counter()

    with OutlookSessionManager() as session:
        me = ""
        try:
            me = session.namespace.CurrentUser.Name
        except Exception:
            pass
        root = session.get_folder(folder_path) if folder_path else session.namespace.GetDefaultFolder(6)
        if root is None:
            raise ValueError(f"Folder '{folder_path}' not found")

        folders = [root]
        if include_subfolders:
            stack = [root]
            while stack:
                f = stack.pop()
                for i in range(1, f.Folders.Count + 1):
                    sub = f.Folders.Item(i)
                    folders.append(sub)
                    stack.append(sub)

        for folder in folders:
            try:
                table = folder.GetTable(dasl)
                table.Columns.RemoveAll()
                for c in COLUMNS:
                    table.Columns.Add(c)
            except Exception as e:
                logger.debug(f"GetTable failed on {getattr(folder, 'Name', '?')}: {e}")
                continue
            while not table.EndOfTable:
                row = table.GetNextRow()
                try:
                    subj, sender, topic, received, cats, mclass, disp_to, disp_cc = row.GetValues()
                except Exception:
                    continue
                if mclass and not str(mclass).startswith("IPM.Note"):
                    continue  # skip meeting requests, receipts, etc.
                total += 1
                try:
                    by_month[received.strftime("%Y-%m")] += 1
                except Exception:
                    pass
                if _contains_me(disp_to, me):
                    role = "to"
                elif _contains_me(disp_cc, me):
                    role = "cc"
                else:
                    role = "other"  # distribution list, BCC, sent folder...
                addressing[role] += 1
                sender = sender or "(unknown)"
                senders[sender][role] += 1
                key = _norm_topic(topic or subj)
                t = topics.setdefault(key, {"count": 0, "to_me": 0, "last": "", "senders": Counter()})
                t["count"] += 1
                t["to_me"] += role == "to"
                t["senders"][sender] += 1
                try:
                    r = received.strftime("%Y-%m-%d")
                    t["last"] = max(t["last"], r)
                except Exception:
                    pass
                for c in re.split(r"[;,]", cats or ""):
                    if c.strip():
                        categories[c.strip()] += 1

    sender_rows = sorted(senders.items(), key=lambda kv: -sum(kv[1].values()))[:top_senders]
    topic_rows = sorted(topics.items(), key=lambda kv: -kv[1]["count"])[:top_topics]
    return {
        "folder": folder_path or "Inbox",
        "days": days,
        "me": me,
        "total": total,
        "by_month": dict(sorted(by_month.items())),
        "addressing": dict(addressing),
        "automated_share": round(
            sum(sum(c.values()) for s, c in senders.items() if _AUTO.search(s)) / total, 3
        ) if total else 0,
        "categories": dict(categories.most_common(20)),
        "top_senders": [
            {"sender": s, "count": sum(c.values()), "to": c["to"], "cc": c["cc"], "other": c["other"],
             "automated": bool(_AUTO.search(s))}
            for s, c in sender_rows
        ],
        "top_topics": [
            {"topic": k, "count": v["count"], "to_me": v["to_me"], "last": v["last"],
             "main_senders": [s for s, _ in v["senders"].most_common(3)]}
            for k, v in topic_rows
        ],
    }
