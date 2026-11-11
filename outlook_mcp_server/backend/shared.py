"""Numbered listing of the last load or search (the "email cache").

The cache is only a handle table: position N -> EntryID (+ StoreID and the row
returned by the query engine). Mail data itself always comes live from Outlook.

Guarantees
- ``email_cache`` and ``email_cache_order`` are mutated in place, never rebound:
  other modules import them by name.
- A load replaces the whole listing atomically (``replace_email_cache``).
- Numbers stay stable until the next load: moving or deleting a message keeps
  its slot (``update_cache_entry``), so later numbers never shift.
"""

import bisect
import json
import os
import threading
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, Iterable, List, Optional, Union

from .config import cache_config, performance_config
from .logging_config import get_logger

logger = get_logger(__name__)

CACHE_BASE_DIR = cache_config.CACHE_BASE_DIR
CACHE_EXPIRY_HOURS = cache_config.CACHE_EXPIRY_HOURS
MAX_CACHE_SIZE = performance_config.MAX_CACHE_SIZE

email_cache: Dict[str, Dict[str, Any]] = {}
email_cache_order: List[str] = []
cache_meta: Dict[str, Any] = {}

_cache_lock = threading.RLock()
_email_time_cache: Dict[str, datetime] = {}
_LEGACY_FORMATS = ("%m/%d/%y %H:%M:%S", "%m/%d/%Y %H:%M:%S")


def _get_cache_file() -> str:
    return os.path.join(CACHE_BASE_DIR, "email_cache.json")


def _parse_email_time(received_time_str: str) -> datetime:
    """Parse a received time (ISO with or without 'T', or legacy US format). datetime.min if unknown."""
    cached = _email_time_cache.get(received_time_str)
    if cached is not None:
        return cached
    parsed = datetime.min
    if isinstance(received_time_str, str) and received_time_str:
        try:
            parsed = datetime.fromisoformat(received_time_str)
        except ValueError:
            for fmt in _LEGACY_FORMATS:
                try:
                    parsed = datetime.strptime(received_time_str, fmt).replace(tzinfo=timezone.utc)
                    break
                except ValueError:
                    continue
    _email_time_cache[received_time_str] = parsed
    return parsed


def _aware(value: datetime) -> datetime:
    return value if value.tzinfo else value.replace(tzinfo=timezone.utc)


def _time_key(email_data: Optional[Dict[str, Any]]) -> float:
    """Sort key, newest first (smaller = more recent). Unknown dates go last."""
    parsed = _parse_email_time((email_data or {}).get("received_time", ""))
    if parsed == datetime.min:
        return float("inf")
    try:
        return -_aware(parsed).timestamp()
    except (OverflowError, OSError, ValueError):
        return float("inf")


def _forget_time(email_data: Optional[Dict[str, Any]]) -> None:
    if email_data:
        _email_time_cache.pop(email_data.get("received_time", ""), None)


def add_email_to_cache(email_id: str, email_data: Dict[str, Any]) -> None:
    """Add one email, keeping the listing sorted newest first and within MAX_CACHE_SIZE."""
    with _cache_lock:
        if email_id in email_cache:
            try:
                email_cache_order.remove(email_id)
            except ValueError:
                pass
        email_cache[email_id] = email_data
        keys = [_time_key(email_cache.get(i)) for i in email_cache_order]
        email_cache_order.insert(bisect.bisect_right(keys, _time_key(email_data)), email_id)
        while len(email_cache_order) > MAX_CACHE_SIZE:
            _forget_time(email_cache.pop(email_cache_order.pop(), None))


def replace_email_cache(emails: Iterable[Dict[str, Any]], source: str = "", persist: bool = True) -> int:
    """Replace the listing with these emails, in the given order. Returns the number kept."""
    new_cache: Dict[str, Dict[str, Any]] = {}
    new_order: List[str] = []
    for email_data in emails:
        entry_id = (email_data or {}).get("entry_id")
        if not entry_id or entry_id in new_cache:
            continue
        new_cache[entry_id] = email_data
        new_order.append(entry_id)
        if len(new_order) >= MAX_CACHE_SIZE:
            break
    with _cache_lock:
        email_cache.clear()
        email_cache.update(new_cache)
        email_cache_order[:] = new_order
        _email_time_cache.clear()
        cache_meta.clear()
        cache_meta.update({"source": source, "loaded_at": datetime.now().isoformat(timespec="seconds")})
        if persist:
            immediate_save_cache()
    return len(new_order)


def update_cache_entry(old_entry_id: str, new_entry_id: Optional[str] = None, **fields: Any) -> None:
    """Update an entry in place (same number). Used after a move, which changes the EntryID."""
    with _cache_lock:
        data = email_cache.get(old_entry_id)
        if data is None:
            return
        data.update(fields)
        if new_entry_id and new_entry_id != old_entry_id:
            data["entry_id"] = new_entry_id
            if "id" in data:
                data["id"] = new_entry_id
            position = email_cache_order.index(old_entry_id)
            if new_entry_id in email_cache:  # already listed elsewhere: keep this slot only
                email_cache_order.remove(new_entry_id)
                position = email_cache_order.index(old_entry_id)
            del email_cache[old_entry_id]
            email_cache[new_entry_id] = data
            email_cache_order[position] = new_entry_id
        immediate_save_cache()


def clear_email_cache() -> None:
    """Clear the listing in memory and on disk."""
    with _cache_lock:
        email_cache.clear()
        email_cache_order.clear()
        _email_time_cache.clear()
        cache_meta.clear()
        try:
            cache_file = _get_cache_file()
            if os.path.exists(cache_file):
                os.remove(cache_file)
        except Exception as e:
            logger.warning(f"Failed to clear email cache from disk: {e}")


def clear_cache() -> None:
    """Deprecated alias of clear_email_cache."""
    clear_email_cache()


def get_cache_size() -> int:
    return len(email_cache_order)


def get_cache_stats() -> Dict[str, Any]:
    oldest = newest = None
    if email_cache_order:
        newest = email_cache.get(email_cache_order[0], {}).get("received_time", "")
        oldest = email_cache.get(email_cache_order[-1], {}).get("received_time", "")
    return {
        "total_emails": len(email_cache),
        "cache_size": len(email_cache),
        "cache_order_size": len(email_cache_order),
        "time_cache_size": len(_email_time_cache),
        "oldest_email": oldest,
        "newest_email": newest,
    }


def cleanup_cache() -> None:
    """Drop entries whose received time is older than CACHE_EXPIRY_HOURS."""
    threshold = datetime.now(timezone.utc) - timedelta(hours=CACHE_EXPIRY_HOURS)
    with _cache_lock:
        expired = []
        for email_id in email_cache_order:
            parsed = _parse_email_time(email_cache.get(email_id, {}).get("received_time", ""))
            if parsed != datetime.min and _aware(parsed) < threshold:
                expired.append(email_id)
        for email_id in expired:
            email_cache_order.remove(email_id)
            _forget_time(email_cache.pop(email_id, None))
    if expired:
        logger.info(f"Cleaned up {len(expired)} expired emails from cache")


def _as_datetime(value: Union[datetime, str]) -> Optional[datetime]:
    if isinstance(value, str):
        try:
            value = datetime.fromisoformat(value)
        except ValueError:
            return None
    return _aware(value)


def get_emails_by_date_range(start_date: Union[datetime, str], end_date: Union[datetime, str]) -> List[Dict[str, Any]]:
    """Cached emails received between two dates (10 s tolerance on both ends)."""
    start, end = _as_datetime(start_date), _as_datetime(end_date)
    if start is None or end is None:
        return []
    start, end = start - timedelta(seconds=10), end + timedelta(seconds=10)
    result = []
    for email_id in list(email_cache_order):
        data = email_cache.get(email_id, {})
        parsed = _parse_email_time(data.get("received_time", ""))
        if parsed != datetime.min and start <= _aware(parsed) <= end:
            result.append(data)
    return result


def _filter_cached(fields: Iterable[str], needle: str) -> List[Dict[str, Any]]:
    if not needle:
        return []
    needle = needle.lower()
    result = []
    for email_id in list(email_cache_order):
        data = email_cache.get(email_id, {})
        if any(needle in str(data.get(f) or "").lower() for f in fields):
            result.append(data)
    return result


def get_emails_by_sender(sender: str) -> List[Dict[str, Any]]:
    return _filter_cached(("from", "sender", "sender_email"), sender)


def get_emails_by_subject(subject: str) -> List[Dict[str, Any]]:
    return _filter_cached(("subject",), subject)


# Deprecated aliases kept for compatibility.
get_emails_by_date_range_cached = get_emails_by_date_range
get_emails_by_sender_cached = get_emails_by_sender
get_emails_by_subject_cached = get_emails_by_subject


def immediate_save_cache() -> None:
    """Write the listing to disk (atomic replace). Failures are logged, never raised."""
    with _cache_lock:
        if not email_cache_order:
            return
        try:
            os.makedirs(CACHE_BASE_DIR, exist_ok=True)
            cache_file = _get_cache_file()
            temp_file = cache_file + ".tmp"
            with open(temp_file, "w", encoding="utf-8") as f:
                json.dump({
                    "cache": email_cache,
                    "cache_order": email_cache_order,
                    "meta": cache_meta,
                    "timestamp": datetime.now().isoformat(),
                }, f, ensure_ascii=False, default=str)
            os.replace(temp_file, cache_file)
        except Exception as e:
            logger.warning(f"Failed to save email cache: {e}")


def save_email_cache(force_save: bool = False) -> None:
    """Deprecated: the cache is saved on every change. Kept for compatibility."""
    immediate_save_cache()


def refresh_email_cache_with_new_data() -> bool:
    """Deprecated: clear the listing before a load. Use replace_email_cache instead."""
    clear_email_cache()
    return True


def load_email_cache() -> None:
    """Load the last listing from disk if it is not expired."""
    try:
        cache_file = _get_cache_file()
        if not os.path.exists(cache_file):
            return
        with open(cache_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        stamp = datetime.fromisoformat(data.get("timestamp", "2000-01-01T00:00:00"))
        if datetime.now() - stamp > timedelta(hours=CACHE_EXPIRY_HOURS):
            return
        cache = data.get("cache")
        if not isinstance(cache, dict):
            return
        order = data.get("cache_order")
        order = [i for i in order if i in cache] if isinstance(order, list) else list(cache)
        with _cache_lock:
            email_cache.clear()
            email_cache.update({i: cache[i] for i in order})
            email_cache_order[:] = order
            cache_meta.clear()
            if isinstance(data.get("meta"), dict):
                cache_meta.update(data["meta"])
        logger.info(f"Loaded {len(order)} emails from persistent cache")
    except Exception as e:
        logger.warning(f"Failed to load email cache: {e}")


def get_email_from_cache(email_identifier: Union[int, str]) -> Optional[Dict[str, Any]]:
    """Email data by 1-based number (ValueError if out of range) or by EntryID."""
    if isinstance(email_identifier, str):
        return email_cache.get(email_identifier)
    if email_identifier < 1 or email_identifier > len(email_cache_order):
        raise ValueError(
            f"Email number {email_identifier} is out of range. Available range: 1-{len(email_cache_order)}"
        )
    return email_cache.get(email_cache_order[email_identifier - 1])


load_email_cache()
