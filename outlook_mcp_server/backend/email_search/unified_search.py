"""Single entry point for the field searches (kept for compatibility)."""

from typing import Any, Dict, List, Optional, Tuple

from .search_common import search_field

_FIELDS = {"subject": "subject", "sender": "sender", "recipient": "recipient", "body": "body"}


def unified_search(
    search_term: str, days: int = 7, folder_name: Optional[str] = None, match_all: bool = True,
    search_type: str = "subject",
) -> Tuple[List[Dict[str, Any]], str]:
    """Live search of one field. Words are separate terms; "quoted text" is one phrase.

    match_all=True requires every term, False any of them. Matching is a
    case-insensitive substring match done by Outlook on the whole folder.
    """
    if search_type not in _FIELDS:
        return [], f"Search type '{search_type}' is not supported"
    return search_field(_FIELDS[search_type], search_term, days, folder_name, match_all)
