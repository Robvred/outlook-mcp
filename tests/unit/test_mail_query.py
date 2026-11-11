"""Query engine, bulk move, stable numbering and tools, against a fake Outlook (any OS)."""

from datetime import datetime, timedelta

import pytest

from outlook_mcp_server.backend import bulk_move, shared
from outlook_mcp_server.backend.email_search import search_common
from outlook_mcp_server.backend.mail_query import (
    MailCriteria,
    build_filter,
    parse_search_terms,
    query_emails,
)
from outlook_mcp_server.backend.outlook_session import email_operations
from outlook_mcp_server.tools import search_tools, triage_tools, viewing_tools

from fakes import make_mailbox


@pytest.fixture
def box(monkeypatch):
    session, inbox, sub, archive, deleted = make_mailbox()
    for module in (bulk_move, search_common, email_operations, viewing_tools):
        monkeypatch.setattr(module, "OutlookSessionManager", session)
    monkeypatch.setattr(bulk_move, "log_event", lambda *a, **k: None)
    monkeypatch.setattr(email_operations, "log_event", lambda *a, **k: None)
    shared.clear_email_cache()
    yield session, inbox, sub, archive, deleted
    shared.clear_email_cache()


def ids(result):
    return [e["entry_id"] for e in result.emails]


# ------------------------------------------------------------------ engine

def test_rows_are_newest_first_and_carry_metadata(box):
    _, inbox, *_ = box
    result = query_emails(inbox, MailCriteria(days=30))
    assert ids(result) == ["A1", "A2", "A3", "A4"]
    first = result.emails[0]
    assert first["categories"] == "Action" and first["unread"] is True
    assert first["store_id"] == "STORE-user@company.com"
    assert first["folder"] == "user@company.com/Inbox"
    assert result.emails[1]["categories"] == "Done, Info"
    assert result.emails[1]["sender_email"] == "billing@example.com"
    assert result.emails[3]["flag_status"] == 2
    assert result.method == "table" and result.total_matched == 4 and not result.truncated


def test_max_results_keeps_the_most_recent_and_reports_total(box):
    _, inbox, *_ = box
    result = query_emails(inbox, MailCriteria(), max_results=2)
    assert ids(result) == ["A1", "A2"]
    assert result.total_matched == 5 and result.truncated


@pytest.mark.parametrize("setup", ["table", "no_sort", "no_array", "column_major", "no_table", "few_columns"])
def test_same_result_on_every_access_path(box, setup):
    _, inbox, *_ = box
    if setup == "no_sort":
        inbox.sort_ok = False
    elif setup == "no_array":
        inbox.array_mode = "off"
    elif setup == "column_major":
        inbox.array_mode = "columns"
    elif setup == "no_table":
        inbox.table_ok = False
    elif setup == "few_columns":
        inbox.blocked_columns = {"Size", "Importance", "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"}
    result = query_emails(inbox, MailCriteria(days=30, categories=["done"]), max_results=10)
    assert ids(result) == ["A2", "A3"]
    assert result.total_matched == 2
    assert result.method == ("items" if setup == "no_table" else "table")


def test_text_criteria(box):
    _, inbox, *_ = box
    assert ids(query_emails(inbox, MailCriteria(subject=["project a"]))) == ["A1", "A3"]
    assert ids(query_emails(inbox, MailCriteria(sender=["vendor"]))) == ["A2", "A5"]
    assert ids(query_emails(inbox, MailCriteria(sender=["example.com"]))) == ["A2"]
    assert ids(query_emails(inbox, MailCriteria(recipient=["doe"]))) == ["A3"]
    assert ids(query_emails(inbox, MailCriteria(body=["budget"]))) == ["A3"]
    assert ids(query_emails(inbox, MailCriteria(subject=["project", "minutes"], match_all_terms=True))) == ["A3"]
    assert ids(query_emails(inbox, MailCriteria(subject=["lunch", "invoice"]))) == ["A2", "A4"]


def test_state_and_date_criteria(box):
    _, inbox, *_ = box
    assert ids(query_emails(inbox, MailCriteria(unread_only=True))) == ["A1"]
    assert ids(query_emails(inbox, MailCriteria(flagged_only=True))) == ["A4"]
    assert ids(query_emails(inbox, MailCriteria(older_than_days=30))) == ["A5"]
    start = datetime.now() - timedelta(days=3, hours=1)
    end = datetime.now() - timedelta(days=1)
    assert ids(query_emails(inbox, MailCriteria(received_from=start, received_to=end))) == ["A2", "A3"]


def test_category_criteria(box):
    session, inbox, *_ = box
    assert ids(query_emails(inbox, MailCriteria(categories=["Done"]))) == ["A2", "A3", "A5"]
    assert ids(query_emails(inbox, MailCriteria(categories=["Done", "Action"]))) == ["A1", "A2", "A3", "A5"]
    both = MailCriteria(categories=["Done", "Info"], categories_match_all=True)
    assert ids(query_emails(inbox, both)) == ["A2"]
    assert ids(query_emails(inbox, MailCriteria(days=30, exclude_categories=["done"]))) == ["A1", "A4"]
    # "Do" must not match "Done": categories are whole names, not substrings.
    assert ids(query_emails(inbox, MailCriteria(categories=["Do"]))) == []
    unknown = query_emails(inbox, MailCriteria(categories=["Donne"]), namespace=session.namespace)
    assert unknown.emails == [] and "Donne" in unknown.warnings[0]


def test_excluded_categories_are_checked_on_rows_not_in_the_filter():
    flt = build_filter(MailCriteria(days=7, exclude_categories=["Action"]))
    assert "Keywords" not in flt and "datereceived" in flt


def test_subfolders(box):
    _, inbox, *_ = box
    result = query_emails(inbox, MailCriteria(subject=["project"]), include_subfolders=True)
    assert ids(result) == ["A1", "A3", "B1"]
    assert result.emails[2]["folder"] == "user@company.com/Inbox/ProjectA"


def test_criteria_validation_and_term_cleaning():
    assert MailCriteria(subject=["", "  ", " a "]).subject == ["a"]
    assert MailCriteria(subject=[""]).is_empty()
    with pytest.raises(ValueError):
        MailCriteria(days=5, older_than_days=10)
    with pytest.raises(ValueError):
        MailCriteria(days=-1)
    assert "''" in build_filter(MailCriteria(subject=["it's"]))
    assert parse_search_terms('project "status report"  a') == ["project", "status report", "a"]


# ------------------------------------------------------------------ bulk move

def test_dry_run_moves_nothing(box):
    _, inbox, _, archive, _ = box
    out = bulk_move.move_by_filter("Archive", days=30, categories=["Done"], dry_run=True)
    assert out["matched"] == 2 and out["moved"] == 0 and out["remaining"] == 2
    assert out["target"] == "user@company.com/Archive" and out["source"] == "user@company.com/Inbox"
    assert len(inbox.mails) == 5 and not archive.mails
    assert "Invoice INV-0001" in out["sample"][0]


def test_move_only_emails_carrying_the_categories(box):
    _, inbox, _, archive, _ = box
    out = bulk_move.move_by_filter("Archive", days=None, categories=["Done"])
    assert out["moved"] == 3 and out["remaining"] == 0 and out["failed"] == 0
    assert sorted(m.Subject for m in archive.mails) == ["Invoice INV-0001", "Old report", "Project A minutes"]
    assert sorted(m.Subject for m in inbox.mails) == ["Lunch", "Project A status"]
    again = bulk_move.move_by_filter("Archive", days=None, categories=["Done"])
    assert again["matched"] == 0 and again["moved"] == 0


def test_move_requires_all_categories_and_honours_exclusions(box):
    _, _, _, archive, _ = box
    out = bulk_move.move_by_filter("Archive", days=None, categories=["Done", "Info"], categories_match="all")
    assert [m.Subject for m in archive.mails] == ["Invoice INV-0001"] and out["moved"] == 1
    out = bulk_move.move_by_filter("Archive", days=30, exclude_categories=["Action"])
    assert out["moved"] == 2 and out["remaining"] == 0
    assert "Project A status" not in [m.Subject for m in archive.mails]


def test_move_guards(box):
    with pytest.raises(ValueError):
        bulk_move.move_by_filter("Archive", days=None)  # no criterion
    with pytest.raises(ValueError):
        bulk_move.move_by_filter("Archive", days=None, subject_contains=["", " "])  # empty terms only
    with pytest.raises(ValueError):
        bulk_move.move_by_filter("Inbox", days=7)  # same folder
    with pytest.raises(ValueError):
        bulk_move.move_by_filter("Archive", categories=["Done"], exclude_categories=["done"])
    with pytest.raises(ValueError):
        bulk_move.move_by_filter("Archive", categories=["Done"], categories_match="some")


def test_move_stops_on_time_budget_and_reports_remaining(box):
    _, inbox, *_ = box
    out = bulk_move.move_by_filter("Archive", days=None, categories=["Done"], time_budget_s=-1)
    assert out["moved"] == 0 and out["remaining"] == 3 and len(inbox.mails) == 5


def test_move_tool_passes_categories(box):
    _, _, _, archive, _ = box
    out = triage_tools.move_emails_by_filter_tool("Archive", days=30, categories="Done", dry_run=False)
    assert out["type"] == "json" and out["data"]["moved"] == 2
    assert len(archive.mails) == 2
    preview = triage_tools.move_emails_by_filter_tool("Archive", days=None, categories=["Done"])
    assert preview["data"]["dry_run"] is True and preview["data"]["matched"] == 1


# ------------------------------------------------------------------ listing and tools

def test_load_folder_exposes_categories(box):
    out = viewing_tools.load_emails_by_folder_tool("Inbox", days=7)
    assert "Found 4 emails" in out["text"]
    page = viewing_tools.view_email_cache_tool(1, page_size=10)["data"]
    assert [e["categories"] for e in page["emails"]] == ["Action", "Done, Info", "Done", ""]
    assert page["emails"][0]["status"] == "Unread" and page["emails"][3]["flag"] == "flagged"
    assert page["total_emails"] == 4 and page["total_pages"] == 1
    assert viewing_tools.view_email_cache_tool(1)["data"]["total_pages"] == 1


def test_load_folder_by_count_and_relative_path(box):
    out = viewing_tools.load_emails_by_folder_tool("Inbox", max_emails=2)
    assert "Found 2 emails" in out["text"]
    assert shared.email_cache_order == ["A1", "A2"]
    out = viewing_tools.load_emails_by_folder_tool("Inbox/ProjectA", max_emails=5)
    assert shared.email_cache_order == ["B1"]
    out = viewing_tools.load_emails_by_folder_tool("user@company.com/Inbox/projecta", days=30)
    assert shared.email_cache_order == ["B1"]
    out = viewing_tools.load_emails_by_folder_tool("Nowhere", days=7)
    assert out["text"].startswith("Error")


def test_search_tool_returns_matches_directly(box):
    out = search_tools.search_emails_tool(subject_contains="project", categories="Done", days=30)
    data = out["data"]
    assert data["total_matched"] == 1 and data["emails"][0]["subject"] == "Project A minutes"
    assert data["emails"][0]["number"] == 1 and shared.email_cache_order == ["A3"]
    empty = search_tools.search_emails_tool(subject_contains="nothing like this")
    assert empty["data"]["emails"] == [] and shared.email_cache_order == []
    sub = search_tools.search_emails_tool(subject_contains="kickoff", include_subfolders=True)
    assert sub["data"]["emails"][0]["folder"] == "user@company.com/Inbox/ProjectA"


def test_legacy_search_tools_use_and_or_terms(box):
    out = search_tools.search_email_by_subject_tool("project minutes", days=30)
    assert "Found 1 matching" in out["text"]
    out = search_tools.search_email_by_subject_tool("lunch invoice", days=30, match_all=False)
    assert "Found 2 matching" in out["text"]
    out = search_tools.search_email_by_body_tool("budget", days=30)
    assert "Found 1 matching" in out["text"]
    out = search_tools.search_email_by_sender_name_tool("vendor", days=30)
    assert shared.email_cache_order == ["A2"]


def test_numbers_stay_stable_after_a_move(box):
    _, inbox, _, archive, deleted = box
    viewing_tools.load_emails_by_folder_tool("Inbox", days=30)
    before = list(shared.email_cache_order)
    msg = email_operations.move_email_to_folder(2, "Archive")
    assert "moved to 'user@company.com/Archive'" in msg and "Invoice" in msg
    assert len(shared.email_cache_order) == len(before)
    assert shared.email_cache_order[2] == before[2]  # number 3 is still the same email
    assert shared.email_cache_order[1] == "A2M"  # number 2 follows the moved email
    assert email_operations.move_email_to_folder(3, "Archive").startswith("Email #3 'Project A minutes'")
    assert sorted(m.Subject for m in archive.mails) == ["Invoice INV-0001", "Project A minutes"]
    assert "already in" in email_operations.move_email_to_folder(2, "Archive")
    assert "Deleted Items" in email_operations.delete_email_by_number(1)
    assert [m.Subject for m in deleted.mails] == ["Project A status"]
    assert email_operations.move_email_to_folder(99, "Archive").startswith("Error")


def test_replace_cache_is_ordered_and_deduplicated():
    rows = [{"entry_id": "Y", "received_time": "2026-01-05 10:00:00"},
            {"entry_id": "X", "received_time": "2026-01-01 10:00:00"},
            {"entry_id": "Y", "received_time": "2026-01-05 10:00:00"}, {"subject": "no id"}]
    assert shared.replace_email_cache(rows, source="test") == 2
    assert shared.email_cache_order == ["Y", "X"] and shared.cache_meta["source"] == "test"
    shared.add_email_to_cache("Z", {"entry_id": "Z", "received_time": "2026-01-03 10:00:00"})
    assert shared.email_cache_order == ["Y", "Z", "X"]
    shared.clear_email_cache()
