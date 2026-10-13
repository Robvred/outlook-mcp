"""Calendar listing with a fake Outlook COM layer (runs on any OS)."""

import sys
from datetime import datetime
from unittest import mock

for _name in ("pythoncom", "pywintypes", "win32com", "win32com.client", "win32api", "win32con"):
    sys.modules.setdefault(_name, mock.MagicMock(name=_name))

from outlook_mcp_server.backend import calendar_operations as cal  # noqa: E402


class FakeRecipient:
    def __init__(self, name, rtype):
        self.Name, self.Type, self.Address, self.AddressEntry = name, rtype, "", None


class FakeRecipients:
    def __init__(self, items):
        self.items = items

    @property
    def Count(self):
        return len(self.items)

    def Item(self, i):
        return self.items[i - 1]


class FakeAppt:
    def __init__(self, subject, start, end, status=1, response=3, body=""):
        self.Subject, self.Start, self.End = subject, start, end
        self.MeetingStatus, self.ResponseStatus, self.Body = status, response, body
        self.EntryID = subject
        self.Recipients = FakeRecipients([FakeRecipient("DOE John", 1), FakeRecipient("Jane Smith", 2)])


class FakeItems:
    def __init__(self, appts, restrict_ok=True):
        self.appts, self.restrict_ok, self.i = sorted(appts, key=lambda a: a.Start), restrict_ok, 0
        self.IncludeRecurrences = False

    def Sort(self, *_):
        pass

    def Restrict(self, flt):
        if not self.restrict_ok:
            raise RuntimeError("bad date format")
        # Simulates Outlook applying the window filter (fixture window: 2026-01-05).
        return FakeItems([a for a in self.appts if a.Start.day == 5])

    def GetFirst(self):
        self.i = 0
        return self.GetNext()

    def GetNext(self):
        if self.i >= len(self.appts):
            return None
        self.i += 1
        return self.appts[self.i - 1]


def _run(appts, restrict_ok=True, **kw):
    folder = mock.MagicMock()
    folder.Items = FakeItems(appts, restrict_ok)
    session = mock.MagicMock()
    session.namespace.GetDefaultFolder.return_value = folder
    with mock.patch.object(cal, "OutlookSessionManager") as m:
        m.return_value.__enter__.return_value = session
        return cal.list_calendar_events(**kw)


APPTS = [
    FakeAppt("Before", datetime(2026, 1, 4, 9), datetime(2026, 1, 4, 10)),
    FakeAppt("Weekly sync", datetime(2026, 1, 5, 9), datetime(2026, 1, 5, 10),
             body="Agenda: point A\n\n________________________________\nMicrosoft Teams https://teams.microsoft.com/l/x"),
    FakeAppt("Declined", datetime(2026, 1, 5, 11), datetime(2026, 1, 5, 12), response=4),
    FakeAppt("Canceled", datetime(2026, 1, 5, 14), datetime(2026, 1, 5, 15), status=5),
    FakeAppt("After", datetime(2026, 1, 6, 9), datetime(2026, 1, 6, 10)),
]


def test_day_window_filters_and_cleans():
    for ok in (True, False):
        res = _run(APPTS, restrict_ok=ok, start_date="2026-01-05")
        assert [e["subject"] for e in res["events"]] == ["Weekly sync"]
        ev = res["events"][0]
        assert ev["agenda"] == "Agenda: point A"
        assert ev["online_link_in_body"] is True
        assert ev["attendees"]["required"] == ["DOE John"] and ev["attendees"]["optional"] == ["Jane Smith"]
        assert res["diagnostics"]["method"].startswith("restrict" if ok else "scan")


def test_include_declined_and_canceled():
    res = _run(APPTS, start_date="2026-01-05", include_declined=True, include_canceled=True)
    assert [e["subject"] for e in res["events"]] == ["Weekly sync", "Declined", "Canceled"]


def test_end_date_inclusive_day():
    res = _run(APPTS, restrict_ok=False, start_date="2026-01-05", end_date="2026-01-06")
    assert [e["subject"] for e in res["events"]][-1] == "After"
