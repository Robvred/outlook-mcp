"""Forward drafts and Cc clearing on reply drafts, with a fake Outlook COM layer.

Runs on any OS: the Windows COM modules are stubbed before the package is imported.
"""

import sys
import types
from unittest import mock

import pytest

for _name in ("pythoncom", "pywintypes", "win32com", "win32com.client", "win32api", "win32con"):
    sys.modules.setdefault(_name, mock.MagicMock(name=_name))

from outlook_mcp_server.backend import email_composition as ec  # noqa: E402
from outlook_mcp_server.backend.com_helpers import OL_CC, OL_TO  # noqa: E402
from outlook_mcp_server.backend.validators import EmailForwardParams, EmailReplyParams  # noqa: E402


class FakeRecipient:
    def __init__(self, name, rtype=OL_TO):
        self.Name = name
        self.Type = rtype
        self.Resolved = True
        self.Address = name if "@" in name else ""
        self.AddressEntry = None


class FakeRecipients:
    def __init__(self, initial=None):
        self.items = list(initial or [])

    @property
    def Count(self):
        return len(self.items)

    def Item(self, i):
        return self.items[i - 1]

    def Remove(self, i):
        del self.items[i - 1]

    def Add(self, name):
        r = FakeRecipient(name)
        self.items.append(r)
        return r

    def ResolveAll(self):
        return True


class FakeAttachments:
    def __init__(self, n):
        self.items = [object() for _ in range(n)]

    @property
    def Count(self):
        return len(self.items)

    def Remove(self, i):
        del self.items[i - 1]


class FakeDraft:
    def __init__(self, recipients=None, attachments=0):
        self.Recipients = FakeRecipients(recipients)
        self.Attachments = FakeAttachments(attachments)
        self.HTMLBody = "<html><body><p>original</p></body></html>"
        self.EntryID = "DRAFT-0001"
        self.saved = False
        self.sent = False

    @property
    def GetInspector(self):
        return None

    def Save(self):
        self.saved = True

    def Send(self):  # must never be called
        self.sent = True


class FakeMail:
    def __init__(self, draft):
        self.draft = draft

    def Forward(self):
        return self.draft

    def ReplyAll(self):
        return self.draft


@pytest.fixture
def outlook(monkeypatch):
    holder = {}

    class FakeSession:
        def __enter__(self):
            ns = types.SimpleNamespace(GetItemFromID=lambda _id: holder["mail"])
            return types.SimpleNamespace(namespace=ns)

        def __exit__(self, *a):
            return False

    monkeypatch.setattr(ec, "OutlookSessionManager", FakeSession)
    monkeypatch.setattr(ec, "email_cache_order", ["ID-1"])
    monkeypatch.setattr(ec, "email_cache", {"ID-1": {"id": "ID-1"}})
    monkeypatch.setattr(ec, "log_event", lambda *a, **k: None)
    return holder


def _names(draft, rtype):
    return [r.Name for r in draft.Recipients.items if r.Type == rtype]


def test_forward_sets_only_given_recipients_and_keeps_attachments(outlook):
    draft = FakeDraft(attachments=2)
    outlook["mail"] = FakeMail(draft)
    out = ec.forward_email_by_number(1, "DOE John", "user@company.com", forward_text="See below.")
    assert draft.saved and not draft.sent
    assert _names(draft, OL_TO) == ["DOE John"]
    assert _names(draft, OL_CC) == ["user@company.com"]
    assert "See below." in draft.HTMLBody
    assert draft.HTMLBody.index("See below.") < draft.HTMLBody.index("original")
    assert "2 attachment(s)" in out


def test_forward_can_drop_attachments(outlook):
    draft = FakeDraft(attachments=3)
    outlook["mail"] = FakeMail(draft)
    out = ec.forward_email_by_number(1, ["Jane Smith"], include_attachments=False)
    assert draft.Attachments.Count == 0
    assert "0 attachment(s)" in out


def test_forward_requires_a_to_recipient():
    with pytest.raises(Exception):
        EmailForwardParams(email_number=1, to_recipients=[])


def test_forward_splits_semicolon_strings():
    p = EmailForwardParams(email_number=1, to_recipients="a@example.com; b@example.com ;")
    assert p.to_recipients == ["a@example.com", "b@example.com"]


@pytest.mark.parametrize("value", [[], "", "  ", [""]])
def test_reply_empty_cc_means_clear(value):
    assert EmailReplyParams(email_number=1, reply_text="x", cc_recipients=value).cc_recipients == []


def test_reply_none_cc_keeps_reply_all():
    assert EmailReplyParams(email_number=1, reply_text="x").cc_recipients is None


def test_reply_with_empty_cc_removes_original_cc(outlook):
    draft = FakeDraft(
        recipients=[FakeRecipient("Jane Smith", OL_TO), FakeRecipient("Vendor A contact", OL_CC)]
    )
    outlook["mail"] = FakeMail(draft)
    ec.reply_to_email_by_number(1, "Hello", to_recipients="DOE John", cc_recipients=[])
    assert _names(draft, OL_TO) == ["DOE John"]
    assert _names(draft, OL_CC) == []
    assert draft.saved and not draft.sent
