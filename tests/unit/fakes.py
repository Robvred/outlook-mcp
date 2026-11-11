"""Minimal fake of the Outlook COM objects used by the query engine (runs on any OS).

The fake evaluates the DASL filters produced by mail_query, so tests exercise the
filter text itself, not only the Python-side checks.
"""

import re
from datetime import datetime, timedelta, timezone

PROPS = {
    "urn:schemas:httpmail:subject": lambda m: m.Subject,
    "urn:schemas:httpmail:fromname": lambda m: m.SenderName,
    "urn:schemas:httpmail:fromemail": lambda m: m.SenderEmailAddress,
    "http://schemas.microsoft.com/mapi/proptag/0x5D01001F": lambda m: m.smtp,
    "urn:schemas:httpmail:displayto": lambda m: m.To,
    "urn:schemas:httpmail:displaycc": lambda m: m.CC,
    "urn:schemas:httpmail:textdescription": lambda m: m.Body,
}
COLUMNS = {
    "EntryID": lambda m: m.EntryID,
    "Subject": lambda m: m.Subject,
    "SenderName": lambda m: m.SenderName,
    "ReceivedTime": lambda m: m.ReceivedTime,
    "Categories": lambda m: m.Categories,
    "MessageClass": lambda m: m.MessageClass,
    "urn:schemas:httpmail:displayto": lambda m: m.To,
    "urn:schemas:httpmail:displaycc": lambda m: m.CC,
    "http://schemas.microsoft.com/mapi/proptag/0x5D01001F": lambda m: m.smtp,
    "SenderEmailAddress": lambda m: m.SenderEmailAddress,
    "http://schemas.microsoft.com/mapi/proptag/0x0E070003": lambda m: 0 if m.UnRead else 1,
    "http://schemas.microsoft.com/mapi/proptag/0x10900003": lambda m: m.FlagStatus,
    "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B": lambda m: bool(m.Attachments.Count),
    "Size": lambda m: 1234,
    "Importance": lambda m: 1,
    "LastModificationTime": lambda m: m.ReceivedTime,
}


def _cats(mail):
    return [c.strip().lower() for c in re.split(r"[;,]", mail.Categories or "") if c.strip()]


def _atom(expr, mail):
    expr = expr.strip()
    m = re.match(r'^"([^"]+)"\s+LIKE\s+\'%(.*)%\'$', expr)
    if m:
        return m.group(2).replace("''", "'").lower() in (PROPS[m.group(1)](mail) or "").lower()
    m = re.match(r'^"urn:schemas-microsoft-com:office:office#Keywords"\s+=\s+\'(.*)\'$', expr)
    if m:
        return m.group(1).replace("''", "'").lower() in _cats(mail)
    m = re.match(r'^"urn:schemas:httpmail:datereceived"\s+(>=|<)\s+\'(.*)\'$', expr)
    if m:
        bound = datetime.strptime(m.group(2), "%Y-%m-%d %H:%M").replace(tzinfo=timezone.utc)
        received = mail.ReceivedTime.astimezone().astimezone(timezone.utc)
        return received >= bound if m.group(1) == ">=" else received < bound
    if expr == '"urn:schemas:httpmail:read" = 0':
        return mail.UnRead
    if expr == '"http://schemas.microsoft.com/mapi/proptag/0x0E2B0003" = 1':
        return mail.FlagStatus in (1, 2)
    raise ValueError(f"Cannot parse condition: {expr}")


def _split(expr, sep):
    parts, depth, quote, last, i = [], 0, False, 0, 0
    while i < len(expr):
        c = expr[i]
        if c == "'":
            quote = not quote
        elif not quote and c == "(":
            depth += 1
        elif not quote and c == ")":
            depth -= 1
        elif not quote and depth == 0 and expr.startswith(sep, i):
            parts.append(expr[last:i])
            last = i + len(sep)
            i += len(sep) - 1
        i += 1
    parts.append(expr[last:])
    return parts


def _wrapped(expr):
    if not (expr.startswith("(") and expr.endswith(")")):
        return False
    depth, quote = 0, False
    for i, c in enumerate(expr):
        if c == "'":
            quote = not quote
        elif not quote and c == "(":
            depth += 1
        elif not quote and c == ")":
            depth -= 1
            if depth == 0 and i < len(expr) - 1:
                return False
    return True


def evaluate(expr, mail):
    expr = expr.strip()
    while _wrapped(expr):
        expr = expr[1:-1].strip()
    for sep, combine in ((" OR ", any), (" AND ", all)):
        parts = _split(expr, sep)
        if len(parts) > 1:
            return combine(evaluate(p, mail) for p in parts)
    return _atom(expr, mail)


def matches(flt, mail):
    if not flt:
        return True
    assert flt.startswith("@SQL=")
    return evaluate(flt[5:], mail)


class FakeAttachments:
    def __init__(self, count=0):
        self.Count = count


class FakeMail:
    def __init__(self, entry_id, subject, sender="DOE John", received=None, categories="",
                 unread=False, flag=0, to="Jane Smith", cc="", body="", smtp="user@company.com",
                 attachments=0):
        self.EntryID = entry_id
        self.Subject = subject
        self.SenderName = sender
        self.SenderEmailAddress = "/O=ORG/CN=RECIPIENTS/CN=X"
        self.smtp = smtp
        self.ReceivedTime = received or datetime.now()
        self.LastModificationTime = self.ReceivedTime
        self.Categories = categories
        self.UnRead = unread
        self.FlagStatus = flag
        self.To, self.CC, self.Body = to, cc, body
        self.MessageClass = "IPM.Note"
        self.Size, self.Importance = 1234, 1
        self.Attachments = FakeAttachments(attachments)
        self.Parent = None
        self.moves = 0

    def Move(self, target):
        self.Parent.mails.remove(self)
        target.mails.append(self)
        self.Parent = target
        self.moves += 1
        self.EntryID = self.EntryID + "M"  # like Exchange, a move changes the EntryID
        target.store.index[self.EntryID] = self
        return self


class FakeRow:
    def __init__(self, values):
        self.values = values

    def GetValues(self):
        return self.values


class FakeColumns:
    def __init__(self, table):
        self.table = table

    def RemoveAll(self):
        self.table.columns = []

    def Add(self, name):
        if name not in COLUMNS or name in self.table.folder.blocked_columns:
            raise RuntimeError(f"property {name} is not supported")
        self.table.columns.append(name)


class FakeTable:
    def __init__(self, folder, mails):
        self.folder, self.mails, self.columns, self.pos = folder, list(mails), [], 0
        self.Columns = FakeColumns(self)

    def Sort(self, prop, descending):
        if not self.folder.sort_ok:
            raise RuntimeError("sort not supported")
        assert prop == "[ReceivedTime]"
        self.mails.sort(key=lambda m: m.ReceivedTime, reverse=bool(descending))
        self.pos = 0

    def GetRowCount(self):
        return len(self.mails)

    @property
    def EndOfTable(self):
        return self.pos >= len(self.mails)

    def _values(self, mail):
        return tuple(COLUMNS[c](mail) for c in self.columns)

    def GetNextRow(self):
        self.pos += 1
        return FakeRow(self._values(self.mails[self.pos - 1]))

    def GetArray(self, n):
        if self.folder.array_mode == "off":
            raise RuntimeError("GetArray not available")
        chunk = self.mails[self.pos:self.pos + n]
        self.pos += len(chunk)
        rows = tuple(self._values(m) for m in chunk)
        if self.folder.array_mode == "columns":
            return tuple(zip(*rows))
        return rows


class FakeItems:
    def __init__(self, folder, mails):
        self.folder, self.mails, self.pos = folder, list(mails), 0

    @property
    def Count(self):
        return len(self.mails)

    def Restrict(self, flt):
        return FakeItems(self.folder, [m for m in self.mails if matches(flt, m)])

    def Sort(self, prop, descending):
        self.mails.sort(key=lambda m: m.ReceivedTime, reverse=bool(descending))

    def GetFirst(self):
        self.pos = 0
        return self.GetNext()

    def GetNext(self):
        if self.pos >= len(self.mails):
            return None
        self.pos += 1
        return self.mails[self.pos - 1]


class FakeFolders:
    def __init__(self, folders):
        self.folders = folders

    @property
    def Count(self):
        return len(self.folders)

    def Item(self, i):
        return self.folders[i - 1]

    def __iter__(self):
        return iter(self.folders)

    def __getitem__(self, name):
        for f in self.folders:
            if f.Name.lower() == name.lower():
                return f
        raise KeyError(name)


class FakeStore:
    def __init__(self, name="user@company.com"):
        self.name, self.index = name, {}
        self.StoreID = "STORE-" + name


class FakeFolder:
    def __init__(self, name, store, parent=None, mails=()):
        self.Name, self.store, self.Parent = name, store, parent
        self.EntryID = f"F-{id(self)}"
        self.StoreID = store.StoreID
        self.mails, self.children = [], []
        self.table_ok, self.sort_ok, self.array_mode = True, True, "rows"
        self.blocked_columns = set()
        self.table_filters = []
        for m in mails:
            self.add(m)
        if parent is not None:
            parent.children.append(self)

    def add(self, mail):
        mail.Parent = self
        self.mails.append(mail)
        self.store.index[mail.EntryID] = mail
        return mail

    @property
    def FolderPath(self):
        if self.Parent is None:
            return "\\\\" + self.Name
        return self.Parent.FolderPath + "\\" + self.Name

    @property
    def Folders(self):
        return FakeFolders(self.children)

    @property
    def Items(self):
        return FakeItems(self, self.mails)

    def GetTable(self, flt=""):
        self.table_filters.append(flt)
        if not self.table_ok:
            raise RuntimeError("GetTable not supported")
        return FakeTable(self, [m for m in self.mails if matches(flt, m)])


class FakeCategory:
    def __init__(self, name):
        self.Name = name


class FakeCategories:
    def __init__(self, names):
        self.items = [FakeCategory(n) for n in names]

    @property
    def Count(self):
        return len(self.items)

    def Item(self, i):
        return self.items[i - 1]


class FakeNamespace:
    def __init__(self, store, root, inbox, categories=()):
        self.store, self.root, self.inbox = store, root, inbox
        self.Categories = FakeCategories(categories)
        self.Folders = FakeFolders([root])

    def GetDefaultFolder(self, folder_type):
        if folder_type == 6:
            return self.inbox
        for f in self.root.children:
            if getattr(f, "default_type", None) == folder_type:
                return f
        raise RuntimeError("no such default folder")

    def GetItemFromID(self, entry_id, store_id=None):
        if entry_id not in self.store.index:
            raise RuntimeError("item not found")
        return self.store.index[entry_id]


class FakeSession:
    """Stands in for OutlookSessionManager (context manager included)."""

    def __init__(self, namespace):
        self.namespace = self.outlook_namespace = namespace
        from outlook_mcp_server.backend.outlook_session.folder_operations import FolderOperations
        self._ops = FolderOperations(self)

    def __call__(self):
        return self

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def get_folder(self, name=None):
        return self._ops.get_folder(name)

    def get_folder_emails(self, folder_name="Inbox", max_emails=100, fast_mode=True, days_filter=None):
        return self._ops.get_folder_emails(folder_name, max_emails, fast_mode, days_filter)


def make_mailbox(categories=("Action", "Done", "Info")):
    """Mailbox with Inbox (+ subfolder ProjectA), Archive and Deleted Items."""
    now = datetime.now().replace(microsecond=0)
    store = FakeStore()
    root = FakeFolder("user@company.com", store)
    inbox = FakeFolder("Inbox", store, root, [
        FakeMail("A1", "Project A status", received=now - timedelta(hours=1), categories="Action", unread=True),
        FakeMail("A2", "Invoice INV-0001", sender="Vendor A", received=now - timedelta(days=2),
                 categories="Done; Info", smtp="billing@example.com", attachments=1),
        FakeMail("A3", "Project A minutes", received=now - timedelta(days=3), categories="Done",
                 to="Jane Smith; DOE John", body="budget review"),
        FakeMail("A4", "Lunch", sender="Jane Smith", received=now - timedelta(days=5), flag=2),
        FakeMail("A5", "Old report", sender="Vendor B", received=now - timedelta(days=40), categories="Done"),
    ])
    sub = FakeFolder("ProjectA", store, inbox, [
        FakeMail("B1", "Project A kickoff", received=now - timedelta(days=4), categories="Info"),
    ])
    archive = FakeFolder("Archive", store, root)
    deleted = FakeFolder("Deleted Items", store, root)
    deleted.default_type = 3
    namespace = FakeNamespace(store, root, inbox, categories)
    return FakeSession(namespace), inbox, sub, archive, deleted
