"""
Unit test for the Table-based list_emails/search_emails fast path, the Items
fallback, and Instant Search (ci_phrasematch) query building.

Pure unit test with fake COM objects: no Outlook, no pywin32, any platform.
Run: python tests/mail_table_test.py
"""
import json
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server  # noqa: E402
from outlook_desktop_mcp.utils.formatting import format_email_summary  # noqa: E402

server._format_locale_datetime = lambda dt: f"{dt.day:02d}/{dt.month:02d}/{dt.year:04d} {dt.hour:02d}:{dt.minute:02d}"

HASATTACH = server._PR_HASATTACH_DASL


class FakeColumns:
    def __init__(self):
        self.added, self.removed = [], False

    def RemoveAll(self):
        self.removed = True

    def Add(self, name):
        self.added.append(name)


class FakeRow:
    def __init__(self, values):
        self._v = values

    def GetValues(self):
        return self._v


class FakeTable:
    def __init__(self, rows):
        self.Columns = FakeColumns()
        self._rows, self.sort_args, self.served = list(rows), None, 0

    def Sort(self, prop, desc):
        self.sort_args = (prop, desc)

    def GetNextRow(self):
        if not self._rows:
            return None
        self.served += 1
        return FakeRow(self._rows.pop(0))


class FakeStore:
    def __init__(self, instant):
        self.IsInstantSearchEnabled = instant


class FakeAttachments:
    def __init__(self, n):
        self.Count = n


class FakeItem:
    def __init__(self, i):
        self.EntryID = f"E{i}"
        self.Subject = f"S{i}"
        self.SenderEmailAddress = "a@x.com"
        self.SenderName = "A"
        self.ReceivedTime = datetime(2026, 3, 10, 9, i)
        self.UnRead = True
        self.Attachments = FakeAttachments(2)


class FakeItems:
    def __init__(self, n):
        self._items = [FakeItem(i) for i in range(n)]
        self.log = []
        self.Count = n

    def Sort(self, p, d):
        self.log.append(("Sort", p, d))

    def Restrict(self, f):
        self.log.append(("Restrict", f))
        return self

    def Item(self, i):
        if i == 2:
            raise RuntimeError("bad item")
        return self._items[i - 1]


class FakeFolder:
    def __init__(self, nrows=5, instant=False, table_raises=False):
        self.Store = FakeStore(instant)
        self.table_raises, self.table = table_raises, None
        self.get_table_args = None
        self.nrows = nrows
        self.Items = FakeItems(nrows)

    def GetTable(self, f, kind):
        self.get_table_args = (f, kind)
        if self.table_raises:
            raise RuntimeError("GetTable failed")
        self.table = FakeTable(
            [(f"E{i}", f"S{i}", "a@x.com", "A", datetime(2026, 3, 10, 9, i), True, True)
             for i in range(self.nrows)]
        )
        return self.table


class FakeBridge:
    async def call(self, fn, *a):
        return fn(None, None, *a)


def _run(coro):
    import asyncio
    return asyncio.run(coro)


def _setup(folder):
    server._require_store = lambda ns, acct="": object()
    server._resolve_folder = lambda ns, name, store=None: folder
    server.bridge = FakeBridge()


def test_table_path():
    f = FakeFolder(nrows=5)
    _setup(f)
    out = json.loads(_run(server.list_emails("inbox", 3, True)))
    assert f.get_table_args == ("[UnRead] = True", 0), f.get_table_args
    t = f.table
    assert t.Columns.removed
    assert t.Columns.added == ["EntryID", "Subject", "SenderEmailAddress", "SenderName",
                               "ReceivedTime", "UnRead", HASATTACH]
    assert t.sort_args == ("[ReceivedTime]", True)
    assert len(out) == 3 and t.served == 3
    ref = format_email_summary(FakeItem(0))
    assert set(out[0]) == set(ref), (set(out[0]), set(ref))
    assert out[0]["received_time"] == ref["received_time"]
    assert out[0]["has_attachments"] is True and out[0]["attachment_count"] is None
    assert out[0]["unread"] is True and out[0]["subject"] == "S0"
    assert f.Items.log == []  # Items path untouched
    print("ok table path")


def test_table_fewer_rows_and_empty_fields():
    f = FakeFolder(nrows=2)
    _setup(f)
    assert len(json.loads(_run(server.list_emails("inbox", 50)))) == 2
    r = server._table_row_to_summary(("E", None, None, None, "t", 0, 0))
    assert r["subject"] == "(no subject)" and r["sender"] == "unknown"
    assert r["has_attachments"] is False and r["unread"] is False
    print("ok short/empty")


def test_fallback():
    f = FakeFolder(nrows=4, table_raises=True)
    _setup(f)
    out = json.loads(_run(server.list_emails("inbox", 10, True)))
    assert f.Items.log == [("Sort", "[ReceivedTime]", True), ("Restrict", "[UnRead] = True")]
    assert [e["entry_id"] for e in out] == ["E0", "E2", "E3"]  # item 2 skipped
    assert out[0]["attachment_count"] == 2  # old formatter values
    # also when a later Table step raises
    f = FakeFolder(nrows=3)
    orig = f.GetTable
    def bad(*a):
        t = orig(*a)
        t.Sort = lambda *x: (_ for _ in ()).throw(RuntimeError("boom"))
        return t
    f.GetTable = bad
    _setup(f)
    out = json.loads(_run(server.search_emails("x", "inbox", 10)))
    assert f.Items.log[0][0] == "Restrict" and f.Items.log[1][0] == "Sort"
    assert len(out) == 2
    print("ok fallback")


def test_phrasematch():
    f = FakeFolder(instant=True)
    _setup(f)
    _run(server.search_emails("it's 100%_x", "inbox", 5))
    q = f.get_table_args[0]
    assert q.startswith("@SQL=")
    assert q.count("ci_phrasematch 'it''s 100%_x'") == 2, q
    assert "LIKE" not in q and "[%]" not in q and "[_]" not in q
    f = FakeFolder(instant=False)
    _setup(f)
    _run(server.search_emails("it's 100%", "inbox", 5))
    q = f.get_table_args[0]
    assert "ci_phrasematch" not in q and "LIKE '%it''s 100[%]%'" in q, q
    f = FakeFolder(instant=True)
    f.Store = object()  # property missing -> treated as disabled
    _setup(f)
    _run(server.search_emails("a", "inbox", 5))
    assert "LIKE" in f.get_table_args[0]
    f = FakeFolder(instant=True)
    _setup(f)
    _run(server.search_emails("a", "inbox", 5, start_date="2026-03-10"))
    q = f.get_table_args[0]
    assert 'datereceived" >= ' in q and 'datereceived" <= ' in q, q
    print("ok phrasematch")


if __name__ == "__main__":
    test_table_path()
    test_table_fewer_rows_and_empty_fields()
    test_fallback()
    test_phrasematch()
    print("ALL PASSED")
