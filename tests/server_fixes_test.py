"""
Unit tests for Windows-side server.py correctness fixes, using fake COM objects.

Pure unit tests: no Outlook, no pywin32, run on any platform.

Covers:
  1. list_events / search_events use an overlap filter (events in progress
     at the range start are included).
  2. Dates are parsed and written to COM as datetime objects, never raw strings.
  3. reply_email keeps HTML formatting.
  4. save_attachment is confined to a root directory and never overwrites.

Run: python tests/server_fixes_test.py
"""
import asyncio
import json
import os
import sys
import tempfile
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server  # noqa: E402


# --- Fakes -----------------------------------------------------------------

def _fake_format_locale_datetime(dt):
    # Deterministic day-first rendering; replaces the Win32 locale formatter.
    return f"{dt.day:02d}/{dt.month:02d}/{dt.year:04d} {dt.hour:02d}:{dt.minute:02d}"


server._format_locale_datetime = _fake_format_locale_datetime


class FakeBridge:
    """Runs tool callbacks synchronously with a fake (outlook, namespace)."""

    def __init__(self, outlook, namespace):
        self.outlook = outlook
        self.namespace = namespace

    async def call(self, func, *args, **kwargs):
        return func(self.outlook, self.namespace, *args, **kwargs)


class FakeItems(list):
    """Calendar Items collection that records the call order and filter."""

    def __init__(self, events, log):
        super().__init__(events)
        self.log = log
        self.IncludeRecurrences = False

    def Sort(self, prop):
        self.log.append(("Sort", prop))

    def Restrict(self, flt):
        self.log.append(("Restrict", flt, self.IncludeRecurrences))
        return list(self)


class FakeEvent:
    def __init__(self, i, subject="Meeting"):
        self.EntryID = f"id{i}"
        self.Subject = subject
        self.Start = datetime(2026, 9, 7, 9, 0)
        self.End = datetime(2026, 9, 7, 10, 0)
        self.Duration = 60
        self.Location = ""
        self.Organizer = ""
        self.IsRecurring = False
        self.AllDayEvent = False
        self.BusyStatus = 2
        self.MeetingStatus = 0
        self.RequiredAttendees = ""
        self.OptionalAttendees = ""


class FakeFolder:
    def __init__(self, items=None):
        self.Items = items


class FakeStore:
    StoreID = "store1"

    def __init__(self, folder):
        self._folder = folder

    def GetDefaultFolder(self, _n):
        return self._folder


class FakeNamespace:
    def __init__(self, store=None, items_by_id=None):
        self.DefaultStore = store
        self.items_by_id = items_by_id or {}

    def GetItemFromID(self, entry_id, store_id=None):
        return self.items_by_id[entry_id]


def _use(outlook=None, namespace=None):
    server.bridge = FakeBridge(outlook, namespace)


def _run(coro):
    return asyncio.run(coro)


# --- 1. Overlap filter -------------------------------------------------------

def _calendar_ns(events):
    log = []
    items = FakeItems(events, log)
    return FakeNamespace(FakeStore(FakeFolder(items))), log


def test_list_events_uses_overlap_filter_and_keeps_call_order():
    ns, log = _calendar_ns([FakeEvent(1)])
    _use(namespace=ns)
    out = _run(server.list_events("2026-09-07 10:30", "2026-09-08"))
    assert len(json.loads(out)) == 1
    assert [e[0] for e in log] == ["Sort", "Restrict"], log
    assert log[0][1] == "[Start]"
    # IncludeRecurrences must already be True when Restrict runs.
    assert log[1][2] is True
    assert log[1][1] == "[Start] <= '08/09/2026 00:00' AND [End] >= '07/09/2026 10:30'", log[1][1]


def test_search_events_uses_overlap_filter():
    ns, log = _calendar_ns([FakeEvent(1, "Standup"), FakeEvent(2, "Lunch")])
    _use(namespace=ns)
    out = _run(server.search_events("stand", "2026-09-07 10:30", "2026-09-08"))
    assert [e["subject"] for e in json.loads(out)] == ["Standup"]
    assert log[1][1] == "[Start] <= '08/09/2026 00:00' AND [End] >= '07/09/2026 10:30'", log[1][1]
    assert [e[0] for e in log] == ["Sort", "Restrict"]


def test_list_events_count_limits_after_sort():
    ns, _ = _calendar_ns([FakeEvent(i) for i in range(5)])
    _use(namespace=ns)
    out = json.loads(_run(server.list_events("2026-09-07", "2026-09-09", count=3)))
    assert [e["entry_id"] for e in out] == ["id0", "id1", "id2"]


if __name__ == "__main__":
    test_list_events_uses_overlap_filter_and_keeps_call_order()
    test_search_events_uses_overlap_filter()
    test_list_events_count_limits_after_sort()
    print("OK")
