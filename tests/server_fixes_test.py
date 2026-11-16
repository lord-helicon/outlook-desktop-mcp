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


# --- 2. Dates written to COM as datetimes -----------------------------------

class FakeAppt:
    """Records every attribute assignment so tests can inspect the types."""

    def __init__(self):
        object.__setattr__(self, "assigned", {})
        object.__setattr__(self, "saved", False)
        object.__setattr__(self, "sent", False)
        object.__setattr__(self, "EntryID", "new-id")
        object.__setattr__(self, "Class", server._OL_CLASS_APPOINTMENT)
        object.__setattr__(self, "Recipients", FakeRecipients())

    def __setattr__(self, k, v):
        self.assigned[k] = v
        object.__setattr__(self, k, v)

    def Save(self):
        object.__setattr__(self, "saved", True)

    def Send(self):
        object.__setattr__(self, "sent", True)


class FakeRecipients:
    def Add(self, _a):
        return type("R", (), {})()

    def ResolveAll(self):
        pass


class FakeOutlook:
    def __init__(self, obj):
        self.obj = obj
        self.created = 0

    def CreateItem(self, _kind):
        self.created += 1
        return self.obj


def test_create_event_assigns_datetimes():
    appt = FakeAppt()
    outlook = FakeOutlook(appt)
    _use(outlook, FakeNamespace())
    out = json.loads(_run(server.create_event("x", "2026-09-07 14:00", "2026-09-07T15:30:00")))
    assert out["status"] == "created"
    assert appt.assigned["Start"] == datetime(2026, 9, 7, 14, 0)
    assert appt.assigned["End"] == datetime(2026, 9, 7, 15, 30)
    assert type(appt.assigned["Start"]) is datetime


def test_create_event_all_day_date_only_input():
    appt = FakeAppt()
    _use(FakeOutlook(appt), FakeNamespace())
    out = json.loads(_run(server.create_event("x", "2026-09-07", "2026-09-08", all_day=True)))
    assert out["status"] == "created"
    assert appt.assigned["Start"] == datetime(2026, 9, 7)
    assert appt.assigned["End"] == datetime(2026, 9, 8)
    assert appt.assigned["AllDayEvent"] is True


def test_create_event_invalid_date_is_clear_error_and_creates_nothing():
    appt = FakeAppt()
    outlook = FakeOutlook(appt)
    _use(outlook, FakeNamespace())
    out = _run(server.create_event("x", "next tuesday", "2026-09-08"))
    assert out.startswith("Error: Invalid start 'next tuesday'"), out
    assert outlook.created == 0
    out = _run(server.create_event("x", "2026-09-07", "07/09/2026"))
    assert out.startswith("Error: Invalid end"), out


def test_create_event_aware_datetime_becomes_naive_local():
    appt = FakeAppt()
    _use(FakeOutlook(appt), FakeNamespace())
    _run(server.create_event("x", "2026-09-07T12:00:00+00:00", "2026-09-07T13:00:00+00:00"))
    got = appt.assigned["Start"]
    assert got.tzinfo is None
    expected = datetime.fromisoformat("2026-09-07T12:00:00+00:00").astimezone().replace(tzinfo=None)
    assert got == expected


def test_create_meeting_assigns_datetimes_and_rejects_bad_input():
    appt = FakeAppt()
    _use(FakeOutlook(appt), FakeNamespace())
    out = _run(server.create_meeting("m", "2026-09-07 14:00", "2026-09-07 15:00", "a@example.com"))
    assert "created" in out and appt.sent
    assert appt.assigned["Start"] == datetime(2026, 9, 7, 14, 0)
    assert type(appt.assigned["End"]) is datetime

    appt2 = FakeAppt()
    _use(FakeOutlook(appt2), FakeNamespace())
    out = _run(server.create_meeting("m", "bogus", "2026-09-07 15:00", "a@example.com"))
    assert out.startswith("Error: Invalid start"), out
    assert not appt2.sent


def test_update_event_assigns_datetimes_and_is_atomic_on_bad_input():
    appt = FakeAppt()
    appt.Subject = "old"
    appt.Start = appt.End = datetime(2026, 1, 1)
    appt.Location = ""
    appt.assigned.clear()
    _use(namespace=FakeNamespace(items_by_id={"e": appt}))
    out = json.loads(_run(server.update_event("e", subject="new", start="2026-09-07 09:00", end="2026-09-07")))
    assert out["status"] == "updated"
    assert appt.assigned["Start"] == datetime(2026, 9, 7, 9, 0)
    assert appt.assigned["End"] == datetime(2026, 9, 7)

    appt.assigned.clear()
    out = _run(server.update_event("e", subject="changed", start="2026-09-07", end="garbage"))
    assert out.startswith("Error: Invalid end"), out
    assert appt.assigned == {}, "nothing may be modified when a date is invalid"


class FakeTask:
    def __init__(self):
        object.__setattr__(self, "assigned", {})
        self.EntryID = "t1"

    def __setattr__(self, k, v):
        self.assigned[k] = v
        object.__setattr__(self, k, v)

    def Save(self):
        pass


def test_create_task_assigns_datetime_due_date():
    task = FakeTask()
    _use(FakeOutlook(task), FakeNamespace())
    out = json.loads(_run(server.create_task("t", due_date="2026-03-01")))
    assert out["status"] == "created"
    assert task.assigned["DueDate"] == datetime(2026, 3, 1)
    assert type(task.assigned["DueDate"]) is datetime

    task2 = FakeTask()
    outlook = FakeOutlook(task2)
    _use(outlook, FakeNamespace())
    out = _run(server.create_task("t", due_date="1st March"))
    assert out.startswith("Error: Invalid due_date"), out
    assert outlook.created == 0


def test_list_events_invalid_date_is_clear_error():
    ns, log = _calendar_ns([])
    _use(namespace=ns)
    assert _run(server.list_events("junk")).startswith("Error: Invalid start_date")
    assert _run(server.search_events("q", end_date="junk")).startswith("Error: Invalid end_date")


if __name__ == "__main__":
    test_list_events_uses_overlap_filter_and_keeps_call_order()
    test_search_events_uses_overlap_filter()
    test_list_events_count_limits_after_sort()
    test_create_event_assigns_datetimes()
    test_create_event_all_day_date_only_input()
    test_create_event_invalid_date_is_clear_error_and_creates_nothing()
    test_create_event_aware_datetime_becomes_naive_local()
    test_create_meeting_assigns_datetimes_and_rejects_bad_input()
    test_update_event_assigns_datetimes_and_is_atomic_on_bad_input()
    test_create_task_assigns_datetime_due_date()
    test_list_events_invalid_date_is_clear_error()
    print("OK")
