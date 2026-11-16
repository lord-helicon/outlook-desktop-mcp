"""
Unit test for the email formatters in utils/formatting.py.

Pure unit test: no Outlook, no pywin32, runs on any platform. Fake items
mimic late-bound pywin32 COM objects, where reading a property the item
type doesn't have raises AttributeError named after the call that produced
the object (e.g. "GetItemFromID.To").

Regression for upstream issue #17: read_email failed with
"AttributeError: GetItemFromID.To" when the newest Inbox item was a meeting
request or delivery report, which have no To/CC properties.

Run: .venv\\Scripts\\python tests\\formatting_test.py
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp.utils.formatting import (  # noqa: E402
    format_email_full,
    format_email_summary,
)
from outlook_desktop_mcp.utils.errors import format_com_error  # noqa: E402


class FakeComError(Exception):
    """Stands in for pywintypes.com_error raised by a property getter."""


class FakeCollection:
    def __init__(self, items):
        self._items = items

    @property
    def Count(self):
        return len(self._items)

    def Item(self, index):
        return self._items[index - 1]


class FakeRecipient:
    def __init__(self, name, type_):
        self.Name = name
        self.Type = type_


class FakeItem:
    """Late-bound COM item: unknown properties raise AttributeError."""

    def __init__(self, **props):
        self._props = props

    def __getattr__(self, name):
        props = self.__dict__["_props"]
        if name not in props:
            raise AttributeError(f"GetItemFromID.{name}")
        value = props[name]
        if isinstance(value, Exception):
            raise value
        return value


def mail_item():
    return FakeItem(
        EntryID="00A1",
        Subject="Quarterly numbers",
        SenderEmailAddress="alice@example.com",
        SenderName="Alice",
        ReceivedTime="2026-11-16 09:00:00",
        UnRead=True,
        Attachments=FakeCollection(["a.pdf"]),
        MessageClass="IPM.Note",
        To="Bob",
        CC="Carol",
        Body="Hello",
    )


def meeting_request():
    """A MeetingItem: no To/CC, recipients only via Recipients."""
    return FakeItem(
        EntryID="00B2",
        Subject="Planning",
        SenderEmailAddress="alice@example.com",
        SenderName="Alice",
        ReceivedTime="2026-11-16 10:00:00",
        UnRead=True,
        Attachments=FakeCollection([]),
        MessageClass="IPM.Schedule.Meeting.Request",
        Recipients=FakeCollection([
            FakeRecipient("Bob", 1),
            FakeRecipient("Dan", 1),
            FakeRecipient("Carol", 2),
        ]),
        Body="Agenda",
    )


def delivery_report():
    """A ReportItem: no sender, no ReceivedTime, Body getter fails."""
    return FakeItem(
        EntryID="00C3",
        Subject="Undeliverable: Planning",
        CreationTime="2026-11-16 11:00:00",
        UnRead=False,
        Attachments=FakeCollection([]),
        MessageClass="REPORT.IPM.Note.NDR",
        Body=FakeComError("property not available"),
    )


def test_mail_item_unchanged():
    r = format_email_full(mail_item())
    assert r["to"] == "Bob" and r["cc"] == "Carol"
    assert r["body"] == "Hello"
    assert r["attachment_count"] == 1 and r["has_attachments"] is True
    assert r["message_class"] == "IPM.Note"


def test_meeting_request_reads_recipients():
    r = format_email_full(meeting_request())
    assert r["to"] == "Bob; Dan", r["to"]
    assert r["cc"] == "Carol", r["cc"]
    assert r["body"] == "Agenda"
    assert r["message_class"] == "IPM.Schedule.Meeting.Request"


def test_delivery_report_does_not_raise():
    r = format_email_full(delivery_report())
    assert r["sender"] == "unknown" and r["sender_name"] == "unknown"
    assert r["received_time"] == "2026-11-16 11:00:00"
    assert r["to"] == "" and r["cc"] == ""
    assert r["body"] == ""


def test_summary_for_every_item_type():
    for item in (mail_item(), meeting_request(), delivery_report()):
        assert format_email_summary(item)["entry_id"]


def test_own_error_messages_reach_the_caller():
    assert format_com_error(ValueError("Account 'x' not found.")) == "Account 'x' not found."
    assert "timed out" in format_com_error(TimeoutError("COM operation timed out"))
    assert format_com_error(RuntimeError("C:\\secret\\path")) == "An unexpected error occurred."


if __name__ == "__main__":
    test_mail_item_unchanged()
    test_meeting_request_reads_recipients()
    test_delivery_report_does_not_raise()
    test_summary_for_every_item_type()
    test_own_error_messages_reach_the_caller()
    print("OK")
