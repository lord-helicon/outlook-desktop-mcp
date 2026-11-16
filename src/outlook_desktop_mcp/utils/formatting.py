"""Helpers for extracting and formatting Outlook item data."""
import re

from outlook_desktop_mcp.tools._folder_constants import (
    BUSY_STATUS_NAMES,
    MEETING_STATUS_NAMES,
    RESPONSE_NAMES,
    TASK_STATUS_NAMES,
    IMPORTANCE_NAMES,
)


def truncate(text: str, max_length: int = 2000) -> str:
    if len(text) <= max_length:
        return text
    return text[:max_length] + "\n... [truncated]"


def strip_html(html: str) -> str:
    text = re.sub(r"<[^>]+>", "", html)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _safe(item, name: str, default=None):
    """Read a COM property, returning default when the item type lacks it.

    Mail folders also hold MeetingItem, ReportItem (delivery/read receipts)
    and similar items that are missing some MailItem properties. On a
    late-bound pywin32 object, reading one raises AttributeError (e.g.
    "GetItemFromID.To"); some properties raise com_error instead.
    getattr(item, name, default) only covers the first case.
    """
    try:
        value = getattr(item, name)
    except Exception:
        return default
    return default if value is None else value


_OL_TO = 1
_OL_CC = 2


def _recipients_of_type(item, recipient_type: int) -> str:
    """Build a "; "-joined recipient list from item.Recipients."""
    names = []
    try:
        recipients = item.Recipients
        for i in range(recipients.Count):
            r = recipients.Item(i + 1)
            if r.Type == recipient_type:
                names.append(r.Name)
    except Exception:
        pass
    return "; ".join(names)


def _attachment_count(item) -> int:
    try:
        return int(item.Attachments.Count)
    except Exception:
        return 0


def format_email_summary(item) -> dict:
    """Extract key fields from a MailItem (or other item in a mail folder)."""
    received = _safe(item, "ReceivedTime") or _safe(item, "CreationTime")
    attachment_count = _attachment_count(item)
    return {
        "entry_id": item.EntryID,
        "subject": _safe(item, "Subject") or "(no subject)",
        "sender": _safe(item, "SenderEmailAddress", "unknown"),
        "sender_name": _safe(item, "SenderName", "unknown"),
        "received_time": str(received) if received is not None else None,
        "unread": bool(_safe(item, "UnRead", False)),
        "has_attachments": attachment_count > 0,
        "attachment_count": attachment_count,
    }


def format_email_full(item, body_max_length: int = 5000) -> dict:
    """Extract full email details including body.

    Works for meeting requests and delivery reports too: those have no
    To/CC properties, so the lists are rebuilt from item.Recipients.
    message_class (e.g. "IPM.Schedule.Meeting.Request") tells the caller
    what kind of item it is.
    """
    result = format_email_summary(item)
    result["message_class"] = _safe(item, "MessageClass", "")
    to = _safe(item, "To")
    cc = _safe(item, "CC")
    result["to"] = to if to is not None else _recipients_of_type(item, _OL_TO)
    result["cc"] = cc if cc is not None else _recipients_of_type(item, _OL_CC)
    result["body"] = truncate(_safe(item, "Body", ""), body_max_length)
    return result


# --- Calendar formatting ---


def format_event_summary(item) -> dict:
    """Extract key fields from an Outlook AppointmentItem."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "start": str(item.Start),
        "end": str(item.End),
        "duration": item.Duration,
        "location": item.Location or "",
        "organizer": item.Organizer or "",
        "is_recurring": bool(item.IsRecurring),
        "all_day": bool(item.AllDayEvent),
        "busy_status": BUSY_STATUS_NAMES.get(item.BusyStatus, "unknown"),
        "meeting_status": MEETING_STATUS_NAMES.get(item.MeetingStatus, "unknown"),
        "required_attendees": item.RequiredAttendees or "",
        "optional_attendees": item.OptionalAttendees or "",
    }


def format_event_full(item, body_max_length: int = 5000) -> dict:
    """Full event details including body."""
    result = format_event_summary(item)
    result["body"] = truncate(item.Body or "", body_max_length)
    result["reminder_set"] = bool(item.ReminderSet)
    result["reminder_minutes"] = (
        item.ReminderMinutesBeforeStart if item.ReminderSet else None
    )
    result["categories"] = item.Categories or ""
    result["response_status"] = RESPONSE_NAMES.get(item.ResponseStatus, "unknown")
    return result


# --- Task formatting ---


def format_task_summary(item) -> dict:
    """Extract key fields from an Outlook TaskItem."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "status": TASK_STATUS_NAMES.get(item.Status, "unknown"),
        "percent_complete": item.PercentComplete,
        "due_date": str(item.DueDate) if str(item.DueDate) != "01/01/4501" else None,
        "start_date": str(item.StartDate) if str(item.StartDate) != "01/01/4501" else None,
        "importance": IMPORTANCE_NAMES.get(item.Importance, "normal"),
        "complete": bool(item.Complete),
        "categories": item.Categories or "",
        "owner": item.Owner or "",
    }


def format_task_full(item, body_max_length: int = 5000) -> dict:
    """Full task details including body."""
    result = format_task_summary(item)
    result["body"] = truncate(item.Body or "", body_max_length)
    result["reminder_set"] = bool(item.ReminderSet)
    result["date_completed"] = (
        str(item.DateCompleted) if item.Complete else None
    )
    return result
