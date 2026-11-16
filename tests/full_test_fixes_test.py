"""
Regression tests for the issues found in the full live test of 0.4.4 against
New Outlook 16.115. Runs without Outlook (fake bridge); the AppleScript
handler checks run osascript only on macOS.

Run: python tests/full_test_fixes_test.py
"""
import asyncio
import json
import os
import re
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server_mac  # noqa: E402
from outlook_desktop_mcp.applescript_bridge import friendly_error  # noqa: E402
from outlook_desktop_mcp.utils.applescript_helpers import (  # noqa: E402
    DELIM, RECORD_DELIM, SANITIZE_HANDLER,
)
from outlook_desktop_mcp.utils.output import set_server_version  # noqa: E402

server_mac._UI_FOLDER_SETTLE_SECONDS = 0  # no waiting in tests
passed = total = 0


def check(name, cond, detail=""):
    global passed, total
    total += 1
    if cond:
        passed += 1
        print(f"  PASS: {name}")
    else:
        print(f"  FAIL: {name} {detail}")


class Fake:
    """Fake bridge: `accounts` answers the account-count script; `handler`
    (if given) answers everything else, otherwise `out`."""

    def __init__(self, accounts="1", out="", handler=None,
                 sidebar=("Inbox; 0 unread messages",), selected=1):
        self.accounts, self.out, self.handler, self.scripts = accounts, out, handler, []
        self.sidebar, self.selected, self.sidebar_calls = sidebar, selected, []

    def selects(self):
        return [int(m.group(1)) for x in self.sidebar_calls
                for m in [re.search(r"select row (\d+) of outline", x)] if m]

    async def run(self, script, timeout=None):
        if "outline 1 of scroll area 1" in script:  # sidebar
            self.sidebar_calls.append(script)
            m = re.search(r"select row (\d+) of outline", script)
            if m:
                self.selected = int(m.group(1))
                return ""
            if self.sidebar is None:
                raise RuntimeError("AppleScript error: Invalid index. (-1719)")
            return "".join(f"{i}{DELIM}{label}{DELIM}{'true' if i == self.selected else 'false'}{RECORD_DELIM}"
                           for i, label in enumerate(self.sidebar, 1))
        self.scripts.append(script)
        if "count of exchange accounts" in script:
            return self.accounts
        if self.handler:
            return self.handler(script)
        return self.out


def call(coro):
    return asyncio.run(coro)


def test_mail_tools_refuse_without_account():
    print("--- mail-writing tools refuse when Outlook has no visible account ---")
    tools = {
        "send_email": lambda: server_mac.send_email(to="a@example.invalid", subject="s", body="b"),
        "save_draft": lambda: server_mac.save_draft(to="a@example.invalid", subject="s", body="b"),
        "send_draft": lambda: server_mac.send_draft(draft_id="12"),
        "reply_email": lambda: server_mac.reply_email(entry_id="12", body="b"),
        "create_draft_reply": lambda: server_mac.create_draft_reply(entry_id="12", body="b"),
        "create_meeting": lambda: server_mac.create_meeting(
            subject="s", start="2026-12-01 10:00", end="2026-12-01 11:00",
            required_attendees="a@example.invalid"),
    }
    for name, fn in tools.items():
        fake = Fake(accounts="0", out="1")
        server_mac.bridge = fake
        out = call(fn())
        ran = [s for s in fake.scripts if "count of exchange accounts" not in s]
        check(f"{name}: refused with Legacy Outlook hint", out.startswith("Error: cannot") and "Legacy Outlook" in out, out[:100])
        check(f"{name}: no Outlook script run", ran == [], str(len(ran)))
    fake = Fake(accounts="2", out="1" + DELIM + "s")
    server_mac.bridge = fake
    call(server_mac.save_draft(to="a@example.invalid", subject="s", body="b"))
    check("with an account, save_draft runs", any("make new outgoing message" in s for s in fake.scripts))


def test_sanitize_handler_ids_and_dates():
    print("--- sf() writes big ids as digits and dates as ISO ---")
    check("handler has isod/idt", "on isod(d)" in SANITIZE_HANDLER and "on idt(n)" in SANITIZE_HANDLER)
    fake = Fake(out="")
    server_mac.bridge = fake
    call(server_mac.list_tasks())
    check("task script passes the raw id to sf", "my sf(item i of idList)" in "".join(fake.scripts))
    if sys.platform == "darwin" and shutil.which("osascript"):
        body = ('set d to current date\nset year of d to 2026\nset month of d to 12\nset day of d to 1\n'
                'set time of d to 37230\nreturn my sf(2130706432) & "|" & my sf(d) & "|" & my sf(7)')
        r = subprocess.run(["osascript", "-e", SANITIZE_HANDLER + body], capture_output=True, text=True, timeout=30)
        check("osascript: id digits, ISO date", r.stdout.strip() == "2130706432|2026-12-01T10:20:30|7", r.stdout + r.stderr)
    else:
        print("  SKIP: osascript checks need macOS")


def test_dates_validated():
    print("--- dates: friendly errors, dotted day-first accepted, end >= start ---")
    server_mac.bridge = Fake(out="")
    out = call(server_mac.list_events(start_date="not-a-date"))
    check("list_events bad date -> Error text", out.startswith("Error: start_date 'not-a-date'"), out)
    out = call(server_mac.list_events(start_date="16.11.2026", end_date="23.11.2026"))
    check("dotted day-first accepted", not out.startswith("Error"), out[:80])
    out = call(server_mac.list_events(start_date="2026-12-10", end_date="2026-12-01"))
    check("list_events end before start", out == "Error: end_date is before start_date.", out)
    out = call(server_mac.create_event(subject="s", start="2026-12-01 12:00", end="2026-12-01 11:00"))
    check("create_event end before start", out == "Error: end is before start.", out)
    out = call(server_mac.create_event(subject="s", start="next tuesday", end="2026-12-01 11:00"))
    check("create_event bad date", out.startswith("Error: start 'next tuesday'"), out)
    out = call(server_mac.create_event(subject="s", start="2026-12-01 10:00", end="2026-12-01 11:00", reminder_minutes=-5))
    check("negative reminder rejected", out.startswith("Error: reminder_minutes"), out)
    out = call(server_mac.update_event(entry_id="1", start="2026-12-01 12:00", end="2026-12-01 11:00"))
    check("update_event end before start", out == "Error: end is before start.", out)
    out = call(server_mac.create_task(subject="s", due_date="someday"))
    check("create_task bad due date", out.startswith("Error: due_date 'someday'"), out)
    out = call(server_mac.create_task(subject="s", importance="urgent"))
    check("create_task bad importance", out.startswith("Error: importance"), out)
    out = call(server_mac.search_events(query="x", start_date="xx"))
    check("search_events bad date", out.startswith("Error: start_date 'xx'"), out)


def test_search_events_range():
    print("--- search_events applies its date range ---")
    def ev(i, start, end):
        return DELIM.join([str(i), f"E{i}", start, end, "", "", "false"]) + RECORD_DELIM
    raw = ev(1, "2026-12-01T10:00:00", "2026-12-01T11:00:00") + ev(2, "2027-03-01T10:00:00", "2027-03-01T11:00:00")
    server_mac.bridge = Fake(out=raw)
    res = json.loads(call(server_mac.search_events(query="E", start_date="2026-11-30", end_date="2026-12-05")))
    check("only the event inside the range", [r["entry_id"] for r in res] == ["1"], str(res))


def test_counts_validated():
    print("--- counts and indexes must be positive ---")
    server_mac.bridge = Fake(out="")
    cases = {
        "list_emails count 0": server_mac.list_emails(count=0),
        "list_emails count -1": server_mac.list_emails(count=-1),
        "search_emails count 0": server_mac.search_emails(query="a", count=0),
        "list_events count 0": server_mac.list_events(count=0),
        "search_events count 0": server_mac.search_events(query="a", count=0),
        "list_tasks count 0": server_mac.list_tasks(count=0),
        "list_drafts limit 0": server_mac.list_drafts(limit=0),
        "list_drafts offset -1": server_mac.list_drafts(offset=-1),
        "list_folders depth -1": server_mac.list_folders(max_depth=-1),
        "save_attachment index 0": server_mac.save_attachment(entry_id="1", attachment_index=0),
    }
    for name, coro in cases.items():
        out = call(coro)
        check(name, out.startswith("Error:") and "must be a whole number" in out, out[:90])


def test_task_priority_clean():
    print("--- task priority without the enum prefix ---")
    raw = DELIM.join(["1", "T", "2026-12-05T00:00:00", "false", "priority high"]) + RECORD_DELIM
    server_mac.bridge = Fake(out=raw)
    res = json.loads(call(server_mac.list_tasks()))
    check("list_tasks priority 'high'", res and res[0]["priority"] == "high", str(res))


def test_friendly_errors():
    print("--- not-found errors are readable ---")
    check("message id real -> digits",
          friendly_error("execution error: Microsoft Outlook got an error: Can’t get message id 9.99999999E+8. (-1728)")
          == "Outlook has no message with id 999999999.")
    check("folder with quote",
          friendly_error('Can’t get mail folder "a\\"b". (-1728)').startswith("Outlook has no folder named 'a\"b'"))
    check("other errors untouched", friendly_error("Can’t get window 1. (-1719)") is None)


ROW = "Alice, Quarterly report,    10:00,"
ROW2 = "Bob, Lunch,    11:00,"


def ui_handler(pane=""):
    def h(script):
        if "set rowCount to count of rows" in script:
            return ROW + RECORD_DELIM + ROW2 + RECORD_DELIM
        if "select row" in script:
            return ""
        if "emailSubject" in script:
            return pane
        if "NOT_FOUND" in script:
            return "NOT_FOUND"  # subject lookup in the empty AppleScript inbox
        return ""  # AppleScript inbox is empty (New Outlook)
    return h


def test_inbox_search_via_window():
    print("--- search_emails / read_email by subject use the window inbox ---")
    server_mac.bridge = Fake(accounts="0", handler=ui_handler())
    res = json.loads(call(server_mac.search_emails(query="lunch")))
    check("search finds the window row", [r["sender_name"] for r in res] == ["Bob"], str(res))
    pane = DELIM.join(["Lunch", "Bob", "Mon 11:00", "me", "Body text"])
    server_mac.bridge = Fake(accounts="0", handler=ui_handler(pane))
    res = json.loads(call(server_mac.read_email(subject_search="quarterly")))
    check("read_email by subject reads the matching row", res.get("entry_id", "").startswith("ui-1-"), str(res)[:150])


SIDEBAR = ("Favorites", "Inbox; 1 unread messages", "Archive", "Drafts", "Sent", "Deleted Items")


def test_folder_switching():
    print("--- New Outlook: other folders are selected in the sidebar, then restored ---")
    fake = Fake(accounts="0", handler=ui_handler(), sidebar=SIDEBAR, selected=2)
    server_mac.bridge = fake
    res = json.loads(call(server_mac.list_emails(folder="sent", count=5)))
    check("sent read through the window", len(res) == 2, str(res)[:100])
    check("Sent selected, then Inbox again", fake.selects() == [5, 2], str(fake.selects()))
    check("selection restored", fake.selected == 2)

    fake = Fake(accounts="0", handler=ui_handler(), sidebar=SIDEBAR, selected=2)
    server_mac.bridge = fake
    call(server_mac.list_emails(count=5))
    check("inbox already shown: no switching", fake.selects() == [], str(fake.selects()))

    fake = Fake(accounts="0", handler=ui_handler(), sidebar=SIDEBAR, selected=5)
    server_mac.bridge = fake
    call(server_mac.list_emails(folder="inbox", count=5))
    check("Sent shown, inbox wanted: switch and restore", fake.selects() == [2, 5], str(fake.selects()))

    fake = Fake(accounts="0", handler=ui_handler(), sidebar=SIDEBAR, selected=2)
    server_mac.bridge = fake
    out = call(server_mac.list_emails(folder="Nope", count=5))
    check("unknown folder: clear error", out.startswith("Error listing emails: Outlook's folder list has no folder 'Nope'"), out[:120])
    check("unknown folder: nothing selected", fake.selects() == [])

    fake = Fake(accounts="0", handler=ui_handler(), sidebar=("Posteingang; 2 ungelesene Nachrichten", "Gesendete Elemente"), selected=2)
    server_mac.bridge = fake
    call(server_mac.list_emails(folder="inbox", count=5))
    check("German names: Posteingang selected for inbox", fake.selects() == [1, 2], str(fake.selects()))

    fake = Fake(accounts="0", handler=ui_handler(), sidebar=SIDEBAR, selected=2)
    server_mac.bridge = fake
    res = json.loads(call(server_mac.search_emails(query="lunch", folder="archive")))
    check("search in Archive switches and restores", fake.selects() == [3, 2] and len(res) == 1, str(fake.selects()))

    pane = DELIM.join(["Lunch", "Bob", "Mon 11:00", "me", "Body text"])
    fake = Fake(accounts="0", handler=ui_handler(pane), sidebar=SIDEBAR, selected=2)
    server_mac.bridge = fake
    fp = server_mac._ui_fingerprint("Bob", "Lunch", "11:00")
    res = json.loads(call(server_mac.read_email(entry_id=f"ui-2-{fp}", folder="sent")))
    check("read_email ui id in Sent", res.get("subject") == "Lunch" and fake.selects() == [5, 2], str(res)[:100] + str(fake.selects()))

    fake = Fake(accounts="0", handler=ui_handler(), sidebar=SIDEBAR + ("Junk", "Junk Email; 44 messages"), selected=2)
    server_mac.bridge = fake
    call(server_mac.list_emails(folder="junk", count=5))
    check("'junk' picks Junk Email over a custom folder named Junk", fake.selects() == [8, 2], str(fake.selects()))
    fake = Fake(accounts="0", handler=ui_handler(), sidebar=SIDEBAR + ("Projects",), selected=2)
    server_mac.bridge = fake
    call(server_mac.list_emails(folder="projects", count=5))
    check("custom folder by name", fake.selects() == [7, 2], str(fake.selects()))

    fake = Fake(accounts="0", handler=ui_handler(), sidebar=None)
    server_mac.bridge = fake
    out = call(server_mac.list_emails(folder="sent", count=5))
    check("sidebar unreadable: falls back to the AppleScript result", "Legacy Outlook" in out or out.strip() == "[]", out[:100])
    out = call(server_mac.list_emails(count=5))
    check("sidebar unreadable: inbox still read", len(json.loads(out)) == 2, out[:100])


def test_has_files_marker():
    print("--- 'Has files' marks an attachment ---")
    row = "Carol, Contract,    09:00,    Has files,    preview text,"
    fake = Fake(handler=lambda s: row + RECORD_DELIM if "set rowCount" in s else "")
    res = call(server_mac._ui_list_messages(fake, 5))
    check("'Has files' marks an attachment", res and res[0]["has_attachments"] is True, str(res))


def test_batched_rows_script():
    print("--- window rows are read in one batched call ---")
    fake = Fake(handler=lambda s: "")
    call(server_mac._ui_list_messages(fake, 5))
    s = next(x for x in fake.scripts if "set rowCount" in x)
    check("batched description read", "description of UI element 1 of rows 1 thru maxRows" in s)
    check("falls back to per-row loop", "if not batched then" in s and "of row i" in s)


def test_server_version():
    print("--- handshake reports the package version ---")
    class M:
        class _mcp_server:
            version = None
    import importlib.metadata as md
    orig = md.version
    md.version = lambda pkg: "9.9.9" if pkg == "outlook-desktop-mcp-plus" else orig(pkg)
    try:
        set_server_version(M)
    finally:
        md.version = orig
    check("version set from package metadata", M._mcp_server.version == "9.9.9", M._mcp_server.version)


if __name__ == "__main__":
    test_mail_tools_refuse_without_account()
    test_sanitize_handler_ids_and_dates()
    test_dates_validated()
    test_search_events_range()
    test_counts_validated()
    test_task_priority_clean()
    test_friendly_errors()
    test_inbox_search_via_window()
    test_folder_switching()
    test_has_files_marker()
    test_batched_rows_script()
    test_server_version()
    print("=" * 50)
    print(f"{passed}/{total} checks passed")
    sys.exit(0 if passed == total else 1)
