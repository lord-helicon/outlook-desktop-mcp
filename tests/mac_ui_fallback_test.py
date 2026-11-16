"""
Outlook Desktop MCP - macOS UI fallback / id validation test
=============================================================
Unit tests for the New Outlook UI-scripting fallback and for message/event/
task id validation in server_mac. Runs without Outlook: the AppleScript
bridge is replaced with a fake.

Run: python tests/mac_ui_fallback_test.py
"""
import sys
import os
import json
import asyncio

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server_mac
from outlook_desktop_mcp.utils.applescript_helpers import DELIM, RECORD_DELIM


def log(msg):
    print(msg, file=sys.stderr, flush=True)


class FakeBridge:
    """Captures scripts; `handler(script)` returns output or raises."""

    def __init__(self, handler=None, output=""):
        self.scripts = []
        self.handler = handler
        self.output = output

    async def run(self, script, timeout=None):
        self.scripts.append(script)
        if self.handler:
            return self.handler(script)
        return self.output

    async def run_lines(self, script, timeout=None):
        result = await self.run(script, timeout=timeout)
        return [line for line in result.split("\n") if line.strip()]


passed = 0
total = 0


def check(name, condition, detail=""):
    global passed, total
    total += 1
    if condition:
        passed += 1
        log(f"  PASS: {name}")
    else:
        log(f"  FAIL: {name} {detail}")


# New layout: 4 splitter groups + UI element wrapper. Legacy: 3, no wrapper.
def is_new_layout(script):
    return script.count("tell splitter group 1") == 4 and "tell UI element 1" in script


def is_legacy_layout(script):
    return script.count("tell splitter group 1") == 3 and "tell UI element 1" not in script


ROW = "Alice, Hello world,    10:00,"
INJECTIONS = ['1)\ndo shell script "id"\n(', "1 & 2", "ui - 2", "abc", "", "1; 2",
              '1" & (do shell script "id") & "', "²"]


def test_layouts_generate_both_paths():
    log("--- scope generation covers New and legacy layouts ---")
    new = server_mac._ui_scope(server_mac._UI_LAYOUTS[0], "X")
    leg = server_mac._ui_scope(server_mac._UI_LAYOUTS[1], "X")
    check("new: 4 splitters", new.count("tell splitter group 1") == 4)
    check("new: UI element wrapper before scroll area",
          "tell group 1\ntell UI element 1\ntell scroll area 1\ntell table 1" in new, new)
    check("legacy: 3 splitters, no wrapper",
          leg.count("tell splitter group 1") == 3
          and "tell group 1\ntell scroll area 1\ntell table 1" in leg, leg)
    check("tells balanced", new.count("tell ") == new.count("end tell"))


def test_list_falls_back_to_new_layout():
    log("--- _ui_list_messages: tries layouts until one works ---")
    # Only the legacy layout exists on this build; new path errors out.
    def legacy_only(script):
        if is_new_layout(script):
            raise RuntimeError("Can't get UI element 1")
        return ROW + RECORD_DELIM
    fake = FakeBridge(handler=legacy_only)
    res = asyncio.run(server_mac._ui_list_messages(fake, 5))
    check("both layouts attempted", len(fake.scripts) == 2)
    check("legacy result parsed", len(res) == 1 and res[0]["entry_id"] == "ui-1", str(res))

    # Only New Outlook layout works; legacy would error.
    def new_only(script):
        if is_legacy_layout(script):
            raise RuntimeError("Can't get group 1")
        return ROW + RECORD_DELIM
    fake = FakeBridge(handler=new_only)
    res = asyncio.run(server_mac._ui_list_messages(fake, 5))
    check("new layout tried first and sufficient", len(fake.scripts) == 1
          and is_new_layout(fake.scripts[0]))
    check("subject parsed", res and res[0]["subject"] == "Hello world", str(res))

    # New layout returns nothing -> falls through to legacy
    def empty_then_rows(script):
        return "" if is_new_layout(script) else ROW + RECORD_DELIM
    fake = FakeBridge(handler=empty_then_rows)
    res = asyncio.run(server_mac._ui_list_messages(fake, 5))
    check("empty new layout falls through to legacy", len(res) == 1 and len(fake.scripts) == 2)

    # Everything fails -> raises (list_emails swallows it)
    def boom(script):
        raise RuntimeError("nope")
    try:
        asyncio.run(server_mac._ui_list_messages(FakeBridge(handler=boom), 5))
        check("all layouts failing raises", False)
    except RuntimeError:
        check("all layouts failing raises", True)


def test_list_emails_uses_ui_fallback():
    log("--- list_emails falls back to UI scraping on empty inbox ---")
    def handler(script):
        if "System Events" in script:
            return ROW + RECORD_DELIM if is_new_layout(script) else ""
        return ""  # AppleScript inbox is empty
    server_mac.bridge = FakeBridge(handler=handler)
    res = json.loads(asyncio.run(server_mac.list_emails(folder="inbox", count=5)))
    check("ui-1 returned", len(res) == 1 and res[0]["entry_id"] == "ui-1", str(res))


def test_read_email_ui_id():
    log("--- read_email ui-N reads via reading pane ---")
    server_mac._UI_PANE_SETTLE_SECONDS = 0
    pane = DELIM.join(["Subj", "Bob, b@x.com", "Mon", "Me", "line1\nline2"])

    def handler(script):
        if "select row 3" in script:
            return ""
        if "emailSubject" in script:
            return pane
        raise AssertionError("unexpected script " + script[:80])
    fake = FakeBridge(handler=handler)
    server_mac.bridge = fake
    res = json.loads(asyncio.run(server_mac.read_email(entry_id="ui-3")))
    check("subject", res.get("subject") == "Subj", str(res))
    check("body keeps delimiter-free text", res.get("body") == "line1\nline2", str(res))
    check("entry_id echoed", res.get("entry_id") == "ui-3")
    check("no `message id` script run", not any("message id" in s for s in fake.scripts))
    check("select targets new layout first", is_new_layout(fake.scripts[0]))
    check("pane read uses 4 splitters",
          fake.scripts[1].count("tell splitter group 1") == 4)

    # select fails on both layouts -> clear JSON error
    def failing(script):
        raise RuntimeError("no table")
    server_mac.bridge = FakeBridge(handler=failing)
    res = json.loads(asyncio.run(server_mac.read_email(entry_id="ui-3")))
    check("selection failure reported", "error" in res, str(res))

    # ui-0 and junk are not treated as UI rows
    server_mac.bridge = FakeBridge(output="SHOULD NOT RUN")
    out = asyncio.run(server_mac.read_email(entry_id="ui-0"))
    check("ui-0 rejected", out.startswith("Error") and not server_mac.bridge.scripts, out)


def test_read_email_numeric_still_works():
    log("--- read_email numeric id still produces message id script ---")
    rec = DELIM.join(["101", "Hello", "a@b.com", "Alice", "t", "true", "0", "to;", "", "body"])
    fake = FakeBridge(output=rec)
    server_mac.bridge = fake
    res = json.loads(asyncio.run(server_mac.read_email(entry_id=" 101 ")))
    check("script uses stripped id", "set m to message id 101\n" in fake.scripts[0])
    check("parsed", res.get("subject") == "Hello", str(res))


# (tool name, kwargs, expected fragment that must NOT be run)
GUARDED = [
    ("read_email", {}),
    ("mark_as_read", {}),
    ("mark_as_unread", {}),
    ("move_email", {"target_folder": "archive"}),
    ("reply_email", {"body": "hi"}),
    ("create_draft_reply", {"body": "hi"}),
    ("list_attachments", {}),
    ("save_attachment", {}),
    ("get_event", {}),
    ("update_event", {"subject": "x"}),
    ("delete_event", {}),
    ("get_task", {}),
    ("complete_task", {}),
    ("delete_task", {}),
]


def test_injection_rejected_everywhere():
    log("--- every id-interpolating tool rejects bad ids ---")
    for name, kwargs in GUARDED:
        fn = getattr(server_mac, name)
        bad_ok = True
        ran = False
        for bad in INJECTIONS + ["ui-2"]:
            if name == "read_email" and bad in ("", "ui-2"):
                continue  # empty -> subject path; ui-2 is a valid UI id
            fake = FakeBridge(output="OK")
            server_mac.bridge = fake
            try:
                out = asyncio.run(fn(entry_id=bad, **kwargs))
            except TypeError as e:
                check(f"{name} signature", False, str(e))
                break
            if fake.scripts:
                ran = True
                bad_ok = False
                log(f"    {name} ran script for {bad!r}")
            elif not str(out).startswith("Error"):
                bad_ok = False
                log(f"    {name} non-error output {out!r} for {bad!r}")
        check(f"{name} rejects bad ids without running script", bad_ok and not ran)


def test_numeric_ids_still_work():
    log("--- numeric ids still reach the script unquoted ---")
    cases = [
        ("mark_as_read", {}, "message id 42\n"),
        ("mark_as_unread", {}, "message id 42\n"),
        ("list_attachments", {}, "message id 42\n"),
        ("get_event", {}, "calendar event id 42\n"),
        ("delete_event", {}, "calendar event id 42\n"),
        ("get_task", {}, "task id 42\n"),
        ("complete_task", {}, "task id 42\n"),
        ("delete_task", {}, "task id 42\n"),
    ]
    for name, kwargs, frag in cases:
        fake = FakeBridge(output="")
        server_mac.bridge = fake
        asyncio.run(getattr(server_mac, name)(entry_id=" 42 ", **kwargs))
        check(f"{name} runs with numeric id",
              len(fake.scripts) >= 1 and frag in fake.scripts[0], fake.scripts[:1] and fake.scripts[0][:120])


def main():
    test_layouts_generate_both_paths()
    test_list_falls_back_to_new_layout()
    test_list_emails_uses_ui_fallback()
    test_read_email_ui_id()
    test_read_email_numeric_still_works()
    test_injection_rejected_everywhere()
    test_numeric_ids_still_work()

    log("=" * 50)
    log(f"{passed}/{total} checks passed")
    if passed != total:
        sys.exit(1)


if __name__ == "__main__":
    main()
