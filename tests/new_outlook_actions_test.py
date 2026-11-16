"""
Tests for acting on New Outlook messages through the window (ui-N ids):
move_email via the Move dialog and mark_as_read/unread via Toggle Unread.
Runs without Outlook (fake System Events); on macOS the generated scripts are
also compiled with osacompile (nothing is run).

Run: python tests/new_outlook_actions_test.py
"""
import asyncio
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from outlook_desktop_mcp import server_mac  # noqa: E402
from outlook_desktop_mcp.utils.applescript_helpers import (  # noqa: E402
    DELIM, RECORD_DELIM, SANITIZE_HANDLER,
)

server_mac._UI_FOLDER_SETTLE_SECONDS = 0
server_mac._UI_SELECT_SETTLE_SECONDS = 0
passed = total = 0


def check(name, cond, detail=""):
    global passed, total
    total += 1
    if cond:
        passed += 1
        print(f"  PASS: {name}")
    else:
        print(f"  FAIL: {name} {detail}")


SIDEBAR = ("Favorites", "Inbox; 1 unread messages", "Archive", "Drafts", "Sent", "Deleted Items")
FP_BOB = server_mac._ui_fingerprint("Bob", "Lunch", "11:00")


class FakeOutlook:
    """Fake bridge: sidebar + message rows + Move dialog + Toggle Unread."""

    def __init__(self, bob_unread=False, move_reply="MOVED", read_on_select=False):
        self.selected_folder = 2
        self.bob_unread = bob_unread
        self.move_reply = move_reply
        self.read_on_select = read_on_select
        self.scripts, self.toggles, self.selected_rows = [], 0, []

    def rows(self):
        bob = ("Unread,    " if self.bob_unread else "     ") + "Bob, Lunch,    11:00,"
        return "Today" + RECORD_DELIM + "     Alice, Report,    10:00," + RECORD_DELIM + bob + RECORD_DELIM

    async def run(self, script, timeout=None):
        self.scripts.append(script)
        if "outline 1 of scroll area 1" in script:
            m = re.search(r"select row (\d+) of outline", script)
            if m:
                self.selected_folder = int(m.group(1))
                return ""
            return "".join(f"{i}{DELIM}{lab}{DELIM}{'true' if i == self.selected_folder else 'false'}{RECORD_DELIM}"
                           for i, lab in enumerate(SIDEBAR, 1))
        if "set rowCount to count of rows" in script:
            return self.rows()
        if "set focused to true\nselect row" in script:
            self.selected_rows.append(int(re.search(r"select row (\d+)", script).group(1)))
            if self.read_on_select:
                self.bob_unread = False
            return ""
        if "Toggle Unread" in script:
            self.toggles += 1
            self.bob_unread = not self.bob_unread
            return ""
        if "set value of text field 1 of sh to" in script:
            return self.move_reply
        return ""  # activate etc.


def call(coro):
    return asyncio.run(coro)


def test_move():
    print("--- move_email with a ui id goes through the Move dialog ---")
    fake = FakeOutlook(); server_mac.bridge = fake
    res = json.loads(call(server_mac.move_email(entry_id=f"ui-3-{FP_BOB}", target_folder="archive")))
    check("moved", res.get("status") == "moved" and res.get("to_folder") == "Archive", str(res))
    check("message row 3 selected first", fake.selected_rows == [3], str(fake.selected_rows))
    mv = next(s for s in fake.scripts if "set value of text field 1 of sh to" in s)
    check("dialog filtered by the display name", 'set value of text field 1 of sh to "Archive"' in mv)
    check("exact '<folder>, <mailbox>' label required", 'set wanted to "Archive, " & acct' in mv)
    check("no unique match -> cancel before Move",
          mv.index("if n is not 1 then") < mv.index('value of attribute "AXDefaultButton"'))
    check("Outlook brought to the front", any("to activate" in s for s in fake.scripts))

    fake = FakeOutlook(); server_mac.bridge = fake
    res = json.loads(call(server_mac.move_email(entry_id=f"ui-3-{FP_BOB}", target_folder="sent")))
    check("'sent' resolved to the sidebar name 'Sent'", res.get("to_folder") == "Sent", str(res))

    fake = FakeOutlook(); fake.selected_folder = 2; server_mac.bridge = fake
    call(server_mac.move_email(entry_id=f"ui-3-{FP_BOB}", target_folder="inbox", source_folder="sent"))
    folder_selects = [int(m.group(1)) for s in fake.scripts for m in [re.search(r"select row (\d+) of outline", s)] if m]
    check("source folder Sent selected, then Inbox restored", folder_selects == [5, 2], str(folder_selects))

    fake = FakeOutlook(move_reply="NOMATCH:2"); server_mac.bridge = fake
    out = call(server_mac.move_email(entry_id=f"ui-3-{FP_BOB}", target_folder="archive"))
    check("ambiguous folder -> nothing moved", out.startswith("Error moving email:") and "nothing was moved" in out, out)

    fake = FakeOutlook(); server_mac.bridge = fake
    out = call(server_mac.move_email(entry_id="ui-3", target_folder="archive"))
    check("id without fingerprint refused, no script run", "fingerprint" in out and fake.scripts == [], out)
    out = call(server_mac.move_email(entry_id="ui-3-deadbeef", target_folder="archive"))
    check("message gone -> error", "no longer in the first" in out, out)
    fake = FakeOutlook(); server_mac.bridge = fake
    out = call(server_mac.move_email(entry_id=f"ui-3-{FP_BOB}", target_folder="Inbox"))
    check("same folder refused", "already in" in out and fake.scripts == [], out)


def test_read_state():
    print("--- mark_as_read / mark_as_unread with ui ids ---")
    fake = FakeOutlook(bob_unread=True); server_mac.bridge = fake
    res = json.loads(call(server_mac.mark_as_read(entry_id=f"ui-3-{FP_BOB}")))
    check("unread -> read via Toggle Unread", res.get("status") == "marked_read" and fake.toggles == 1, str(res))

    fake = FakeOutlook(bob_unread=False); server_mac.bridge = fake
    res = json.loads(call(server_mac.mark_as_unread(entry_id=f"ui-3-{FP_BOB}")))
    check("read -> unread", res.get("status") == "marked_unread" and fake.toggles == 1, str(res))

    fake = FakeOutlook(bob_unread=True); server_mac.bridge = fake
    res = json.loads(call(server_mac.mark_as_unread(entry_id=f"ui-3-{FP_BOB}")))
    check("already unread -> unchanged, not selected", res.get("status") == "unchanged"
          and fake.toggles == 0 and fake.selected_rows == [], str(res))

    fake = FakeOutlook(bob_unread=True, read_on_select=True); server_mac.bridge = fake
    res = json.loads(call(server_mac.mark_as_read(entry_id=f"ui-3-{FP_BOB}")))
    check("selection already marked it read -> no toggle", res.get("status") == "marked_read" and fake.toggles == 0, str(res))

    out = call(server_mac.mark_as_read(entry_id="x1"))
    check("non-ui bad ids still rejected", "not a valid Outlook message id" in out, out)
    out = server_mac._bad_message_id("ui-3-deadbeef")
    check("id error lists the tools that accept ui ids", "move_email" in out and "mark_as_read" in out, out)


def test_scripts_compile():
    print("--- generated scripts compile (macOS only, nothing is run) ---")
    if sys.platform != "darwin" or not shutil.which("osacompile"):
        print("  SKIP: needs macOS")
        return
    tmp = tempfile.mkdtemp()
    for name, script in (("move", server_mac._ui_move_script('Archive "x"')),
                         ("toggle", server_mac._UI_TOGGLE_UNREAD_SCRIPT),
                         ("activate", server_mac._UI_ACTIVATE)):
        r = subprocess.run(["osacompile", "-o", os.path.join(tmp, name + ".scpt"), "-e", SANITIZE_HANDLER + script],
                           capture_output=True, text=True, timeout=60)
        check(f"{name} script compiles", r.returncode == 0, r.stderr[:200])
    shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    test_move()
    test_read_state()
    test_scripts_compile()
    print("=" * 50)
    print(f"{passed}/{total} checks passed")
    sys.exit(0 if passed == total else 1)
