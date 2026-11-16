"""
Regression tests for the security-audit needs-validation leads that were fixed
defensively (no Outlook needed; one optional macOS check compiles the generated
AppleScript against Outlook's dictionary when Outlook for Mac is installed).

Leads covered:
- Windows: save_directory UNC/WebDAV/device paths are rejected before any
  filesystem call (no SMB/WebDAV connection from the containment check).
- Windows/macOS: saved attachments are marked as downloaded from the internet.
- macOS save_attachment: sender-controlled attachment names cannot choose the
  write path, existing files are not overwritten, the save stays in the root.
- macOS send_draft: only unsent outgoing messages can be sent.
- macOS move_email: a custom destination folder must be in the message's account.
- macOS mail/event/task bodies: plain text is HTML-escaped before it goes into
  Outlook's HTML `content` property; html_body uses `content` (there is no
  `html content` property).

Run: python tests/needs_validation_fixes_test.py
"""
import asyncio
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import types

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

_m, _s, _f = (types.ModuleType(n) for n in ("mcp", "mcp.server", "mcp.server.fastmcp"))


class _FastMCP:
    def __init__(self, *a, **k):
        pass

    def tool(self, *a, **k):
        return lambda fn: fn

    def run(self, *a, **k):
        pass


_f.FastMCP = _FastMCP
for _n, _mod in (("mcp", _m), ("mcp.server", _s), ("mcp.server.fastmcp", _f)):
    sys.modules.setdefault(_n, _mod)

from outlook_desktop_mcp import server_mac  # noqa: E402
from outlook_desktop_mcp.utils import attachments  # noqa: E402
from outlook_desktop_mcp.utils.applescript_helpers import SANITIZE_HANDLER, text_to_html  # noqa: E402

passed = total = 0


def check(name, cond, detail=""):
    global passed, total
    total += 1
    if cond:
        passed += 1
        print(f"  PASS: {name}")
    else:
        print(f"  FAIL: {name} {detail}")


class SaveBridge:
    """Fake AppleScript bridge for save_attachment: answers the name query and
    'performs' the save by creating the file at the path in the script."""

    def __init__(self, att_name):
        self.att_name, self.scripts = att_name, []

    async def run(self, script, timeout=None):
        self.scripts.append(script)
        m = re.search(r'save a in "((?:[^"\\]|\\.)*)"', script)
        if m:
            path = m.group(1).replace('\\"', '"').replace("\\\\", "\\")
            with open(path, "w") as f:
                f.write("dummy attachment")
            return ""
        return self.att_name


class Capture:
    def __init__(self, out=""):
        self.out, self.scripts = out, []

    async def run(self, script, timeout=None):
        self.scripts.append(script)
        return self.out


def test_unc_rejected_before_filesystem():
    print("--- save_directory network/device paths rejected before realpath ---")
    calls = []
    real = attachments.os.path.realpath
    attachments.os.path.realpath = lambda p: (calls.append(p), real(p))[1]
    try:
        root = real(tempfile.mkdtemp())  # attachment_root() always returns a realpath
        for bad in (r"\\attacker.example\share\x", "//attacker.example/share/x",
                    r"\\attacker.example@SSL@443\DavWWWRoot\x", r"\\?\UNC\attacker.example\s\x", r"\\.\pipe\x"):
            calls.clear()
            path, err = attachments.resolve_save_dir(bad, root)
            check(f"rejected without filesystem access: {bad[:28]}", path is None and err and not calls, (err, calls))
        path, err = attachments.resolve_save_dir("sub", root)
        check("relative folder inside root still accepted", err is None and path == os.path.join(root, "sub"), (path, err))
    finally:
        attachments.os.path.realpath = real


def test_safe_filename():
    print("--- attachment names reduced to one safe component ---")
    cases = {"../../evil.sh": "evil.sh", "..": "attachment", "a/b\\c.txt": "c.txt",
             "report.pdf": "report.pdf", "x:y?.txt": "x_y_.txt", "trailing. ": "trailing", "": "attachment"}
    for raw, want in cases.items():
        got = attachments.safe_filename(raw)
        check(f"{raw!r} -> {want!r}", got == want, got)


def test_mark_as_downloaded():
    print("--- saved files are marked as downloaded from the internet ---")
    tmp = tempfile.mkdtemp()
    path = os.path.join(tmp, "f.txt")
    open(path, "w").write("x")
    real_name = attachments.os.name
    try:
        attachments.os.name = "nt"  # exercise the Windows branch on any OS
        check("Windows: Zone.Identifier written", attachments.mark_as_downloaded(path))
        ads = path + ":Zone.Identifier"  # on non-Windows this is a plain sibling file
        check("Windows: ZoneId=3", os.path.exists(ads) and "ZoneId=3" in open(ads).read())
    finally:
        attachments.os.name = real_name
    if sys.platform == "darwin" and os.path.exists("/usr/bin/xattr"):
        os.remove(path + ":Zone.Identifier")
        check("macOS: quarantine set", attachments.mark_as_downloaded(path))
        q = subprocess.run(["/usr/bin/xattr", "-p", "com.apple.quarantine", path], capture_output=True, text=True)
        check("macOS: com.apple.quarantine present", q.returncode == 0 and attachments.QUARANTINE_AGENT in q.stdout, q.stdout)
    shutil.rmtree(tmp, ignore_errors=True)


def test_mac_save_attachment():
    print("--- macOS save_attachment: name cannot choose the path, no overwrite ---")
    root = tempfile.mkdtemp()
    outside = tempfile.mkdtemp()
    os.environ["OUTLOOK_MCP_ATTACHMENT_DIR"] = root
    try:
        fake = SaveBridge("../" + os.path.basename(outside) + "/evil.sh")
        server_mac.bridge = fake
        res = json.loads(asyncio.run(server_mac.save_attachment(entry_id="42")))
        real_root = os.path.realpath(root)
        check("saved inside the root", res.get("path") == os.path.join(real_root, "evil.sh"), res)
        check("nothing written outside the root", os.listdir(outside) == [])
        check("save script never concatenates the attachment name",
              not any("& aname" in s for s in fake.scripts) and 'save a in "' in fake.scripts[-1])
        res2 = json.loads(asyncio.run(server_mac.save_attachment(entry_id="42")))
        check("second save does not overwrite", res2.get("filename") == "evil (1).sh", res2)
        check("result reports the download mark", "marked_as_downloaded" in res2)

        before = set(os.listdir(root))
        out = asyncio.run(server_mac.save_attachment(entry_id="not-an-id", save_directory="newdir"))
        check("bad id rejected before any directory is created",
              out.startswith("Error") and set(os.listdir(root)) == before, out)
        for bad in (outside, os.path.join(root, "..", "x"), r"\\host\share"):
            fake = SaveBridge("a.txt"); server_mac.bridge = fake
            out = asyncio.run(server_mac.save_attachment(entry_id="42", save_directory=bad))
            check(f"save_directory outside root rejected: {bad[-20:]}", out.startswith("Error") and not fake.scripts, out)
    finally:
        del os.environ["OUTLOOK_MCP_ATTACHMENT_DIR"]
        shutil.rmtree(root, ignore_errors=True)
        shutil.rmtree(outside, ignore_errors=True)


def test_mac_send_draft_guard():
    print("--- macOS send_draft refuses anything but an unsent outgoing message ---")
    cap = Capture("s" + "\x01" + "to@example.invalid")
    server_mac.bridge = cap
    asyncio.run(server_mac.send_draft(draft_id="12"))
    s = next(x for x in cap.scripts if "send m" in x)
    guard_pos = min(s.find("is not outgoing message"), s.find("was sent of m"))
    check("checks class outgoing message", "class of m is not outgoing message" in s)
    check("checks was sent", "if was sent of m then error" in s)
    check("checks happen before send", guard_pos != -1 and guard_pos < s.find("send m"))


def test_mac_move_email_account():
    print("--- macOS move_email: custom folder must be in the message's account ---")
    cap = Capture("subject"); server_mac.bridge = cap
    asyncio.run(server_mac.move_email(entry_id="12", target_folder="Archive"))
    s = cap.scripts[0]
    check("account check before move", "id of account of destFolder" in s and s.find("account of destFolder") < s.find("move m to"))
    cap = Capture("subject"); server_mac.bridge = cap
    asyncio.run(server_mac.move_email(entry_id="12", target_folder="inbox"))
    check("built-in folders unchanged (no account check)", "destFolder" not in cap.scripts[0])
    cap = Capture("subject"); server_mac.bridge = cap
    asyncio.run(server_mac.move_email(entry_id="12", target_folder=" Archive "))
    check("folder name normalised", 'mail folder "Archive"' in cap.scripts[0])


def test_mac_bodies_escaped():
    print("--- macOS plain-text bodies are HTML-escaped into `content` ---")
    body = 'Hi <img src="http://tracker.invalid/p.gif"> & <!-- rest\nline2'
    calls = {
        "send_email": lambda: server_mac.send_email(to="a@example.invalid", subject="s", body=body),
        "save_draft": lambda: server_mac.save_draft(to="a@example.invalid", subject="s", body=body),
        "reply_email": lambda: server_mac.reply_email(entry_id="12", body=body),
        "create_draft_reply": lambda: server_mac.create_draft_reply(entry_id="12", body=body),
        "create_event": lambda: server_mac.create_event(subject="s", start="2026-12-01 10:00", end="2026-12-01 11:00", body=body),
        "create_meeting": lambda: server_mac.create_meeting(subject="s", start="2026-12-01 10:00", end="2026-12-01 11:00", required_attendees="a@example.invalid", body=body),
        "update_event": lambda: server_mac.update_event(entry_id="12", body=body),
        "create_task": lambda: server_mac.create_task(subject="s", body=body),
    }
    for name, call in calls.items():
        cap = Capture("1"); server_mac.bridge = cap
        asyncio.run(call())
        s = "\n".join(cap.scripts)
        check(f"{name}: no raw markup", "<img" not in s and "<!--" not in s)
        check(f"{name}: escaped markup present", "&lt;img" in s and "&lt;!--" in s)
    cap = Capture("1"); server_mac.bridge = cap
    asyncio.run(server_mac.send_email(to="a@example.invalid", subject="s", body="", html_body="<p>x</p>"))
    check("html_body uses `content`, not the non-existent `html content`",
          any('content:"<p>x</p>"' in x for x in cap.scripts) and not any("html content" in x for x in cap.scripts))
    check("text_to_html keeps line breaks", text_to_html("a\nb") == "a<br>b")


def test_mac_scripts_compile():
    print("--- generated AppleScript compiles against Outlook's dictionary (macOS + Outlook only) ---")
    if sys.platform != "darwin" or not shutil.which("osacompile") or not os.path.isdir("/Applications/Microsoft Outlook.app"):
        print("  SKIP: needs macOS with Microsoft Outlook installed")
        return
    tmp = tempfile.mkdtemp()
    scripts = []
    for call in (lambda: server_mac.send_draft(draft_id="12"),
                 lambda: server_mac.move_email(entry_id="12", target_folder="Archive"),
                 lambda: server_mac.send_email(to="a@example.invalid", subject="s", body="", html_body="<p>x</p>"),
                 lambda: server_mac.reply_email(entry_id="12", body="x")):
        cap = Capture("1"); server_mac.bridge = cap
        asyncio.run(call()); scripts += cap.scripts
    for i, s in enumerate(scripts):
        p = subprocess.run(["osacompile", "-o", os.path.join(tmp, f"{i}.scpt"), "-e", SANITIZE_HANDLER + s],
                           capture_output=True, text=True, timeout=60)
        check(f"script {i} compiles", p.returncode == 0, p.stderr.strip()[:200])
    shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    test_unc_rejected_before_filesystem()
    test_safe_filename()
    test_mark_as_downloaded()
    test_mac_save_attachment()
    test_mac_send_draft_guard()
    test_mac_move_email_account()
    test_mac_bodies_escaped()
    test_mac_scripts_compile()
    print("=" * 50)
    print(f"{passed}/{total} checks passed")
    sys.exit(0 if passed == total else 1)
