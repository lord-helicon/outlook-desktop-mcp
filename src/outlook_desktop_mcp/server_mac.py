"""
Outlook Desktop MCP Server — macOS
====================================
Exposes Microsoft Outlook for Mac as an MCP server over stdio.
Uses AppleScript automation via osascript — no Microsoft Graph, no Entra app.
Just run this on macOS with Outlook open and you have a full email MCP server.

Entry point: python -m outlook_desktop_mcp (auto-detected on macOS)
"""
import sys
import json
import logging
import os
import re
import asyncio
import hashlib

from mcp.server.fastmcp import FastMCP

from outlook_desktop_mcp.applescript_bridge import AppleScriptBridge
from outlook_desktop_mcp.utils.applescript_helpers import (
    date_component_lines,
    escape,
    format_date,
    parse_date,
    is_builtin_folder,
    resolve_folder_ref,
    text_to_html,
    AS_DELIM,
    AS_RECORD_DELIM,
    DELIM,
    RECORD_DELIM,
)
from outlook_desktop_mcp.utils.output import compact_json_tools, set_server_version
from outlook_desktop_mcp.utils.attachments import (
    attachment_root,
    is_within,
    mark_as_downloaded,
    resolve_save_dir,
    safe_filename,
    unique_path,
)
from datetime import datetime, timedelta

# --- Logging (all to stderr, stdout is reserved for MCP JSON-RPC) ---

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("outlook_desktop_mcp")


def _bad_message_id(message_id: str, kind: str = "message") -> str | None:
    """Return an error if message_id is not a plain Outlook numeric id.

    The id is interpolated unquoted into AppleScript (`message id {id}`,
    `calendar event id {id}`, `task id {id}`), so anything other than ASCII
    digits would be evaluated as script. `kind` only affects the message.
    """
    s = message_id.strip()
    if re.fullmatch(r"ui-[0-9]+(?:-[0-9a-f]{8})?", s):
        return (
            f"Error: '{message_id}' came from UI scraping (New Outlook) and "
            f"cannot be used as a {kind} id; only read_email accepts ui-N ids."
        )
    if not (s.isascii() and s.isdigit()):
        return f"Error: '{message_id}' is not a valid Outlook {kind} id (expected digits)."
    return None

# --- MCP Server ---

mcp = FastMCP(
    "outlook-desktop-mcp-plus",
    instructions=(
        "This MCP server gives you full access to Microsoft Outlook on macOS "
        "via AppleScript automation. It can send emails, read inbox messages, "
        "search across folders, mark messages as read/unread, move messages "
        "between folders, reply to emails, manage calendar events, and "
        "manage tasks.\n\n"
        "All operations use the locally running Outlook app — no "
        "Microsoft Graph API, no Entra app registration, no OAuth tokens needed. "
        "The user's existing Outlook session handles all authentication.\n\n"
        "PREREQUISITE: Microsoft Outlook for Mac must be running.\n\n"
        "NOTE: entry_id values on macOS are numeric IDs (not hex strings like "
        "on Windows). They identify items within their folder context.\n\n"
        "AVAILABLE TOOL CATEGORIES:\n"
        "- Email: send, list, read, search, reply, mark read/unread, move\n"
        "- Calendar: list events, create appointments/meetings, update, delete, "
        "search events\n"
        "- Tasks: create, list, complete, delete to-do items\n"
        "- Attachments: list and save attachments\n"
        "- Folders: list folder hierarchy"
    ),
)

# Every tool result leaves the server as compact JSON (see utils/output.py).
compact_json_tools(mcp)
set_server_version(mcp)

bridge = AppleScriptBridge()


# --- Helper: truncate long text ---

def _truncate(text: str, max_length: int = 5000) -> str:
    if len(text) <= max_length:
        return text
    return text[:max_length] + "\n... [truncated]"


def _clean(value: str) -> str:
    """Replace AppleScript's 'missing value' with empty string."""
    v = value.strip()
    return "" if v == "missing value" else v


_DATE_HINT = "Use YYYY-MM-DD or YYYY-MM-DD HH:MM (DD.MM.YYYY [HH:MM] is also accepted)."


def _parse_when(value: str, field: str):
    """Parse a caller-supplied date/time. Returns (datetime, None) or (None, error).

    Accepts ISO 8601 ("2026-12-01", "2026-12-01 14:00", "2026-12-01T14:00:00")
    and the unambiguous dotted day-first form ("01.12.2026", "01.12.2026 14:00").
    Slash forms are rejected: 01/12/2026 means different days in different locales.
    """
    text = (value or "").strip()
    try:
        return datetime.fromisoformat(text), None
    except ValueError:
        pass
    for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S", "%d.%m.%Y"):
        try:
            return datetime.strptime(text, fmt), None
        except ValueError:
            continue
    return None, f"Error: {field} '{value}' is not a valid date. {_DATE_HINT}"


def _bad_count(name: str, value, minimum: int = 1) -> str | None:
    """Error text if a count/limit argument is not an int >= minimum."""
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        return f"Error: {name} must be a whole number of at least {minimum} (got {value!r})."
    return None


def _priority_name(value: str) -> str:
    """'priority high' (AppleScript enum text) -> 'high'."""
    v = _clean(value)
    return v[len("priority "):] if v.startswith("priority ") else v


async def _no_mail_account_error(action: str) -> str | None:
    """Refuse mail-writing tools when Outlook has no AppleScript-visible account.

    New Outlook accepts `make new outgoing message` and returns an id, but the
    message is never stored or sent. Without this check save_draft reported
    success for drafts that do not exist (found in a live test on 16.115).
    """
    try:
        if (await bridge.run(_ACCOUNT_COUNT_SCRIPT)).strip() == "0":
            return (
                f"Error: cannot {action}. Outlook has no mail account that AppleScript "
                "can see (New Outlook keeps Microsoft 365 mail in the cloud). New Outlook "
                "accepts the request but never saves or sends the message. Switch Outlook "
                "to Legacy Outlook to use this tool."
            )
    except Exception:
        pass  # can't tell: let the tool run as before
    return None


async def _run_with_fallback(fast_script: str, legacy_script: str) -> str:
    """Run the batched (fast) AppleScript; fall back to the per-item script.

    Batched property fetches (e.g. `subject of messages 1 thru N of folder`)
    retrieve a whole column in a single Apple Event instead of one round trip
    per property per item, which is what made large list calls exceed the
    script timeout. If this Outlook version rejects a batch form, retry once
    with the legacy per-item script. Timeouts are not retried — the legacy
    script is strictly slower.
    """
    try:
        return await bridge.run(fast_script)
    except RuntimeError as e:
        if "timed out" in str(e):
            raise
        logger.warning(
            "Batch AppleScript failed; falling back to per-item script: %s", e
        )
        return await bridge.run(legacy_script)


# --- UI Scraping for New Outlook for Mac ---
# New Outlook for Mac stores Exchange/M365 mailbox data in the cloud and
# does NOT expose it through the AppleScript `inbox` keyword (which only
# reaches the empty local "On My Computer" inbox). The only way to access
# Exchange messages is via macOS UI scripting (System Events), reading
# the message list table visible in the Outlook window.


# Element chains from `window 1` down to the message table. The layout
# differs between Outlook builds: New Outlook has four nested splitter groups
# and an extra `UI element` wrapper; older builds have three and no wrapper;
# 16.115 has three and no scroll area around the table.
# These are System Events class names (not localized UI strings), so they do
# not depend on the macOS language. Each variant is tried in order.
_UI_LAYOUTS = (
    # splitters: splitter group depth; wrappers: elements between `group 1` and
    # the table; scroll: whether the table sits in `scroll area 1`.
    {"name": "new", "splitters": 4, "wrappers": ("UI element 1",), "scroll": True},
    {"name": "legacy", "splitters": 3, "wrappers": (), "scroll": True},
    # Outlook 16.115 (New Outlook): the table is a direct child of `group 1`.
    {"name": "flat", "splitters": 3, "wrappers": (), "scroll": False},
)


UI_ACCESS_DENIED_MESSAGE = (
    "Outlook's mailbox is not visible to AppleScript (New Outlook keeps "
    "Microsoft 365 mail in the cloud), and reading it from the Outlook window "
    "needs macOS Accessibility permission, which is not granted. Open System "
    "Settings > Privacy & Security > Accessibility, turn on the app that runs "
    "this MCP server (for example Claude, or your terminal app), then restart "
    "that app. Alternatively switch Outlook to Legacy Outlook."
)


NO_ACCOUNT_NOTE = (
    "Outlook has no mail account that AppleScript can see, so this tool only "
    "reads the local 'On My Computer' items. New Outlook keeps Microsoft 365 "
    "mail, calendars and tasks in the cloud, out of AppleScript's reach. To use "
    "this tool with that account, switch Outlook to Legacy Outlook."
)

_ACCOUNT_COUNT_SCRIPT = (
    'tell application "Microsoft Outlook" to return ((count of exchange accounts) '
    '+ (count of imap accounts) + (count of pop accounts)) as text'
)


async def _results_json(results: list, what: str) -> str:
    """JSON for a list result; for an empty one, explain New Outlook if needed.

    In New Outlook, AppleScript sees no account at all, so every list is empty.
    Returning a bare [] then reads as "your mailbox/calendar is empty". When the
    list is empty and Outlook reports zero accounts, return a note instead.
    """
    if results:
        return json.dumps(results, indent=2, default=str)
    try:
        if (await bridge.run(_ACCOUNT_COUNT_SCRIPT)).strip() == "0":
            return f"No {what} found. {NO_ACCOUNT_NOTE}"
    except Exception:
        pass  # can't tell: fall back to the plain empty list
    return json.dumps([])


def _is_ui_access_denied(error: BaseException) -> bool:
    """True if System Events refused UI scripting (no Accessibility permission).

    osascript reports this as error -25211 ("osascript is not allowed
    assistive access"); the error number does not depend on the macOS language.
    """
    text = str(error)
    return "-25211" in text or "assistive access" in text


# Extra rows read per list: date group headers ("Today", ...) are rows too.
_UI_HEADER_ALLOWANCE = 5
# Rows scanned for unread_only in the UI fallback.
_UI_UNREAD_SCAN_ROWS = 50


def _ui_scope(layout: dict, inner: str, *, pane: bool = False) -> str:
    """Wrap `inner` in nested System Events `tell` blocks for a layout.

    pane=False targets the message table (`... group 1 > wrappers > scroll
    area 1 > table 1`); pane=True stops at the innermost splitter group,
    whose children are the list (group 1), header (group 2) and body (group 3).
    """
    chain = ["application \"System Events\"", "process \"Microsoft Outlook\"",
             "window 1"]
    chain += ["splitter group 1"] * layout["splitters"]
    if not pane:
        chain += ["group 1", *layout["wrappers"]]
        if layout.get("scroll", True):
            chain.append("scroll area 1")
        chain.append("table 1")
    lines = [f"tell {c}" for c in chain] + [inner] + ["end tell"] * len(chain)
    return "\n".join(lines)


async def _ui_list_messages(bridge_obj, count: int = 10, unread_only: bool = False) -> list[dict]:
    """Read visible inbox messages via UI scripting (System Events).

    This is the fallback for New Outlook for Mac where AppleScript's inbox
    keyword only sees the empty local mailbox.
    """
    # Read a few extra rows: date group headers take up rows too.
    max_rows = count + _UI_HEADER_ALLOWANCE
    if unread_only:
        # Unread messages can be anywhere in the list; scan a fixed window
        # (batched, ~0.06 s per row).
        max_rows = max(max_rows, _UI_UNREAD_SCAN_ROWS)
    # One Apple event for all rows (~0.06 s/row) instead of one per row
    # (~0.4 s/row). If the batch fails or its length differs from the row
    # count (a row without a cell), fall back to the per-row loop so row
    # numbers stay exact.
    rows_script = (
        'set rowCount to count of rows\n'
        'set maxRows to rowCount\n'
        f'if maxRows > {max_rows} then set maxRows to {max_rows}\n'
        'if maxRows < 1 then return ""\n'
        'set output to ""\n'
        'set batched to false\n'
        'try\n'
        '    set descList to description of UI element 1 of rows 1 thru maxRows\n'
        '    if (count of descList) = maxRows then\n'
        '        repeat with d in descList\n'
        f'            set output to output & my sf(contents of d) & {AS_RECORD_DELIM}\n'
        '        end repeat\n'
        '        set batched to true\n'
        '    end if\n'
        'end try\n'
        'if not batched then\n'
        '    set output to ""\n'
        '    repeat with i from 1 to maxRows\n'
        '        try\n'
        '            set cellDesc to description of UI element 1 of row i\n'
        f'            set output to output & my sf(cellDesc) & {AS_RECORD_DELIM}\n'
        '        on error\n'
        f'            set output to output & {AS_RECORD_DELIM}\n'
        '        end try\n'
        '    end repeat\n'
        'end if\n'
        'return output'
    )

    raw = ""
    last_error: Exception | None = None
    succeeded = False
    for layout in _UI_LAYOUTS:
        try:
            raw = await bridge_obj.run(_ui_scope(layout, rows_script))
            succeeded = True
        except Exception as e:  # wrong element path for this build -> try next
            if _is_ui_access_denied(e):
                raise  # no layout can work without Accessibility permission
            last_error = e
            continue
        if raw and raw.strip():
            break
    if not succeeded and last_error is not None:
        raise last_error
    if not raw or not raw.strip():
        return []

    results = []
    for idx, record in enumerate(raw.split(RECORD_DELIM), start=1):
        record = record.strip()
        if not record:
            continue
        # Cell description format uses `,` + 4+ spaces as major field
        # separators, while in-content commas have 0-1 trailing spaces.
        # Structure: [UNREAD_FLAG,]    SENDER, SUBJECT,     TIME,    [FLAGS,]
        fields = [f.strip() for f in re.split(r",\s{4,}", record)]
        if len(fields) < 2:
            # Date group headers ("Today", "Last week", ...) are rows too, but
            # have no wide-gap fields. Skip them; idx stays the real row number.
            continue

        is_unread = False
        has_attachment = False
        # Status tokens are locale-dependent. Match known tokens across
        # languages so the parser works regardless of macOS language.
        _UNREAD_TOKENS = {"Ulest", "Unread", "Non lu", "Nicht gelesen",
                          "No leído", "未読", "未读"}
        _ATTACHMENT_TOKENS = {"Har filer", "Has attachments", "Contient des fichiers",
                              "Hat Anlagen", "Tiene archivos adjuntos", "添付ファイルあり",
                              "有附件"}
        _SKIP_PREFIXES = ("Merket som", "Marked as", "Marqué comme",
                          "Markiert als", "Marcado como", "A ")
        _CATEGORY_TOKENS = {"Kategorisert", "Categorized", "Catégorisé",
                            "Kategorisiert", "Categorizado"}
        cleaned = []
        for f in fields:
            if not f:
                continue
            if f in _UNREAD_TOKENS:
                is_unread = True
                continue
            if any(tok in f for tok in _ATTACHMENT_TOKENS):
                has_attachment = True
                continue
            if f in _CATEGORY_TOKENS or any(f.startswith(p) for p in _SKIP_PREFIXES):
                continue
            cleaned.append(f)

        # cleaned is typically: [SENDER_AND_SUBJECT, TIME]
        # or [SENDER_AND_SUBJECT, TIME, extra...]
        # SENDER_AND_SUBJECT is: "Sender, Subject" (comma + 1 space)
        sender_subject = cleaned[0] if cleaned else ""
        time_str = cleaned[1] if len(cleaned) > 1 else ""
        # Strip trailing comma from time
        time_str = time_str.rstrip(",").strip()

        # Split sender from subject on first ", " (comma + single space)
        # Remove thread/unread count prefixes like "2 messages, " or
        # "1 unread message, " in any locale (pattern: digits + words + comma)
        ss = sender_subject
        ss = re.sub(r"^\d+\s+[\w\s]+,\s*", "", ss)
        # Split on first ", " to get sender and subject
        comma_pos = ss.find(", ")
        if comma_pos > 0:
            sender = ss[:comma_pos].strip()
            subject = ss[comma_pos + 2:].strip()
        else:
            sender = ""
            subject = ss.strip()

        results.append({
            "entry_id": f"ui-{idx}-{_ui_fingerprint(sender, subject, time_str)}",
            "subject": subject or "(could not parse subject)",
            "sender": "",
            "sender_name": sender,
            "received_time": time_str,
            "unread": is_unread,
            "has_attachments": has_attachment,
            "attachment_count": 1 if has_attachment else 0,
            "_source": "ui_scraping",
        })

    if unread_only:
        results = [r for r in results if r["unread"]]
    return results[:count]


_UI_PANE_SETTLE_SECONDS = 1.5

# How many rows read_email scans to find a message that moved since list_emails.
_UI_LOOKUP_ROWS = 50
_UI_ID_RE = re.compile(r"ui-([0-9]+)-([0-9a-f]{8})")


def _ui_fingerprint(sender: str, subject: str, time_str: str) -> str:
    """Short hash of a scraped row's stable fields.

    ui-N ids are row positions, which shift when mail arrives or the list is
    re-sorted. The fingerprint lets read_email check that row N still shows
    the message list_emails returned (and find it if it moved). Read state
    and flags are excluded because selecting a row can change them.
    """
    key = "\x1f".join((sender, subject, time_str)).encode("utf-8")
    return hashlib.sha1(key).hexdigest()[:8]


async def _ui_resolve_row(bridge_obj, row_number: int, fingerprint: str) -> int | None:
    """Current row of the message with this fingerprint, or None if gone.

    Prefers the original row; otherwise the matching row nearest to it.
    """
    rows = await _ui_list_messages(bridge_obj, max(row_number, _UI_LOOKUP_ROWS))
    matches = []
    for r in rows:
        m = _UI_ID_RE.fullmatch(r["entry_id"])
        if m and m.group(2) == fingerprint:
            matches.append(int(m.group(1)))
    if not matches:
        return None
    return min(matches, key=lambda n: abs(n - row_number))

async def _ui_search_inbox(query: str, count: int) -> list[dict] | None:
    """Window-read inbox rows whose subject or sender name contains `query`.

    Scans the first _UI_LOOKUP_ROWS rows. Returns None if the window could not
    be read; a missing Accessibility permission is raised to the caller.
    """
    try:
        rows = await _ui_list_messages(bridge, _UI_LOOKUP_ROWS)
    except Exception as e:
        if _is_ui_access_denied(e):
            raise
        return None
    q = (query or "").strip().lower()
    hits = [r for r in rows if q in r["subject"].lower() or q in r["sender_name"].lower()]
    return hits[:count]


async def _inbox_search_result(results: list, folder: str, query: str, count: int) -> str:
    """search_emails result; for an empty New Outlook inbox, search the window."""
    if results or (folder or "").lower().strip() != "inbox":
        return await _results_json(results, "matching emails")
    try:
        hits = await _ui_search_inbox(query, count)
    except Exception:
        return f"Error searching emails: {UI_ACCESS_DENIED_MESSAGE}"
    if hits is None:
        return await _results_json(results, "matching emails")
    return json.dumps(hits, indent=2, default=str)


_UI_READ_PANE_SCRIPT = """set emailSubject to ""
try
    set emailSubject to value of static text 1 of group 2
end try
-- Header fields are matched by text prefix, not index: buttons, CC lines and
-- presence indicators shift positions between messages.
set emailFrom to ""
set emailDate to ""
set emailTo to ""
try
    tell scroll area 1 of group 2
        repeat with hElem in (every UI element)
            try
                set hRole to role of hElem
                set hVal to ""
                try
                    set hVal to value of hElem as text
                end try
                set hName to ""
                try
                    set hName to name of hElem as text
                end try
                if hVal starts with "From: " then
                    set emailFrom to text 7 thru -1 of hVal
                else if hVal starts with "Recipients: " then
                    set emailTo to text 13 thru -1 of hVal
                else if hRole is "AXUnknown" and hName starts with "Sent on: " then
                    set emailDate to hVal
                end if
            end try
        end repeat
    end tell
end try
-- Body is an AXWebArea. `entire contents` must be materialized into a
-- variable first; iterating it inline yields nothing.
set emailBody to ""
try
    tell group 3 to tell group 1 to tell scroll area 1 to tell UI element 1
        set allElems to entire contents
        repeat with e in allElems
            try
                if role of e is "AXStaticText" then
                    set emailBody to emailBody & (value of e as text) & return
                end if
            end try
        end repeat
    end tell
end try
return my sf(emailSubject) & (character id 1) & my sf(emailFrom) & (character id 1) & my sf(emailDate) & (character id 1) & my sf(emailTo) & (character id 1) & my sf(emailBody)"""


async def _ui_read_email_by_row(bridge_obj, row_number: int) -> dict:
    """Read a message by selecting its row and scraping the reading pane.

    Note: UI ids are row positions at list time; if mail arrives in between,
    the row may now hold a different message. Header prefixes ("From: ",
    "Recipients: ", "Sent on: ") are English-only; subject and body are not.
    """
    select_script = (
        f'select row {row_number}\n'
        f'set selected of row {row_number} to true'
    )
    chosen = None
    last_error: Exception | None = None
    for layout in _UI_LAYOUTS:
        try:
            await bridge_obj.run(_ui_scope(layout, select_script))
            chosen = layout
            break
        except Exception as e:
            last_error = e
    if chosen is None:
        return {"error": f"Failed to select row {row_number}: {last_error}"}

    await asyncio.sleep(_UI_PANE_SETTLE_SECONDS)  # let the reading pane load

    try:
        raw = await bridge_obj.run(_ui_scope(chosen, _UI_READ_PANE_SCRIPT, pane=True))
    except Exception as e:
        return {"error": f"Failed to read email content: {e}"}
    if not raw:
        return {"error": "No data returned from reading pane"}
    parts = raw.split(DELIM)
    if len(parts) < 5:
        return {"error": f"Unexpected data format: {raw}"}
    return {
        "entry_id": f"ui-{row_number}",
        "subject": _clean(parts[0]),
        "sender": _clean(parts[1]),
        "sender_name": _clean(parts[1]),
        "received_time": _clean(parts[2]),
        "to": _clean(parts[3]),
        "cc": "",
        "body": _truncate(_clean(DELIM.join(parts[4:]))),
        "unread": False,
        "has_attachments": False,
        "attachment_count": 0,
        "_source": "ui_scraping_read",
    }


# =====================================================================
# TOOL 1: send_email
# =====================================================================

@mcp.tool()
async def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
) -> str:
    """Send an email using the user's Outlook account.

    Creates and sends an email immediately through the default Outlook profile.
    The email will appear in the user's Sent Items folder after sending.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
            Example: "alice@example.com" or "alice@example.com; bob@example.com"
        subject: The email subject line.
        body: The plain-text body of the email. If html_body is also provided,
            both are set and Outlook will prefer the HTML version.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body. When provided, Outlook renders
            the email as HTML. The plain-text body serves as fallback.

    Returns:
        A confirmation message with subject and recipients, or an error.
    """
    if err := await _no_mail_account_error("send email"):
        return err
    # Build recipient lines
    def _recipient_lines(addresses: str, kind: str) -> str:
        lines = ""
        for addr in addresses.split(";"):
            addr = addr.strip()
            if addr:
                lines += f'make new {kind} at newMsg with properties {{email address:{{address:"{escape(addr)}"}}}}\n'
        return lines

    to_lines = _recipient_lines(to, "to recipient")
    cc_lines = _recipient_lines(cc, "cc recipient") if cc else ""
    bcc_lines = _recipient_lines(bcc, "bcc recipient") if bcc else ""

    # `content` is Outlook for Mac's HTML body (there is no `html content`
    # property); plain text is HTML-escaped so markup in it stays literal.
    content_prop = f'content:"{escape(html_body)}"' if html_body else f'content:"{escape(text_to_html(body))}"'

    script = f'''tell application "Microsoft Outlook"
    set newMsg to make new outgoing message with properties {{subject:"{escape(subject)}", {content_prop}}}
    {to_lines}{cc_lines}{bcc_lines}
    send newMsg
end tell'''

    try:
        await bridge.run(script)
        return f"Email sent: '{subject}' to {to}"
    except Exception as e:
        return f"Error sending email: {e}"


# =====================================================================
# TOOL 1B: save_draft
# =====================================================================

@mcp.tool()
async def save_draft(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
) -> str:
    """Save an email as a draft in Outlook's Drafts folder without sending.

    Creates an outgoing message in Outlook and leaves it in the Drafts folder
    so the user can review, edit, and send it later from Outlook. The message
    is NOT sent — use send_email for that. This is the recommended tool when a
    workflow requires human review before delivery.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
            Example: "alice@example.com" or "alice@example.com; bob@example.com"
        subject: The email subject line.
        body: The plain-text body of the email. If html_body is also provided,
            both are set and Outlook will prefer the HTML version.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body. When provided, Outlook renders
            the email as HTML. The plain-text body serves as fallback.

    Returns:
        A confirmation message with subject and recipients, or an error.
    """
    if err := await _no_mail_account_error("save a draft"):
        return err
    # Build recipient lines (same dialect as send_email)
    def _recipient_lines(addresses: str, kind: str) -> str:
        lines = ""
        for addr in addresses.split(";"):
            addr = addr.strip()
            if addr:
                lines += f'make new {kind} at newMsg with properties {{email address:{{address:"{escape(addr)}"}}}}\n'
        return lines

    to_lines = _recipient_lines(to, "to recipient")
    cc_lines = _recipient_lines(cc, "cc recipient") if cc else ""
    bcc_lines = _recipient_lines(bcc, "bcc recipient") if bcc else ""

    # `content` is Outlook for Mac's HTML body (there is no `html content`
    # property); plain text is HTML-escaped so markup in it stays literal.
    content_prop = f'content:"{escape(html_body)}"' if html_body else f'content:"{escape(text_to_html(body))}"'

    # No `send newMsg` here — the outgoing message stays in Drafts.
    script = f'''tell application "Microsoft Outlook"
    set newMsg to make new outgoing message with properties {{subject:"{escape(subject)}", {content_prop}}}
    {to_lines}{cc_lines}{bcc_lines}
    return my sf(id of newMsg)
end tell'''

    try:
        draft_id = await bridge.run(script)
        return json.dumps({
            "status": "draft_saved",
            "entry_id": draft_id.strip(),
            "subject": subject,
            "to": to,
        }, indent=2, default=str)
    except Exception as e:
        return f"Error saving draft: {e}"


# =====================================================================
# TOOL 1C: list_drafts
# =====================================================================
#
# NOTE on macOS support: AppleScript exposes the "drafts" folder for
# legacy/IMAP/POP accounts. Modern "New Outlook for Mac" stores Exchange/M365
# data in the cloud and does not expose drafts through AppleScript — in that
# case the tool returns an empty list. The Windows COM build is recommended
# for full draft management.

@mcp.tool()
async def list_drafts(
    limit: int = 50,
    offset: int = 0,
) -> str:
    """List drafts currently saved in Outlook's Drafts folder.

    Returns a JSON array of draft summaries sorted by last-modified time
    (newest first). Each summary includes entry_id, subject, to, cc, bcc,
    last_modified, has_attachments, and a short body_preview (first ~200
    chars of the plain-text body).

    Args:
        limit: Maximum number of drafts to return. Default 50, capped at 200.
        offset: Number of newest drafts to skip before returning. Default 0.

    Returns:
        JSON array of draft summary objects.
    """
    if err := _bad_count("limit", limit):
        return err
    if err := _bad_count("offset", offset, 0):
        return err
    limit = min(max(1, limit), 200)
    offset = max(0, offset)
    end_index = offset + limit  # 1-indexed slice upper bound applied below

    script = f'''tell application "Microsoft Outlook"
    set folderRef to drafts
    set allMsgs to messages of folderRef
    set msgCount to count of allMsgs
    set startIdx to {offset + 1}
    set endIdx to {end_index}
    if endIdx > msgCount then set endIdx to msgCount
    set output to ""
    if startIdx > msgCount then return output
    repeat with i from startIdx to endIdx
        set m to item i of allMsgs
        set mid to id of m
        set msubject to subject of m
        set mto to ""
        try
            set recips to to recipients of m
            repeat with r in recips
                set mto to mto & address of r & "; "
            end repeat
        end try
        set mcc to ""
        try
            set recips to cc recipients of m
            repeat with r in recips
                set mcc to mcc & address of r & "; "
            end repeat
        end try
        set mbcc to ""
        try
            set recips to bcc recipients of m
            repeat with r in recips
                set mbcc to mbcc & address of r & "; "
            end repeat
        end try
        set mtime to ""
        try
            set mtime to (modification date of m)
        end try
        set mattcount to 0
        try
            set mattcount to count of attachments of m
        end try
        set mbody to ""
        try
            set mbody to plain text content of m
        end try
        set output to output & my sf(mid) & {AS_DELIM} & my sf(msubject) & {AS_DELIM} & my sf(mto) & {AS_DELIM} & my sf(mcc) & {AS_DELIM} & my sf(mbcc) & {AS_DELIM} & my sf(mtime) & {AS_DELIM} & my sf(mattcount) & {AS_DELIM} & my sf(mbody) & {AS_RECORD_DELIM}
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)
        results = []
        if raw:
            for record in raw.split(RECORD_DELIM):
                record = record.strip()
                if not record:
                    continue
                parts = record.split(DELIM, 7)
                if len(parts) < 8:
                    continue
                att_count = int(parts[6]) if parts[6].strip().isdigit() else 0
                body = _clean(parts[7])
                preview = body[:200]
                if len(body) > 200:
                    preview += "..."
                results.append({
                    "entry_id": parts[0].strip(),
                    "subject": parts[1].strip() or "(no subject)",
                    "to": parts[2].strip().rstrip("; "),
                    "cc": parts[3].strip().rstrip("; "),
                    "bcc": parts[4].strip().rstrip("; "),
                    "last_modified": _clean(parts[5]),
                    "has_attachments": att_count > 0,
                    "attachment_count": att_count,
                    "body_preview": preview,
                })
        return await _results_json(results, "drafts")
    except Exception as e:
        return f"Error listing drafts: {e}"


# =====================================================================
# TOOL 1D: read_draft
# =====================================================================

@mcp.tool()
async def read_draft(draft_id: str) -> str:
    """Read the full content of a specific draft.

    Retrieves complete draft details including the full body, all recipients
    (To/CC/BCC), attachment file names, and last-modified time. Use this to
    review a draft before calling send_draft.

    Args:
        draft_id: The numeric ID of the draft from list_drafts results.

    Returns:
        JSON object with full draft details.
    """
    if err := _bad_message_id(draft_id):
        return err
    draft_id = draft_id.strip()
    script = f'''tell application "Microsoft Outlook"
    set m to message id {draft_id}
    set mid to id of m
    set msubject to subject of m
    set mto to ""
    try
        set recips to to recipients of m
        repeat with r in recips
            set mto to mto & address of r & "; "
        end repeat
    end try
    set mcc to ""
    try
        set recips to cc recipients of m
        repeat with r in recips
            set mcc to mcc & address of r & "; "
        end repeat
    end try
    set mbcc to ""
    try
        set recips to bcc recipients of m
        repeat with r in recips
            set mbcc to mbcc & address of r & "; "
        end repeat
    end try
    set mtime to ""
    try
        set mtime to (modification date of m)
    end try
    set mattnames to ""
    try
        repeat with a in attachments of m
            set mattnames to mattnames & (name of a) & "; "
        end repeat
    end try
    set mbody to ""
    try
        set mbody to plain text content of m
    end try
    set mhtml to ""
    try
        set mhtml to content of m
    end try
    return my sf(mid) & {AS_DELIM} & my sf(msubject) & {AS_DELIM} & my sf(mto) & {AS_DELIM} & my sf(mcc) & {AS_DELIM} & my sf(mbcc) & {AS_DELIM} & my sf(mtime) & {AS_DELIM} & my sf(mattnames) & {AS_DELIM} & my sf(mbody) & {AS_DELIM} & my sf(mhtml)
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM, 8)
        if len(parts) < 9:
            return json.dumps({"error": "Failed to parse draft data"})
        att_raw = parts[6].strip().rstrip("; ")
        attachments = [a.strip() for a in att_raw.split(";") if a.strip()] if att_raw else []
        result = {
            "entry_id": parts[0].strip(),
            "subject": parts[1].strip() or "(no subject)",
            "to": parts[2].strip().rstrip("; "),
            "cc": parts[3].strip().rstrip("; "),
            "bcc": parts[4].strip().rstrip("; "),
            "last_modified": _clean(parts[5]),
            "attachments": attachments,
            "has_attachments": len(attachments) > 0,
            "attachment_count": len(attachments),
            "body": _clean(parts[7]),
            "html_body": _clean(parts[8]),
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error reading draft: {e}"


# =====================================================================
# TOOL 1E: send_draft
# =====================================================================

@mcp.tool()
async def send_draft(draft_id: str) -> str:
    """Send an existing draft from Outlook's Drafts folder.

    Looks up the draft by ID and dispatches it via AppleScript `send`.
    After a successful send, Outlook moves the message from Drafts to Sent
    Items automatically (standard Outlook behavior).

    Args:
        draft_id: The numeric ID of the draft to send. Get this from
            list_drafts results.

    Returns:
        JSON object confirming the send, or an error.
    """
    if err := _bad_message_id(draft_id):
        return err
    if err := await _no_mail_account_error("send a draft"):
        return err
    draft_id = draft_id.strip()
    script = f'''tell application "Microsoft Outlook"
    set m to message id {draft_id}
    if class of m is not outgoing message then error "Message " & {draft_id} & " is not a draft (not an outgoing message)."
    if was sent of m then error "Message " & {draft_id} & " has already been sent; it is not a draft."
    set msubject to subject of m
    set mto to ""
    try
        set recips to to recipients of m
        repeat with r in recips
            set mto to mto & address of r & "; "
        end repeat
    end try
    send m
    return my sf(msubject) & {AS_DELIM} & my sf(mto)
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM, 1)
        subject = parts[0].strip() if parts else ""
        to = parts[1].strip().rstrip("; ") if len(parts) > 1 else ""
        return json.dumps({
            "status": "sent",
            "subject": subject or "(no subject)",
            "to": to,
            "message": f"Draft sent: '{subject}' to {to}",
        }, indent=2, default=str)
    except Exception as e:
        return f"Error sending draft: {e}"


# =====================================================================
# TOOL 2: list_emails
# =====================================================================

@mcp.tool()
async def list_emails(
    folder: str = "inbox",
    count: int = 10,
    unread_only: bool = False,
) -> str:
    """List recent emails from a specified Outlook folder.

    Returns a JSON array of email summaries sorted by received time (newest
    first). Each summary includes entry_id, subject, sender, sender_name,
    received_time, unread status, and attachment info.

    Use the entry_id from results to read full content with read_email,
    or to perform actions like mark_as_read, move_email, or reply_email.

    Args:
        folder: The folder to list. Case-insensitive names: "inbox" (default),
            "sent"/"sentmail", "drafts", "deleted"/"trash", "junk"/"spam",
            "outbox", or any custom folder name visible in list_folders output.
        count: Maximum number of emails to return. Default 10, max recommended 50.
        unread_only: If true, only return unread emails. Default false.
            In New Outlook (inbox read from the window) only the first 50
            rows are scanned.

    Returns:
        JSON array of email summary objects.
    """
    if err := _bad_count("count", count):
        return err
    folder_ref = resolve_folder_ref(folder)
    unread_filter = ' whose is read is false' if unread_only else ''

    # Batched property fetches: one Apple Event per property (whole column)
    # instead of one per property per message. For unread_only the whose
    # filter runs inside Outlook and the result is truncated locally.
    if unread_only:
        msg_spec = "(every message of folderRef whose is read is false)"
    else:
        msg_spec = "messages 1 thru maxCount of folderRef"

    fast_script = f'''tell application "Microsoft Outlook"
    set folderRef to {folder_ref}
    set msgCount to count of (messages of folderRef{unread_filter})
    set maxCount to {count}
    if msgCount < maxCount then set maxCount to msgCount
    if maxCount is 0 then return ""
    set idList to id of {msg_spec}
    set subjectList to subject of {msg_spec}
    set timeList to time received of {msg_spec}
    set readList to is read of {msg_spec}
    set senderList to {{}}
    try
        set senderList to sender of {msg_spec}
    end try
    set attLists to {{}}
    try
        set attLists to attachments of {msg_spec}
    end try
    set output to ""
    repeat with i from 1 to maxCount
        set msubject to ""
        try
            set msubject to item i of subjectList
        end try
        set msender to ""
        try
            set msender to address of (item i of senderList)
        end try
        set msenderName to ""
        try
            set msenderName to name of (item i of senderList)
        end try
        set mtime to ""
        try
            set mtime to (item i of timeList)
        end try
        set misread to false
        try
            set misread to item i of readList
        end try
        set mattcount to 0
        try
            set mattcount to count of (item i of attLists)
        end try
        set output to output & my sf(item i of idList) & {AS_DELIM} & my sf(msubject) & {AS_DELIM} & my sf(msender) & {AS_DELIM} & my sf(msenderName) & {AS_DELIM} & my sf(mtime) & {AS_DELIM} & my sf(misread) & {AS_DELIM} & my sf(mattcount) & {AS_RECORD_DELIM}
    end repeat
    return output
end tell'''

    legacy_script = f'''tell application "Microsoft Outlook"
    set folderRef to {folder_ref}
    set allMsgs to messages of folderRef{unread_filter}
    set msgCount to count of allMsgs
    set maxCount to {count}
    if msgCount < maxCount then set maxCount to msgCount
    set output to ""
    repeat with i from 1 to maxCount
        set m to item i of allMsgs
        set mid to id of m
        set msubject to subject of m
        set msender to ""
        try
            set msender to address of sender of m
        end try
        set msenderName to ""
        try
            set msenderName to name of sender of m
        end try
        set mtime to time received of m
        set misread to is read of m
        set mattcount to 0
        try
            set mattcount to count of attachments of m
        end try
        set output to output & my sf(mid) & {AS_DELIM} & my sf(msubject) & {AS_DELIM} & my sf(msender) & {AS_DELIM} & my sf(msenderName) & {AS_DELIM} & my sf(mtime) & {AS_DELIM} & my sf(misread) & {AS_DELIM} & my sf(mattcount) & {AS_RECORD_DELIM}
    end repeat
    return output
end tell'''

    try:
        raw = await _run_with_fallback(fast_script, legacy_script)

        results = []
        if raw:
            for record in raw.split(RECORD_DELIM):
                record = record.strip()
                if not record:
                    continue
                parts = record.split(DELIM)
                if len(parts) < 7:
                    continue
                att_count = int(parts[6]) if parts[6].strip().isdigit() else 0
                results.append({
                    "entry_id": parts[0].strip(),
                    "subject": parts[1].strip() or "(no subject)",
                    "sender": parts[2].strip(),
                    "sender_name": parts[3].strip(),
                    "received_time": _clean(parts[4]),
                    "unread": parts[5].strip().lower() != "true",  # is_read -> unread
                    "has_attachments": att_count > 0,
                    "attachment_count": att_count,
                })

        # Fallback: New Outlook for Mac keeps Exchange messages outside the
        # AppleScript-visible mailbox. If the standard query returned nothing
        # for the inbox, try reading the visible message list via UI scripting.
        if not results and folder.lower().strip() == "inbox":
            try:
                results = await _ui_list_messages(bridge, count, unread_only)
                # The window was read, so an empty list is a real answer
                # (e.g. no unread mail), not a sign of a hidden mailbox.
                return json.dumps(results, indent=2, default=str)
            except Exception as e:
                if _is_ui_access_denied(e):
                    return f"Error listing emails: {UI_ACCESS_DENIED_MESSAGE}"
                # Other UI scraping failures (no window, unknown layout):
                # keep returning the empty AppleScript result.

        return await _results_json(results, "emails")
    except Exception as e:
        return f"Error listing emails: {e}"


# =====================================================================
# TOOL 3: read_email
# =====================================================================

@mcp.tool()
async def read_email(
    entry_id: str = "",
    subject_search: str = "",
    folder: str = "inbox",
) -> str:
    """Read the full content of a specific email.

    Retrieves complete email details including body text, recipients, CC,
    and metadata. Provide EITHER entry_id (preferred, exact match) OR
    subject_search (finds most recent match by subject substring).

    Args:
        entry_id: The numeric ID of the email. Most reliable way to identify
            a specific email. Get this from list_emails or search_emails results.
        subject_search: Alternative to entry_id. A case-insensitive substring
            to search for in email subjects. Returns the most recent match.
        folder: Folder to search when using subject_search. Ignored when
            entry_id is provided. Default "inbox".

    Returns:
        JSON object with full email details (entry_id, subject, sender,
        sender_name, received_time, unread, to, cc, body, attachment info).
    """
    if entry_id:
        entry_id = entry_id.strip()
        # UI-scraped ids ("ui-N-<fingerprint>", from the New Outlook fallback
        # in list_emails) are row positions, not Outlook ids: read via the
        # pane after checking the row still holds the same message.
        ui_match = re.fullmatch(r"ui-([0-9]+)(?:-([0-9a-f]{8}))?", entry_id)
        if ui_match and int(ui_match.group(1)) >= 1:
            row_number, fingerprint = int(ui_match.group(1)), ui_match.group(2)
            if not fingerprint:
                return json.dumps({"error": (
                    f"'{entry_id}' has no row fingerprint, so it could open the "
                    "wrong message. Call list_emails again and use the new id."
                )})
            try:
                current = await _ui_resolve_row(bridge, row_number, fingerprint)
                if current is None:
                    result = {"error": (
                        f"The message for '{entry_id}' is no longer in the first "
                        f"{max(row_number, _UI_LOOKUP_ROWS)} rows of the message "
                        "list. Call list_emails again."
                    )}
                else:
                    result = await _ui_read_email_by_row(bridge, current)
                    if "error" not in result:
                        result["entry_id"] = f"ui-{current}-{fingerprint}"
            except Exception as e:
                result = {"error": UI_ACCESS_DENIED_MESSAGE if _is_ui_access_denied(e)
                          else f"Failed to read UI email: {e}"}
            if "error" in result:
                return json.dumps(result)
            return json.dumps(result, indent=2, default=str)
        if err := _bad_message_id(entry_id):
            return err
        folder_ref = resolve_folder_ref(folder)
        script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set mid to id of m
    set msubject to subject of m
    set msender to ""
    try
        set msender to address of sender of m
    end try
    set msenderName to ""
    try
        set msenderName to name of sender of m
    end try
    set mtime to time received of m
    set misread to is read of m
    set mattcount to 0
    try
        set mattcount to count of attachments of m
    end try
    set mto to ""
    try
        set recips to to recipients of m
        repeat with r in recips
            set mto to mto & address of r & "; "
        end repeat
    end try
    set mcc to ""
    try
        set recips to cc recipients of m
        repeat with r in recips
            set mcc to mcc & address of r & "; "
        end repeat
    end try
    set mbody to ""
    try
        set mbody to plain text content of m
    end try
    return my sf(mid) & {AS_DELIM} & my sf(msubject) & {AS_DELIM} & my sf(msender) & {AS_DELIM} & my sf(msenderName) & {AS_DELIM} & my sf(mtime) & {AS_DELIM} & my sf(misread) & {AS_DELIM} & my sf(mattcount) & {AS_DELIM} & my sf(mto) & {AS_DELIM} & my sf(mcc) & {AS_DELIM} & my sf(mbody)
end tell'''
    elif subject_search:
        folder_ref = resolve_folder_ref(folder)
        safe_query = escape(subject_search)
        script = f'''tell application "Microsoft Outlook"
    set folderRef to {folder_ref}
    set matchMsgs to messages of folderRef whose subject contains "{safe_query}"
    if (count of matchMsgs) = 0 then return "NOT_FOUND"
    set m to item 1 of matchMsgs
    set mid to id of m
    set msubject to subject of m
    set msender to ""
    try
        set msender to address of sender of m
    end try
    set msenderName to ""
    try
        set msenderName to name of sender of m
    end try
    set mtime to time received of m
    set misread to is read of m
    set mattcount to 0
    try
        set mattcount to count of attachments of m
    end try
    set mto to ""
    try
        set recips to to recipients of m
        repeat with r in recips
            set mto to mto & address of r & "; "
        end repeat
    end try
    set mcc to ""
    try
        set recips to cc recipients of m
        repeat with r in recips
            set mcc to mcc & address of r & "; "
        end repeat
    end try
    set mbody to ""
    try
        set mbody to plain text content of m
    end try
    return my sf(mid) & {AS_DELIM} & my sf(msubject) & {AS_DELIM} & my sf(msender) & {AS_DELIM} & my sf(msenderName) & {AS_DELIM} & my sf(mtime) & {AS_DELIM} & my sf(misread) & {AS_DELIM} & my sf(mattcount) & {AS_DELIM} & my sf(mto) & {AS_DELIM} & my sf(mcc) & {AS_DELIM} & my sf(mbody)
end tell'''
    else:
        return json.dumps({"error": "Provide either entry_id or subject_search"})

    try:
        raw = await bridge.run(script)
        if raw == "NOT_FOUND":
            # New Outlook: the AppleScript inbox is empty; look in the window.
            if subject_search and (folder or "").lower().strip() == "inbox":
                try:
                    hits = await _ui_search_inbox(subject_search, 1)
                except Exception:
                    return json.dumps({"error": UI_ACCESS_DENIED_MESSAGE})
                if hits:
                    m = _UI_ID_RE.fullmatch(hits[0]["entry_id"])
                    result = await _ui_read_email_by_row(bridge, int(m.group(1)))
                    if "error" in result:
                        return json.dumps(result)
                    result["entry_id"] = hits[0]["entry_id"]
                    return json.dumps(result, indent=2, default=str)
            return json.dumps({"error": f"No email found matching '{subject_search}'"})

        parts = raw.split(DELIM, 9)  # max 10 parts
        if len(parts) < 10:
            return json.dumps({"error": "Failed to parse email data"})

        att_count = int(parts[6].strip()) if parts[6].strip().isdigit() else 0
        result = {
            "entry_id": parts[0].strip(),
            "subject": parts[1].strip() or "(no subject)",
            "sender": parts[2].strip(),
            "sender_name": parts[3].strip(),
            "received_time": _clean(parts[4]),
            "unread": parts[5].strip().lower() != "true",
            "has_attachments": att_count > 0,
            "attachment_count": att_count,
            "to": parts[7].strip(),
            "cc": parts[8].strip(),
            "body": _truncate(_clean(parts[9])),
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error reading email: {e}"


# =====================================================================
# TOOL 4: mark_as_read
# =====================================================================

@mcp.tool()
async def mark_as_read(entry_id: str) -> str:
    """Mark a specific email as read in Outlook.

    Changes the unread status to read, same as clicking on an email in Outlook.
    The change is persisted immediately and synced to the server.

    Args:
        entry_id: The numeric ID of the email. Get this from list_emails
            or search_emails results.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    if err := _bad_message_id(entry_id):
        return err
    entry_id = entry_id.strip()
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set is read of m to true
    return subject of m
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Marked as read: '{subject}'"
    except Exception as e:
        return f"Error marking email as read: {e}"


# =====================================================================
# TOOL 5: mark_as_unread
# =====================================================================

@mcp.tool()
async def mark_as_unread(entry_id: str) -> str:
    """Mark a specific email as unread in Outlook.

    Restores a previously read email to unread status. Useful for flagging
    emails that need follow-up attention. Persisted immediately.

    Args:
        entry_id: The numeric ID of the email. Get this from list_emails
            or search_emails results.

    Returns:
        Confirmation message with the email subject, or an error.
    """
    if err := _bad_message_id(entry_id):
        return err
    entry_id = entry_id.strip()
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set is read of m to false
    return subject of m
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Marked as unread: '{subject}'"
    except Exception as e:
        return f"Error marking email as unread: {e}"


# =====================================================================
# TOOL 6: move_email
# =====================================================================

@mcp.tool()
async def move_email(
    entry_id: str,
    target_folder: str = "archive",
) -> str:
    """Move an email to a different Outlook folder.

    Moves the specified email from its current location to the target folder.
    IMPORTANT: After moving, the email gets a NEW entry_id — the old one
    becomes invalid.

    Args:
        entry_id: The numeric ID of the email to move.
        target_folder: Destination folder name. Default is "archive". Supports
            same names as list_emails: "inbox", "sent", "deleted"/"trash",
            "drafts", "junk"/"spam", or any custom folder name.

    Returns:
        Confirmation with email subject and destination, or an error.
    """
    dest_ref = resolve_folder_ref(target_folder)
    if err := _bad_message_id(entry_id):
        return err
    entry_id = entry_id.strip()
    # Custom folder names are looked up by name across all accounts; refuse a
    # match in another account so a same-named folder in a shared/co-owned
    # account can never receive this message. Built-in folders are unchanged.
    account_check = "" if is_builtin_folder(target_folder) else (
        f'''
    set destFolder to {dest_ref}
    if (id of account of destFolder) is not (id of account of m) or (class of account of destFolder) is not (class of account of m) then error "Folder {escape(target_folder)} belongs to a different account than the message; not moved."'''
    )
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set msubject to subject of m{account_check}
    move m to {dest_ref}
    return my sf(msubject)
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Moved '{subject}' to {target_folder}"
    except Exception as e:
        return f"Error moving email: {e}"


# =====================================================================
# TOOL 7: reply_email
# =====================================================================

@mcp.tool()
async def reply_email(
    entry_id: str,
    body: str,
    reply_all: bool = False,
) -> str:
    """Reply to an email in Outlook.

    Creates and sends a reply, preserving the original message thread.
    Use reply_all=True to reply to all recipients (sender + CC list).

    Args:
        entry_id: The numeric ID of the email to reply to.
        body: The reply message text. Prepended above the original message
            in the email thread.
        reply_all: If true, reply to all recipients (sender + all CC/To).
            If false (default), reply only to the sender.

    Returns:
        Confirmation indicating the reply was sent, or an error.
    """
    reply_cmd = "reply all to" if reply_all else "reply to"
    if err := _bad_message_id(entry_id):
        return err
    if err := await _no_mail_account_error("reply to email"):
        return err
    entry_id = entry_id.strip()
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set msubject to subject of m
    set replyMsg to {reply_cmd} m
    set content of replyMsg to "{escape(text_to_html(body))}<br><br>" & content of replyMsg
    send replyMsg
    return msubject
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Reply sent to '{subject}' (reply_all={reply_all})"
    except Exception as e:
        return f"Error replying to email: {e}"


def _folder_listing_pass(level: int, max_depth: int) -> str:
    """Generate the AppleScript block that lists folders at one nesting
    level and recurses into subfolders up to max_depth.

    Level 1 reads from the account reference (`item k of acctRefs`);
    deeper levels read from the enclosing level's folder reference.
    Folder names accumulate into a "Parent/Child" path. Names and unread
    counts are fetched as batched columns; only the per-folder message
    count needs one Apple Event per folder.
    """
    parent = "(item k of acctRefs)" if level == 1 else f"(item i{level - 1} of refs{level - 1})"
    path = f"n{level}" if level == 1 else f'path{level - 1} & "/" & n{level}'
    deeper = _folder_listing_pass(level + 1, max_depth) if level < max_depth else ""
    return f'''
        try
            set refs{level} to mail folders of {parent}
            set names{level} to name of mail folders of {parent}
            set unread{level} to unread count of mail folders of {parent}
            repeat with i{level} from 1 to count of refs{level}
                set n{level} to item i{level} of names{level}
                if n{level} is not missing value then
                    set path{level} to {path}
                    set c{level} to 0
                    try
                        set c{level} to count of messages of (item i{level} of refs{level})
                    end try
                    set u{level} to 0
                    try
                        set uu{level} to item i{level} of unread{level}
                        if uu{level} is not missing value then set u{level} to uu{level}
                    end try
                    set output to output & my sf(acctName) & {AS_DELIM} & my sf(path{level}) & {AS_DELIM} & my sf(c{level}) & {AS_DELIM} & my sf(u{level}) & {AS_RECORD_DELIM}{deeper}
                end if
            end repeat
        end try'''


# =====================================================================
# TOOL 7B: create_draft_reply  (ported from upstream PR #21)
# =====================================================================

@mcp.tool()
async def create_draft_reply(
    entry_id: str,
    body: str = "",
    reply_all: bool = False,
    html_body: str = "",
) -> str:
    """Create a reply draft without sending it.

    Creates a reply or reply-all, puts the supplied text above the original
    message thread, and saves it to Drafts so the user can review it before
    sending. This is the safe alternative to reply_email. Send it later with
    send_draft.

    Args:
        entry_id: The numeric message ID of the email to reply to.
        body: Plain-text reply content. Used when html_body is not provided.
        reply_all: If true, reply to all recipients. Default false.
        html_body: Optional. HTML-formatted reply content.

    Returns:
        JSON with status, the draft's entry_id and subject, or an error.
    """
    if err := _bad_message_id(entry_id):
        return err
    if err := await _no_mail_account_error("create a draft reply"):
        return err
    if not body and not html_body:
        return "Error: Provide body or html_body."
    entry_id = entry_id.strip()

    reply_cmd = "reply all to" if reply_all else "reply to"
    if html_body:
        content_line = (
            f'set content of replyMsg to "{escape(html_body)}<br><br>" '
            f'& content of replyMsg'
        )
    else:
        content_line = (
            f'set content of replyMsg to "{escape(text_to_html(body))}<br><br>" '
            f'& content of replyMsg'
        )

    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set originalSubject to subject of m
    set replyMsg to {reply_cmd} m
    {content_line}
    save replyMsg
    set mid to id of replyMsg
    return my sf(mid) & {AS_DELIM} & my sf(originalSubject) & {AS_DELIM} & my sf(subject of replyMsg)
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM) if raw else []
        return json.dumps({
            "status": "draft_reply_created",
            "entry_id": parts[0].strip() if len(parts) > 0 else "",
            "original_subject": parts[1].strip() if len(parts) > 1 else "",
            "subject": parts[2].strip() if len(parts) > 2 else "",
            "reply_all": reply_all,
            "folder": "Drafts",
        }, indent=2, default=str)
    except Exception as e:
        return f"Error creating draft reply: {e}"


# =====================================================================
# TOOL 8: list_folders
# =====================================================================

@mcp.tool()
async def list_folders(max_depth: int = 2) -> str:
    """List all mail folders in the user's Outlook mailbox.

    Returns a JSON array showing the folder hierarchy with item counts,
    grouped per account (e.g. the Exchange mailbox vs the local
    "On My Computer" store, which often contains empty folders with the
    same names). Subfolders appear as "Parent/Child" paths. Use this to
    discover folder names for other tools (list_emails, move_email,
    search_emails) — pass just the folder's own name, not the full path.

    Args:
        max_depth: How many levels deep to recurse into subfolders.
            Default 2 (top-level plus one level of subfolders); values
            above 5 are capped at 5.

    Returns:
        JSON array of folder objects with account, name, item_count,
        and unread_count.
    """
    if err := _bad_count("max_depth", max_depth, 0):
        return err
    depth = max(1, min(int(max_depth), 5))
    listing_pass = _folder_listing_pass(1, depth)

    # Enumerate folders per account so identically named folders (the empty
    # local "On My Computer" Inbox vs the real Exchange Inbox) are
    # distinguishable, and skip the nameless account-root containers that
    # the flat `mail folders` list exposes as "missing value".
    fast_script = f'''tell application "Microsoft Outlook"
    set acctRefs to {{}}
    set acctNames to {{}}
    try
        repeat with a in exchange accounts
            set end of acctRefs to a
            set end of acctNames to name of a
        end repeat
    end try
    try
        repeat with a in imap accounts
            set end of acctRefs to a
            set end of acctNames to name of a
        end repeat
    end try
    try
        repeat with a in pop accounts
            set end of acctRefs to a
            set end of acctNames to name of a
        end repeat
    end try
    try
        set end of acctRefs to on my computer
        set end of acctNames to "On My Computer"
    end try
    set output to ""
    repeat with k from 1 to count of acctRefs
        set acctName to item k of acctNames
{listing_pass}
    end repeat
    return output
end tell'''

    legacy_script = f'''tell application "Microsoft Outlook"
    set allFolders to mail folders
    set output to ""
    repeat with f in allFolders
        set fname to name of f
        set fcount to count of messages of f
        set funread to unread count of f
        set output to output & my sf(fname) & {AS_DELIM} & my sf(fcount) & {AS_DELIM} & my sf(funread) & {AS_RECORD_DELIM}
    end repeat
    return output
end tell'''

    try:
        raw = await _run_with_fallback(fast_script, legacy_script)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) >= 4:
                account, name, count_s, unread_s = (p.strip() for p in parts[:4])
            elif len(parts) == 3:
                # Legacy flat record from the fallback script: no account.
                account = ""
                name, count_s, unread_s = (p.strip() for p in parts)
            else:
                continue
            if not _clean(name):
                continue
            results.append({
                "account": _clean(account),
                "name": name,
                "item_count": int(count_s) if count_s.isdigit() else 0,
                "unread_count": int(unread_s) if unread_s.isdigit() else 0,
            })
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error listing folders: {e}"


# =====================================================================
# TOOL 9: search_emails
# =====================================================================

@mcp.tool()
async def search_emails(
    query: str,
    folder: str = "inbox",
    count: int = 10,
) -> str:
    """Search for emails in Outlook using text search.

    Searches email subjects using Outlook's AppleScript filtering.
    Results include entry_id for further operations.

    Args:
        query: The search term (case-insensitive substring match on subject).
            Examples: "budget report", "meeting notes", "quarterly".
        folder: Folder to search in. Default "inbox". Supports same
            names as list_emails.
        count: Maximum results to return. Default 10.

    Returns:
        JSON array of matching email summaries, or an error.
    """
    if err := _bad_count("count", count):
        return err
    folder_ref = resolve_folder_ref(folder)
    safe_query = escape(query)

    msg_spec = f'(every message of folderRef whose subject contains "{safe_query}")'

    fast_script = f'''tell application "Microsoft Outlook"
    set folderRef to {folder_ref}
    set msgCount to count of {msg_spec}
    set maxCount to {count}
    if msgCount < maxCount then set maxCount to msgCount
    if maxCount is 0 then return ""
    set idList to id of {msg_spec}
    set subjectList to subject of {msg_spec}
    set timeList to time received of {msg_spec}
    set readList to is read of {msg_spec}
    set senderList to {{}}
    try
        set senderList to sender of {msg_spec}
    end try
    set attLists to {{}}
    try
        set attLists to attachments of {msg_spec}
    end try
    set output to ""
    repeat with i from 1 to maxCount
        set msubject to ""
        try
            set msubject to item i of subjectList
        end try
        set msender to ""
        try
            set msender to address of (item i of senderList)
        end try
        set msenderName to ""
        try
            set msenderName to name of (item i of senderList)
        end try
        set mtime to ""
        try
            set mtime to (item i of timeList)
        end try
        set misread to false
        try
            set misread to item i of readList
        end try
        set mattcount to 0
        try
            set mattcount to count of (item i of attLists)
        end try
        set output to output & my sf(item i of idList) & {AS_DELIM} & my sf(msubject) & {AS_DELIM} & my sf(msender) & {AS_DELIM} & my sf(msenderName) & {AS_DELIM} & my sf(mtime) & {AS_DELIM} & my sf(misread) & {AS_DELIM} & my sf(mattcount) & {AS_RECORD_DELIM}
    end repeat
    return output
end tell'''

    legacy_script = f'''tell application "Microsoft Outlook"
    set folderRef to {folder_ref}
    set matchMsgs to messages of folderRef whose subject contains "{safe_query}"
    set msgCount to count of matchMsgs
    set maxCount to {count}
    if msgCount < maxCount then set maxCount to msgCount
    set output to ""
    repeat with i from 1 to maxCount
        set m to item i of matchMsgs
        set mid to id of m
        set msubject to subject of m
        set msender to ""
        try
            set msender to address of sender of m
        end try
        set msenderName to ""
        try
            set msenderName to name of sender of m
        end try
        set mtime to time received of m
        set misread to is read of m
        set mattcount to 0
        try
            set mattcount to count of attachments of m
        end try
        set output to output & my sf(mid) & {AS_DELIM} & my sf(msubject) & {AS_DELIM} & my sf(msender) & {AS_DELIM} & my sf(msenderName) & {AS_DELIM} & my sf(mtime) & {AS_DELIM} & my sf(misread) & {AS_DELIM} & my sf(mattcount) & {AS_RECORD_DELIM}
    end repeat
    return output
end tell'''

    try:
        raw = await _run_with_fallback(fast_script, legacy_script)
        if not raw:
            return await _inbox_search_result([], folder, query, count)

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 7:
                continue
            att_count = int(parts[6].strip()) if parts[6].strip().isdigit() else 0
            results.append({
                "entry_id": parts[0].strip(),
                "subject": parts[1].strip() or "(no subject)",
                "sender": parts[2].strip(),
                "sender_name": parts[3].strip(),
                "received_time": _clean(parts[4]),
                "unread": parts[5].strip().lower() != "true",
                "has_attachments": att_count > 0,
                "attachment_count": att_count,
            })
        return await _inbox_search_result(results, folder, query, count)
    except Exception as e:
        return f"Error searching emails: {e}"


# =====================================================================
# CALENDAR TOOLS
# =====================================================================


# =====================================================================
# TOOL 10: list_events
# =====================================================================

@mcp.tool()
async def list_events(
    start_date: str = "",
    end_date: str = "",
    count: int = 20,
) -> str:
    """List upcoming calendar events from Outlook.

    Returns a JSON array of event summaries within a date range, sorted by
    start time. Each summary has entry_id, subject, start, end, duration,
    location, organizer, and attendee info.

    Use entry_id from results with get_event, update_event, or delete_event.

    Args:
        start_date: Start of date range in ISO 8601 format (e.g. "2026-02-25"
            or "2026-02-25 09:00"). Default: now.
        end_date: End of date range. Default: 7 days from start_date.
        count: Maximum number of events to return. Default 20.

    Returns:
        JSON array of event summary objects.
    """
    if err := _bad_count("count", count):
        return err
    start, err = _parse_when(start_date, "start_date") if start_date else (datetime.now(), None)
    if err:
        return err
    end, err = _parse_when(end_date, "end_date") if end_date else (start + timedelta(days=7), None)
    if err:
        return err
    if end < start:
        return "Error: end_date is before start_date."

    # The date range is applied server-side with a whose filter on start
    # time; the cap only bounds the payload if the range matches thousands
    # of events. Start/end come back as ISO 8601 («class isot») so Python
    # can re-filter, sort, and truncate deterministically.
    fetch_cap = max(count * 3, 100)

    evt_spec = ("(every calendar event whose start time ≥ rangeStart "
                "and start time ≤ rangeEnd)")

    fast_script = f'''{date_component_lines("rangeStart", start)}
{date_component_lines("rangeEnd", end)}
tell application "Microsoft Outlook"
    set evtCount to count of {evt_spec}
    set maxFetch to {fetch_cap}
    if evtCount < maxFetch then set maxFetch to evtCount
    if maxFetch is 0 then return ""
    set idList to id of {evt_spec}
    set subjectList to subject of {evt_spec}
    set startList to start time of {evt_spec}
    set endList to end time of {evt_spec}
    set alldayList to all day flag of {evt_spec}
    set locList to {{}}
    try
        set locList to location of {evt_spec}
    end try
    set orgList to {{}}
    try
        set orgList to organizer of {evt_spec}
    end try
end tell
set output to ""
repeat with i from 1 to maxFetch
    set esubject to ""
    try
        set esubject to item i of subjectList
    end try
    set estart to ""
    try
        set estart to (item i of startList)
    end try
    set eend to ""
    try
        set eend to (item i of endList)
    end try
    set elocation to ""
    try
        set elocation to item i of locList
    end try
    set eorganizer to ""
    try
        set eorganizer to item i of orgList
    end try
    set eallday to false
    try
        set eallday to item i of alldayList
    end try
    set output to output & my sf(item i of idList) & {AS_DELIM} & my sf(esubject) & {AS_DELIM} & my sf(estart) & {AS_DELIM} & my sf(eend) & {AS_DELIM} & my sf(elocation) & {AS_DELIM} & my sf(eorganizer) & {AS_DELIM} & my sf(eallday) & {AS_RECORD_DELIM}
end repeat
return output'''

    # Fallback if this Outlook version rejects the whose-date filter:
    # batched fetch of the first events in calendar order. The Python
    # date filter below still applies, so out-of-range events are dropped
    # rather than returned.
    legacy_script = f'''tell application "Microsoft Outlook"
    set evtCount to count of calendar events
    set maxFetch to {fetch_cap}
    if evtCount < maxFetch then set maxFetch to evtCount
    if maxFetch is 0 then return ""
    set idList to id of calendar events 1 thru maxFetch
    set subjectList to subject of calendar events 1 thru maxFetch
    set startList to start time of calendar events 1 thru maxFetch
    set endList to end time of calendar events 1 thru maxFetch
    set alldayList to all day flag of calendar events 1 thru maxFetch
    set locList to {{}}
    try
        set locList to location of calendar events 1 thru maxFetch
    end try
    set orgList to {{}}
    try
        set orgList to organizer of calendar events 1 thru maxFetch
    end try
    set output to ""
    repeat with i from 1 to maxFetch
        set esubject to ""
        try
            set esubject to item i of subjectList
        end try
        set estart to ""
        try
            set estart to (item i of startList)
        end try
        set eend to ""
        try
            set eend to (item i of endList)
        end try
        set elocation to ""
        try
            set elocation to item i of locList
        end try
        set eorganizer to ""
        try
            set eorganizer to item i of orgList
        end try
        set eallday to false
        try
            set eallday to item i of alldayList
        end try
        set output to output & my sf(item i of idList) & {AS_DELIM} & my sf(esubject) & {AS_DELIM} & my sf(estart) & {AS_DELIM} & my sf(eend) & {AS_DELIM} & my sf(elocation) & {AS_DELIM} & my sf(eorganizer) & {AS_DELIM} & my sf(eallday) & {AS_RECORD_DELIM}
    end repeat
    return output
end tell'''

    def _event_dt(value: str) -> datetime | None:
        for candidate in (value, parse_date(value)):
            try:
                return datetime.fromisoformat(candidate)
            except (ValueError, TypeError):
                continue
        return None

    try:
        raw = await _run_with_fallback(fast_script, legacy_script)
        if not raw:
            return await _results_json([], "events")

        events = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 7:
                continue
            events.append({
                "entry_id": parts[0].strip(),
                "subject": parts[1].strip() or "(no subject)",
                "start": parts[2].strip(),
                "end": parts[3].strip(),
                "location": _clean(parts[4]),
                "organizer": _clean(parts[5]),
                "all_day": parts[6].strip().lower() == "true",
            })

        # Re-filter, sort by start time, and truncate. Events whose start
        # can't be parsed are kept (never silently dropped) and sort last.
        filtered = []
        for ev in events:
            dt = _event_dt(ev["start"])
            if dt is None or start <= dt <= end:
                filtered.append((dt, ev))
        filtered.sort(key=lambda pair: pair[0] or datetime.max)
        results = [ev for _, ev in filtered[:count]]

        return await _results_json(results, "events")
    except Exception as e:
        return f"Error listing events: {e}"


# =====================================================================
# TOOL 11: get_event
# =====================================================================

@mcp.tool()
async def get_event(entry_id: str) -> str:
    """Read the full details of a specific calendar event.

    Retrieves complete event information including body/description,
    attendees, and recurrence status.

    Args:
        entry_id: The numeric ID of the event. Get this from list_events
            or search_events results.

    Returns:
        JSON object with full event details.
    """
    if err := _bad_message_id(entry_id, "event"):
        return err
    entry_id = entry_id.strip()
    script = f'''tell application "Microsoft Outlook"
    set e to calendar event id {entry_id}
    set eid to id of e
    set esubject to subject of e
    set estart to start time of e
    set eend to end time of e
    set elocation to ""
    try
        set elocation to location of e
    end try
    set eorganizer to ""
    try
        set eorganizer to organizer of e
    end try
    set eallday to all day flag of e
    set ebody to ""
    try
        set ebody to plain text content of e
    end try
    set eattendees to ""
    try
        set attList to attendees of e
        repeat with a in attList
            set eattendees to eattendees & address of a & "; "
        end repeat
    end try
    return my sf(eid) & {AS_DELIM} & my sf(esubject) & {AS_DELIM} & my sf(estart) & {AS_DELIM} & my sf(eend) & {AS_DELIM} & my sf(elocation) & {AS_DELIM} & my sf(eorganizer) & {AS_DELIM} & my sf(eallday) & {AS_DELIM} & my sf(ebody) & {AS_DELIM} & my sf(eattendees)
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM, 8)
        if len(parts) < 9:
            return json.dumps({"error": "Failed to parse event data"})

        result = {
            "entry_id": parts[0].strip(),
            "subject": parts[1].strip() or "(no subject)",
            "start": parts[2].strip(),
            "end": parts[3].strip(),
            "location": _clean(parts[4]),
            "organizer": _clean(parts[5]),
            "all_day": parts[6].strip().lower() == "true",
            "body": _truncate(_clean(parts[7])),
            "attendees": parts[8].strip(),
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error reading event: {e}"


# =====================================================================
# TOOL 12: create_event
# =====================================================================

@mcp.tool()
async def create_event(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    all_day: bool = False,
    reminder_minutes: int = 15,
) -> str:
    """Create a personal calendar appointment (no attendees).

    Creates and saves an appointment on the user's calendar. This is a
    personal event — no meeting invitations are sent. Use create_meeting
    instead if you need to invite attendees.

    Args:
        subject: The event title.
        start: Start time in ISO 8601 format. Examples: "2026-02-25 14:00",
            "2026-02-25T14:00:00". For all-day events, use just the date.
        end: End time in ISO 8601 format.
        location: Optional. Event location.
        body: Optional. Description or notes for the event.
        all_day: If true, creates an all-day event. Default false.
        reminder_minutes: Minutes before the event to show a reminder.
            Default 15. Set to 0 to disable reminder.

    Returns:
        Confirmation with event subject and entry_id, or an error.
    """
    start_dt, err = _parse_when(start, "start")
    if err:
        return err
    end_dt, err = _parse_when(end, "end")
    if err:
        return err
    if end_dt < start_dt:
        return "Error: end is before start."
    if isinstance(reminder_minutes, bool) or not isinstance(reminder_minutes, int) or reminder_minutes < 0:
        return f"Error: reminder_minutes must be 0 or more (got {reminder_minutes!r})."

    props = f'subject:"{escape(subject)}", start time:{format_date(start_dt)}, end time:{format_date(end_dt)}'
    if location:
        props += f', location:"{escape(location)}"'
    if body:
        props += f', content:"{escape(text_to_html(body))}"'
    if all_day:
        props += ', all day flag:true'

    script = f'''tell application "Microsoft Outlook"
    set newEvt to make new calendar event with properties {{{props}}}
    return my sf(id of newEvt) & {AS_DELIM} & my sf((subject of newEvt)) & {AS_DELIM} & my sf((start time of newEvt)) & {AS_DELIM} & my sf((end time of newEvt))
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM)
        result = {
            "status": "created",
            "entry_id": parts[0].strip() if len(parts) > 0 else "",
            "subject": parts[1].strip() if len(parts) > 1 else subject,
            "start": parts[2].strip() if len(parts) > 2 else start,
            "end": parts[3].strip() if len(parts) > 3 else end,
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error creating event: {e}"


# =====================================================================
# TOOL 13: create_meeting
# =====================================================================

@mcp.tool()
async def create_meeting(
    subject: str,
    start: str,
    end: str,
    required_attendees: str,
    location: str = "",
    body: str = "",
    optional_attendees: str = "",
) -> str:
    """Create a meeting and send invitations to attendees.

    Creates a calendar meeting and sends meeting requests to all specified
    attendees. Note: invitation delivery depends on Outlook's AppleScript
    meeting support on macOS.

    Args:
        subject: The meeting title.
        start: Start time in ISO 8601 format (e.g. "2026-02-25 14:00").
        end: End time in ISO 8601 format (e.g. "2026-02-25 15:00").
        required_attendees: Required attendee email addresses, separated by
            semicolons. Example: "alice@example.com; bob@example.com"
        location: Optional. Meeting location.
        body: Optional. Meeting description or agenda.
        optional_attendees: Optional. Optional attendee emails, separated
            by semicolons.

    Returns:
        Confirmation that the meeting was created.
    """
    if err := await _no_mail_account_error("create a meeting with invitations"):
        return err
    start_dt, err = _parse_when(start, "start")
    if err:
        return err
    end_dt, err = _parse_when(end, "end")
    if err:
        return err
    if end_dt < start_dt:
        return "Error: end is before start."

    props = f'subject:"{escape(subject)}", start time:{format_date(start_dt)}, end time:{format_date(end_dt)}'
    if location:
        props += f', location:"{escape(location)}"'
    if body:
        props += f', content:"{escape(text_to_html(body))}"'

    attendee_lines = ""
    for addr in required_attendees.split(";"):
        addr = addr.strip()
        if addr:
            attendee_lines += f'make new required attendee at newEvt with properties {{email address:{{address:"{escape(addr)}"}}}}\n'
    if optional_attendees:
        for addr in optional_attendees.split(";"):
            addr = addr.strip()
            if addr:
                attendee_lines += f'make new optional attendee at newEvt with properties {{email address:{{address:"{escape(addr)}"}}}}\n'

    script = f'''tell application "Microsoft Outlook"
    set newEvt to make new calendar event with properties {{{props}}}
    {attendee_lines}
    return my sf(id of newEvt)
end tell'''

    try:
        await bridge.run(script)
        return f"Meeting '{subject}' created with attendees: {required_attendees}"
    except Exception as e:
        return f"Error creating meeting: {e}"


# =====================================================================
# TOOL 14: update_event
# =====================================================================

@mcp.tool()
async def update_event(
    entry_id: str,
    subject: str = "",
    start: str = "",
    end: str = "",
    location: str = "",
    body: str = "",
) -> str:
    """Update an existing calendar event.

    Modifies properties of an appointment or meeting. Only the fields you
    provide will be updated — omitted fields remain unchanged.

    Args:
        entry_id: The numeric ID of the event to update.
        subject: Optional. New event title.
        start: Optional. New start time in ISO 8601 format.
        end: Optional. New end time in ISO 8601 format.
        location: Optional. New location.
        body: Optional. New description/notes.

    Returns:
        Confirmation with updated event details, or an error.
    """
    set_lines = ""
    if subject:
        set_lines += f'set subject of e to "{escape(subject)}"\n'
    start_dt = end_dt = None
    if start:
        start_dt, err = _parse_when(start, "start")
        if err:
            return err
        set_lines += f'set start time of e to {format_date(start_dt)}\n'
    if end:
        end_dt, err = _parse_when(end, "end")
        if err:
            return err
        set_lines += f'set end time of e to {format_date(end_dt)}\n'
    if start_dt and end_dt and end_dt < start_dt:
        return "Error: end is before start."
    if location:
        set_lines += f'set location of e to "{escape(location)}"\n'
    if body:
        set_lines += f'set content of e to "{escape(text_to_html(body))}"\n'

    if not set_lines:
        return json.dumps({"error": "No fields to update"})

    if err := _bad_message_id(entry_id, "event"):
        return err
    entry_id = entry_id.strip()
    script = f'''tell application "Microsoft Outlook"
    set e to calendar event id {entry_id}
    {set_lines}
    return my sf(id of e) & {AS_DELIM} & my sf((subject of e)) & {AS_DELIM} & my sf((start time of e)) & {AS_DELIM} & my sf((end time of e)) & {AS_DELIM} & my sf((location of e))
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM)
        result = {
            "status": "updated",
            "entry_id": parts[0].strip() if len(parts) > 0 else entry_id,
            "subject": parts[1].strip() if len(parts) > 1 else "",
            "start": parts[2].strip() if len(parts) > 2 else "",
            "end": parts[3].strip() if len(parts) > 3 else "",
            "location": _clean(parts[4]) if len(parts) > 4 else "",
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error updating event: {e}"


# =====================================================================
# TOOL 15: delete_event
# =====================================================================

@mcp.tool()
async def delete_event(entry_id: str) -> str:
    """Delete a calendar event.

    For personal appointments, the event is simply deleted. For meetings,
    the event is removed from your calendar.

    Args:
        entry_id: The numeric ID of the event to delete.

    Returns:
        Confirmation with the event subject, or an error.
    """
    if err := _bad_message_id(entry_id, "event"):
        return err
    entry_id = entry_id.strip()
    script = f'''tell application "Microsoft Outlook"
    set e to calendar event id {entry_id}
    set esubject to subject of e
    delete e
    return esubject
end tell'''

    try:
        subject = await bridge.run(script)
        return f"Event deleted: '{subject}'"
    except Exception as e:
        return f"Error deleting event: {e}"


# =====================================================================
# TOOL 16: search_events
# =====================================================================

@mcp.tool()
async def search_events(
    query: str,
    start_date: str = "",
    end_date: str = "",
    count: int = 10,
) -> str:
    """Search for calendar events by keyword.

    Searches event subjects within a date range.

    Args:
        query: The search term (case-insensitive substring match on subject).
            Examples: "standup", "review", "1:1".
        start_date: Start of search range in ISO 8601 format. Default: 30
            days ago.
        end_date: End of search range. Default: 30 days from now.
        count: Maximum results to return. Default 10.

    Returns:
        JSON array of matching event summaries.
    """
    if err := _bad_count("count", count):
        return err
    now = datetime.now()
    range_start, err = _parse_when(start_date, "start_date") if start_date else (now - timedelta(days=30), None)
    if err:
        return err
    range_end, err = _parse_when(end_date, "end_date") if end_date else (now + timedelta(days=30), None)
    if err:
        return err
    if range_end < range_start:
        return "Error: end_date is before start_date."
    safe_query = escape(query)

    evt_spec = f'(every calendar event whose subject contains "{safe_query}")'

    fast_script = f'''tell application "Microsoft Outlook"
    set evtCount to count of {evt_spec}
    set maxCount to {count}
    if evtCount < maxCount then set maxCount to evtCount
    if maxCount is 0 then return ""
    set idList to id of {evt_spec}
    set subjectList to subject of {evt_spec}
    set startList to start time of {evt_spec}
    set endList to end time of {evt_spec}
    set alldayList to all day flag of {evt_spec}
    set locList to {{}}
    try
        set locList to location of {evt_spec}
    end try
    set orgList to {{}}
    try
        set orgList to organizer of {evt_spec}
    end try
    set output to ""
    repeat with i from 1 to maxCount
        set esubject to ""
        try
            set esubject to item i of subjectList
        end try
        set estart to ""
        try
            set estart to (item i of startList)
        end try
        set eend to ""
        try
            set eend to (item i of endList)
        end try
        set elocation to ""
        try
            set elocation to item i of locList
        end try
        set eorganizer to ""
        try
            set eorganizer to item i of orgList
        end try
        set eallday to false
        try
            set eallday to item i of alldayList
        end try
        set output to output & my sf(item i of idList) & {AS_DELIM} & my sf(esubject) & {AS_DELIM} & my sf(estart) & {AS_DELIM} & my sf(eend) & {AS_DELIM} & my sf(elocation) & {AS_DELIM} & my sf(eorganizer) & {AS_DELIM} & my sf(eallday) & {AS_RECORD_DELIM}
    end repeat
    return output
end tell'''

    legacy_script = f'''tell application "Microsoft Outlook"
    set evts to calendar events whose subject contains "{safe_query}"
    set evtCount to count of evts
    set maxCount to {count}
    if evtCount < maxCount then set maxCount to evtCount
    set output to ""
    repeat with i from 1 to maxCount
        set e to item i of evts
        set eid to id of e
        set esubject to subject of e
        set estart to start time of e
        set eend to end time of e
        set elocation to ""
        try
            set elocation to location of e
        end try
        set eorganizer to ""
        try
            set eorganizer to organizer of e
        end try
        set eallday to all day flag of e
        set output to output & my sf(eid) & {AS_DELIM} & my sf(esubject) & {AS_DELIM} & my sf(estart) & {AS_DELIM} & my sf(eend) & {AS_DELIM} & my sf(elocation) & {AS_DELIM} & my sf(eorganizer) & {AS_DELIM} & my sf(eallday) & {AS_RECORD_DELIM}
    end repeat
    return output
end tell'''

    try:
        raw = await _run_with_fallback(fast_script, legacy_script)
        if not raw:
            return await _results_json([], "matching events")

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 7:
                continue
            results.append({
                "entry_id": parts[0].strip(),
                "subject": parts[1].strip() or "(no subject)",
                "start": parts[2].strip(),
                "end": parts[3].strip(),
                "location": _clean(parts[4]),
                "organizer": _clean(parts[5]),
                "all_day": parts[6].strip().lower() == "true",
            })

        # Keep events that overlap the range; unparseable starts are kept.
        def _dt(v):
            try:
                return datetime.fromisoformat(v)
            except ValueError:
                return None
        kept = []
        for ev in results:
            st, en = _dt(ev["start"]), _dt(ev["end"])
            if st is None or (st <= range_end and (en or st) >= range_start):
                kept.append((st, ev))
        kept.sort(key=lambda pair: pair[0] or datetime.max)
        results = [ev for _, ev in kept[:count]]
        return await _results_json(results, "matching events")
    except Exception as e:
        return f"Error searching events: {e}"


# =====================================================================
# TASK TOOLS
# =====================================================================

@mcp.tool()
async def list_tasks(
    include_completed: bool = False,
    count: int = 20,
) -> str:
    """List tasks from the Outlook Tasks folder.

    Returns a JSON array of task summaries. Each task includes entry_id,
    subject, due_date, and completion status.

    Args:
        include_completed: If true, include completed tasks. Default false
            (only pending/in-progress tasks).
        count: Maximum number of tasks to return. Default 20.

    Returns:
        JSON array of task summary objects.
    """
    if err := _bad_count("count", count):
        return err
    completed_filter = "" if include_completed else " whose todo flag is not completed"

    if include_completed:
        task_spec = "tasks 1 thru maxCount"
    else:
        task_spec = "(every task whose todo flag is not completed)"

    fast_script = f'''tell application "Microsoft Outlook"
    set taskCount to count of (tasks{completed_filter})
    set maxCount to {count}
    if taskCount < maxCount then set maxCount to taskCount
    if maxCount is 0 then return ""
    set idList to id of {task_spec}
    set nameList to name of {task_spec}
    set flagList to todo flag of {task_spec}
    set priList to priority of {task_spec}
    set dueList to {{}}
    try
        set dueList to due date of {task_spec}
    end try
    set output to ""
    repeat with i from 1 to maxCount
        set tname to ""
        try
            set tname to item i of nameList
        end try
        set tdue to ""
        try
            set tdue to (item i of dueList)
        end try
        set tflag to ""
        try
            set tflag to (item i of flagList) as text
        end try
        set tpriority to ""
        try
            set tpriority to (item i of priList) as text
        end try
        set output to output & my sf(item i of idList) & {AS_DELIM} & my sf(tname) & {AS_DELIM} & my sf(tdue) & {AS_DELIM} & my sf(tflag) & {AS_DELIM} & my sf(tpriority) & {AS_RECORD_DELIM}
    end repeat
    return output
end tell'''

    legacy_script = f'''tell application "Microsoft Outlook"
    set taskList to tasks{completed_filter}
    set taskCount to count of taskList
    set maxCount to {count}
    if taskCount < maxCount then set maxCount to taskCount
    set output to ""
    repeat with i from 1 to maxCount
        set t to item i of taskList
        set tid to id of t
        set tname to name of t
        set tdue to ""
        try
            set tdue to due date of t
        end try
        set tflag to todo flag of t
        set tpriority to priority of t
        set output to output & my sf(tid) & {AS_DELIM} & my sf(tname) & {AS_DELIM} & my sf(tdue) & {AS_DELIM} & my sf(tflag) & {AS_DELIM} & my sf(tpriority) & {AS_RECORD_DELIM}
    end repeat
    return output
end tell'''

    try:
        raw = await _run_with_fallback(fast_script, legacy_script)
        if not raw:
            return await _results_json([], "tasks")

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 5:
                continue
            results.append({
                "entry_id": parts[0].strip(),
                "subject": parts[1].strip() or "(no subject)",
                "due_date": _clean(parts[2]) or None,
                "complete": parts[3].strip() == "completed",
                "priority": _priority_name(parts[4]),
            })
        return await _results_json(results, "tasks")
    except Exception as e:
        return f"Error listing tasks: {e}"


@mcp.tool()
async def get_task(entry_id: str) -> str:
    """Read the full details of a specific task.

    Args:
        entry_id: The numeric ID of the task.

    Returns:
        JSON object with full task details including body.
    """
    if err := _bad_message_id(entry_id, "task"):
        return err
    entry_id = entry_id.strip()
    script = f'''tell application "Microsoft Outlook"
    set t to task id {entry_id}
    set tid to id of t
    set tname to name of t
    set tdue to ""
    try
        set tdue to due date of t
    end try
    set tflag to todo flag of t
    set tpriority to priority of t
    set tbody to ""
    try
        set tbody to plain text content of t
    end try
    set tstartdate to ""
    try
        set tstartdate to start date of t
    end try
    return my sf(tid) & {AS_DELIM} & my sf(tname) & {AS_DELIM} & my sf(tdue) & {AS_DELIM} & my sf(tflag) & {AS_DELIM} & my sf(tpriority) & {AS_DELIM} & my sf(tbody) & {AS_DELIM} & my sf(tstartdate)
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM, 6)
        if len(parts) < 7:
            return json.dumps({"error": "Failed to parse task data"})

        result = {
            "entry_id": parts[0].strip(),
            "subject": parts[1].strip() or "(no subject)",
            "due_date": _clean(parts[2]) or None,
            "complete": parts[3].strip() == "completed",
            "priority": _priority_name(parts[4]),
            "body": _truncate(_clean(parts[5])),
            "start_date": _clean(parts[6]) or None,
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error reading task: {e}"


@mcp.tool()
async def create_task(
    subject: str,
    body: str = "",
    due_date: str = "",
    importance: str = "normal",
) -> str:
    """Create a new task in Outlook.

    Args:
        subject: The task title.
        body: Optional. Task description or notes.
        due_date: Optional. Due date in ISO 8601 format (e.g. "2026-03-01").
        importance: Optional. "low", "normal" (default), or "high".

    Returns:
        Confirmation with task subject and entry_id.
    """
    imp_map = {"low": "priority low", "normal": "priority normal", "high": "priority high"}
    imp_key = (importance or "normal").strip().lower()
    if imp_key not in imp_map:
        return f"Error: importance must be 'low', 'normal' or 'high' (got {importance!r})."
    imp_val = imp_map[imp_key]

    props = f'name:"{escape(subject)}", priority:{imp_val}'
    if due_date:
        due_dt, err = _parse_when(due_date, "due_date")
        if err:
            return err
        props += f', due date:{format_date(due_dt)}'
    if body:
        props += f', content:"{escape(text_to_html(body))}"'

    script = f'''tell application "Microsoft Outlook"
    set newTask to make new task with properties {{{props}}}
    return my sf(id of newTask) & {AS_DELIM} & my sf((name of newTask))
end tell'''

    try:
        raw = await bridge.run(script)
        parts = raw.split(DELIM)
        result = {
            "status": "created",
            "entry_id": parts[0].strip() if len(parts) > 0 else "",
            "subject": parts[1].strip() if len(parts) > 1 else subject,
            "due_date": due_dt.isoformat() if due_date else None,
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error creating task: {e}"


@mcp.tool()
async def complete_task(entry_id: str) -> str:
    """Mark a task as complete.

    Sets the task status to complete.

    Args:
        entry_id: The numeric ID of the task.

    Returns:
        Confirmation with the task subject.
    """
    if err := _bad_message_id(entry_id, "task"):
        return err
    entry_id = entry_id.strip()
    script = f'''tell application "Microsoft Outlook"
    set t to task id {entry_id}
    set todo flag of t to completed
    return name of t
end tell'''

    try:
        name = await bridge.run(script)
        return f"Task completed: '{name}'"
    except Exception as e:
        return f"Error completing task: {e}"


@mcp.tool()
async def delete_task(entry_id: str) -> str:
    """Delete a task from Outlook.

    Args:
        entry_id: The numeric ID of the task to delete.

    Returns:
        Confirmation with the task subject.
    """
    if err := _bad_message_id(entry_id, "task"):
        return err
    entry_id = entry_id.strip()
    script = f'''tell application "Microsoft Outlook"
    set t to task id {entry_id}
    set tname to name of t
    delete t
    return tname
end tell'''

    try:
        name = await bridge.run(script)
        return f"Task deleted: '{name}'"
    except Exception as e:
        return f"Error deleting task: {e}"


# =====================================================================
# ATTACHMENT TOOLS
# =====================================================================

@mcp.tool()
async def list_attachments(entry_id: str) -> str:
    """List all attachments on an email.

    Args:
        entry_id: The numeric ID of the email to check for attachments.

    Returns:
        JSON array of attachment objects with index and filename.
    """
    if err := _bad_message_id(entry_id):
        return err
    entry_id = entry_id.strip()
    script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set attList to attachments of m
    set attCount to count of attList
    set output to ""
    repeat with i from 1 to attCount
        set a to item i of attList
        set aname to name of a
        set asize to file size of a
        set output to output & my sf(i) & {AS_DELIM} & my sf(aname) & {AS_DELIM} & my sf(asize) & {AS_RECORD_DELIM}
    end repeat
    return output
end tell'''

    try:
        raw = await bridge.run(script)
        if not raw:
            return json.dumps([])

        results = []
        for record in raw.split(RECORD_DELIM):
            record = record.strip()
            if not record:
                continue
            parts = record.split(DELIM)
            if len(parts) < 3:
                continue
            results.append({
                "index": int(parts[0].strip()) if parts[0].strip().isdigit() else 0,
                "filename": parts[1].strip(),
                "size": int(parts[2].strip()) if parts[2].strip().isdigit() else 0,
            })
        return json.dumps(results, indent=2, default=str)
    except Exception as e:
        return f"Error listing attachments: {e}"


@mcp.tool()
async def save_attachment(
    entry_id: str,
    attachment_index: int = 1,
    save_directory: str = "",
) -> str:
    """Save an attachment from an email to disk.

    Saves the attachment inside the configured attachment root (default
    ~/Downloads, or the OUTLOOK_MCP_ATTACHMENT_DIR environment variable).
    The file name is reduced to a safe single component, existing files are
    never overwritten (a suffix such as " (1)" is added), and the saved file
    is marked as downloaded from the internet (com.apple.quarantine).

    Args:
        entry_id: The numeric ID of the email containing the attachment.
        attachment_index: Which attachment to save (1-based index). Default 1.
            Use list_attachments to see available indices.
        save_directory: Optional folder inside the attachment root (absolute,
            or relative to the root). Default: the root itself.

    Returns:
        JSON with the saved file name and path, or an error.
    """
    if err := _bad_count("attachment_index", attachment_index):
        return err
    if err := _bad_message_id(entry_id):
        return err
    entry_id = entry_id.strip()
    root = attachment_root()
    target_dir, err = resolve_save_dir(save_directory, root)
    if err:
        return err

    # Step 1: read the attachment's name (chosen by the sender) without saving.
    name_script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set attList to attachments of m
    set attCount to count of attList
    if attCount < {attachment_index} then return "ERROR:Only " & attCount & " attachment(s)"
    return my sf(name of (item {attachment_index} of attList))
end tell'''
    try:
        raw = await bridge.run(name_script)
        if raw.startswith("ERROR:"):
            return raw
        os.makedirs(target_dir, exist_ok=True)
        save_path = unique_path(target_dir, safe_filename(raw))
        if not is_within(os.path.realpath(save_path), target_dir):
            return "Error: Attachment filename would escape the target directory."

        # Step 2: save to the exact path chosen here, never to a path built
        # from the sender's attachment name.
        save_script = f'''tell application "Microsoft Outlook"
    set m to message id {entry_id}
    set a to item {attachment_index} of (attachments of m)
    save a in "{escape(save_path)}"
end tell'''
        await bridge.run(save_script)
        if not os.path.isfile(save_path) or not is_within(os.path.realpath(save_path), target_dir):
            return f"Error: Outlook did not save the attachment to {save_path}."
        result = {
            "status": "saved",
            "filename": os.path.basename(save_path),
            "path": save_path,
            "marked_as_downloaded": mark_as_downloaded(save_path),
        }
        return json.dumps(result, indent=2, default=str)
    except Exception as e:
        return f"Error saving attachment: {e}"


# =====================================================================
# Entry point
# =====================================================================

def main():
    import asyncio

    async def _start():
        logger.info("Starting Outlook Desktop MCP server (macOS)...")
        await bridge.start()
        logger.info("AppleScript bridge ready. Starting MCP stdio transport...")

    asyncio.run(_start())
    try:
        mcp.run(transport="stdio")
    finally:
        bridge.stop()
