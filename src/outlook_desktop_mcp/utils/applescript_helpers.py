"""Helpers for building and parsing AppleScript safely."""
import re
from datetime import datetime


def escape(text: str) -> str:
    """Escape a string for safe embedding inside AppleScript double quotes.

    Handles backslashes, double quotes, and other special characters.
    """
    text = text.replace("\\", "\\\\")
    text = text.replace('"', '\\"')
    text = text.replace("\n", "\\n")
    text = text.replace("\r", "\\r")
    text = text.replace("\t", "\\t")
    return text


def text_to_html(text: str) -> str:
    """Plain text -> HTML that displays the same characters.

    Outlook for Mac's `content` property is HTML (see its AppleScript
    dictionary), so plain text written there must be HTML-escaped, or markup
    in it (e.g. quoted from a sender's email) would render as live HTML.
    """
    import html
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return html.escape(text, quote=False).replace("\n", "<br>")


def format_date(dt: datetime) -> str:
    """Convert a Python datetime to a locale-independent AppleScript date.

    AppleScript's ``date "..."`` string coercion parses the string according
    to the *system locale*, so handing it an ISO string such as
    ``date "2026-03-22 14:00:00"`` is mis-read on non-US locales. On es_ES,
    for example, AppleScript turns that into 16 September 2027 — the date is
    scrambled while only the time survives.

    To stay locale-independent we build the date by assigning its numeric
    components (which AppleScript never reinterprets) and wrap the whole thing
    in ``run script`` so the result is still a single inline expression that
    can be interpolated everywhere format_date() is used (event properties,
    ``set start time of e to ...``, task due dates, etc.). The day is reset to
    1 before setting year/month to avoid end-of-month rollover (e.g. assigning
    a month while the day is 31).
    """
    inner = (
        "set d to current date\\n"
        "set day of d to 1\\n"
        f"set year of d to {dt.year}\\n"
        f"set month of d to {dt.month}\\n"
        f"set day of d to {dt.day}\\n"
        f"set hours of d to {dt.hour}\\n"
        f"set minutes of d to {dt.minute}\\n"
        f"set seconds of d to {dt.second}\\n"
        "d"
    )
    return f'(run script "{inner}")'


def parse_date(text: str) -> str:
    """Parse an AppleScript date string to ISO 8601 format.

    AppleScript dates look like: "Sunday, March 22, 2026 at 2:00:00 PM"
    or various locale-specific formats. We attempt several common patterns.
    """
    text = text.strip()
    # Remove day name prefix if present (e.g., "Sunday, ")
    text = re.sub(r"^\w+day,\s*", "", text)
    # Remove " at " between date and time
    text = text.replace(" at ", " ")
    # Try common formats
    for fmt in (
        "%B %d, %Y %I:%M:%S %p",   # March 22, 2026 2:00:00 PM
        "%d. %B %Y %H:%M:%S",       # 22. mars 2026 14:00:00 (Norwegian)
        "%Y-%m-%d %H:%M:%S",        # 2026-03-22 14:00:00
        "%d/%m/%Y %H:%M:%S",        # 22/03/2026 14:00:00
        "%m/%d/%Y %H:%M:%S",        # 03/22/2026 14:00:00
    ):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.isoformat()
        except ValueError:
            continue
    # Fallback: return as-is
    return text


def date_component_lines(var: str, dt: datetime) -> str:
    """Build AppleScript that constructs a date in `var` from components.

    Unlike `date "..."` literals, component assignment does not depend on
    the system locale. Day is reset to 1 first so month assignment cannot
    overflow (e.g. current date Aug 31 -> month 2 would roll into March).
    """
    secs = dt.hour * 3600 + dt.minute * 60 + dt.second
    return (
        f"set {var} to current date\n"
        f"set day of {var} to 1\n"
        f"set year of {var} to {dt.year}\n"
        f"set month of {var} to {dt.month}\n"
        f"set day of {var} to {dt.day}\n"
        f"set time of {var} to {secs}"
    )


# Locale-independent AppleScript folder keywords
FOLDER_MAP = {
    "inbox": "inbox",
    "sent": "sent items",
    "sentmail": "sent items",
    "sent items": "sent items",
    "drafts": "drafts",
    "deleted": "deleted items",
    "deleted items": "deleted items",
    "trash": "deleted items",
    "junk": "junk mail",
    "spam": "junk mail",
    "outbox": "outbox",
}


def resolve_folder_ref(folder_name: str) -> str:
    """Map a user-facing folder name to an AppleScript folder reference.

    Returns an AppleScript expression like 'inbox' or 'mail folder "Archive"'.
    Built-in folders use locale-independent keywords; custom folders use name lookup.
    """
    key = folder_name.lower().strip()
    if key in FOLDER_MAP:
        return FOLDER_MAP[key]
    # Custom folder — search by name
    return f'mail folder "{escape(folder_name.strip())}"'


def is_builtin_folder(folder_name: str) -> bool:
    """True if resolve_folder_ref maps this name to a built-in folder keyword."""
    return folder_name.lower().strip() in FOLDER_MAP


# Separators for structured AppleScript output.
#
# Outlook item fields (subjects, sender names, attachment names, bodies) are
# chosen by whoever sent the mail. Printable separators such as "|||" could be
# planted in a subject to forge record and field boundaries (fake entry_id,
# fake sender). So the separators are ASCII control characters, and every
# field value is passed through the AppleScript handler `sf` (SANITIZE_HANDLER,
# prepended to every script by AppleScriptBridge.run), which removes them.
DELIM = "\x01"           # SOH: between fields
RECORD_DELIM = "\x02"    # STX: between records
# (Not ASCII 28-31: Python treats those as whitespace, so str.strip() and
# str.splitlines() would eat separators and shift fields.)
# The same separators as AppleScript expressions, for building scripts.
AS_DELIM = "(character id 1)"
AS_RECORD_DELIM = "(character id 2)"

SANITIZE_HANDLER = """on sf(v)
\tif v is missing value then return ""
\ttry
\t\tset t to v as text
\ton error
\t\treturn ""
\tend try
\tset od to AppleScript's text item delimiters
\tset AppleScript's text item delimiters to {character id 1, character id 2}
\tset parts to text items of t
\tset AppleScript's text item delimiters to " "
\tset t to parts as text
\tset AppleScript's text item delimiters to od
\treturn t
end sf
"""
