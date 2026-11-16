"""Shared rules for writing email attachments to disk (Windows and macOS).

Attachment names and bytes come from whoever sent the email, and the save
directory is chosen by the model, which may have been steered by mailbox
content. These helpers keep every save:

- inside one root directory (default ~/Downloads, or OUTLOOK_MCP_ATTACHMENT_DIR),
- under a sanitised single-component file name,
- without overwriting an existing file,
- without touching network/device paths before rejecting them, and
- marked as downloaded from the internet (Windows Mark-of-the-Web /
  macOS quarantine), so SmartScreen, Office Protected View and Gatekeeper
  still apply when the user opens the file later.
"""
import os
import re
import subprocess
import sys
import time

ATTACHMENT_DIR_ENV = "OUTLOOK_MCP_ATTACHMENT_DIR"
QUARANTINE_AGENT = "outlook-desktop-mcp-plus"


def attachment_root() -> str:
    """Root directory attachments may be saved under (realpath).

    Default ~/Downloads; override with the OUTLOOK_MCP_ATTACHMENT_DIR env var.
    """
    root = os.environ.get(ATTACHMENT_DIR_ENV) or os.path.join(
        os.path.expanduser("~"), "Downloads"
    )
    return os.path.realpath(os.path.expanduser(root))


def is_within(path: str, root: str) -> bool:
    """True if realpath ``path`` equals or is inside realpath ``root``."""
    try:
        return os.path.commonpath([os.path.normcase(path), os.path.normcase(root)]) == os.path.normcase(root)
    except ValueError:  # different drives, or mixed absolute/relative
        return False


def _is_network_or_device_path(path: str) -> bool:
    """UNC (\\\\host\\share, //host/share), WebDAV and device (\\\\?\\, \\\\.\\) paths.

    Checked lexically, before any filesystem call: resolving such a path with
    realpath() makes Windows open it, which can start an SMB/WebDAV connection
    (and NTLM authentication) to an attacker-chosen host.
    """
    return path.replace("/", "\\").startswith("\\\\")


def resolve_save_dir(save_directory: str, root: str):
    """Resolve the caller's save_directory; returns (path, error).

    Empty means the root itself. A relative path is taken relative to the
    root. Network and device paths are rejected before any filesystem access.
    The realpath (symlinks resolved) must lie inside the root.
    """
    if not save_directory:
        return root, None
    if _is_network_or_device_path(save_directory):
        return None, (
            "Error: save_directory must be a local folder inside the attachment "
            f"root '{root}'; network and device paths are not allowed."
        )
    candidate = os.path.expanduser(save_directory)
    if not os.path.isabs(candidate):
        candidate = os.path.join(root, candidate)
    resolved = os.path.realpath(candidate)
    if not is_within(resolved, root):
        return None, (
            f"Error: save_directory must be inside the attachment root '{root}'. "
            f"Set the {ATTACHMENT_DIR_ENV} environment variable to change the root."
        )
    return resolved, None


def safe_filename(name: str) -> str:
    """Reduce a sender-chosen attachment name to one safe path component."""
    name = (name or "").replace("\\", "/").rsplit("/", 1)[-1]
    name = re.sub(r"[^\w.\- ]", "_", name).strip(" .")
    return name or "attachment"


def unique_path(directory: str, filename: str) -> str:
    """Path in ``directory`` for ``filename`` that does not exist yet.

    Appends " (1)", " (2)", ... before the extension rather than overwriting.
    """
    path = os.path.join(directory, filename)
    if not os.path.lexists(path):
        return path
    stem, ext = os.path.splitext(filename)
    n = 1
    while True:
        path = os.path.join(directory, f"{stem} ({n}){ext}")
        if not os.path.lexists(path):
            return path
        n += 1


def mark_as_downloaded(path: str) -> bool:
    """Mark a saved attachment as coming from the internet. Returns success.

    Windows: writes the Zone.Identifier alternate data stream (ZoneId=3,
    Internet), which SmartScreen and Office Protected View check.
    macOS: sets com.apple.quarantine (unless Outlook already set it), which
    Gatekeeper checks on first open. Other platforms: no-op (False).
    """
    try:
        if os.name == "nt":
            with open(path + ":Zone.Identifier", "w") as f:
                f.write("[ZoneTransfer]\r\nZoneId=3\r\n")
            return True
        if sys.platform == "darwin":
            xattr = "/usr/bin/xattr"
            existing = subprocess.run([xattr, "-p", "com.apple.quarantine", path],
                                      capture_output=True, timeout=10)
            if existing.returncode == 0:
                return True  # already quarantined (e.g. by Outlook itself)
            value = f"0081;{int(time.time()):08x};{QUARANTINE_AGENT};"
            done = subprocess.run([xattr, "-w", "com.apple.quarantine", value, path],
                                  capture_output=True, timeout=10)
            return done.returncode == 0
    except (OSError, subprocess.SubprocessError):
        return False
    return False
