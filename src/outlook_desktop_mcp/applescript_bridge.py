"""
AppleScript Bridge
==================
Runs Outlook automation via osascript subprocess calls on macOS.
Each call is stateless — no persistent COM-like objects.

Every tool builds an AppleScript string and passes it to bridge.run().
"""
import asyncio
import logging
import os
import re

from outlook_desktop_mcp.utils.applescript_helpers import SANITIZE_HANDLER

logger = logging.getLogger("outlook_desktop_mcp.applescript_bridge")

# "Can't get message id 9.99999999E+8. (-1728)" -> "Outlook has no message with id 999999999."
_NOT_FOUND = re.compile(
    r"Can.t get (message|calendar event|task|contact|mail folder)(?: id)? (.+?)\. \(-1728\)"
)


def friendly_error(err: str) -> str | None:
    """Readable text for Outlook's "not found" errors (-1728), else None.

    AppleScript prints ids above 2^29 as reals in scientific notation and
    folder names with escaped quotes; both are turned back into what the
    caller passed in.
    """
    m = _NOT_FOUND.search(err)
    if not m:
        return None
    kind, target = m.group(1), m.group(2).strip()
    if len(target) >= 2 and target[0] == target[-1] == '"':
        target = target[1:-1].replace('\\"', '"')
    if re.fullmatch(r"[0-9]+(?:[.,][0-9]+)?E\+[0-9]+", target):
        target = str(int(float(target.replace(",", "."))))
    if kind == "mail folder":
        return (f"Outlook has no folder named '{target}'. "
                "Call list_folders to see the available folders.")
    return f"Outlook has no {kind} with id {target}."


def _env_timeout(name: str, default: float) -> float:
    """Read a timeout override (seconds) from the environment."""
    raw = os.environ.get(name)
    if raw:
        try:
            return float(raw)
        except ValueError:
            logger.warning("Invalid %s=%r; using default %ss", name, raw, default)
    return default


STARTUP_TIMEOUT = _env_timeout("OUTLOOK_MCP_STARTUP_TIMEOUT", 10)
SCRIPT_TIMEOUT = _env_timeout("OUTLOOK_MCP_SCRIPT_TIMEOUT", 120)


class AppleScriptBridge:
    """Manages AppleScript execution for Outlook on macOS."""

    def __init__(self):
        self._version: str | None = None

    async def start(self):
        """Verify Outlook is running and accessible. Call once at server startup."""
        try:
            self._version = await self.run(
                'tell application "Microsoft Outlook" to get version',
                timeout=STARTUP_TIMEOUT,
            )
            logger.info("AppleScript bridge ready. Outlook version: %s", self._version)
        except Exception as e:
            raise RuntimeError(
                f"Cannot connect to Microsoft Outlook via AppleScript. "
                f"Is Outlook running? Error: {e}"
            ) from e

    async def run(self, script: str, timeout: float = SCRIPT_TIMEOUT) -> str:
        """Execute an AppleScript and return stdout as a string.

        Raises RuntimeError on non-zero exit or timeout.
        """
        # Every script gets the `sf` handler that strips the output separators
        # from field values (see utils/applescript_helpers.py).
        proc = await asyncio.create_subprocess_exec(
            "osascript", "-e", SANITIZE_HANDLER + script,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        try:
            stdout, stderr = await asyncio.wait_for(
                proc.communicate(), timeout=timeout
            )
        except asyncio.TimeoutError:
            proc.kill()
            await proc.communicate()
            raise RuntimeError(
                f"AppleScript timed out after {timeout}s. If the script had "
                f"already reached Outlook, the operation may still have "
                f"completed; check before retrying anything that sends or "
                f"changes data."
            )
        except BaseException:
            # The request was cancelled (MCP notifications/cancelled, client
            # disconnect, server shutdown). Stop osascript so a later Apple
            # Event in the script (send, delete, move) is never dispatched.
            # Do not await here: the task is being cancelled.
            if proc.returncode is None:
                try:
                    proc.kill()
                except ProcessLookupError:
                    pass
            raise

        if proc.returncode != 0:
            err = stderr.decode("utf-8", errors="replace").strip()
            raise RuntimeError(friendly_error(err) or f"AppleScript error: {err}")

        return stdout.decode("utf-8", errors="replace").strip()

    async def run_lines(self, script: str, timeout: float = SCRIPT_TIMEOUT) -> list[str]:
        """Execute an AppleScript and return output split into non-empty lines."""
        result = await self.run(script, timeout=timeout)
        return [line for line in result.split("\n") if line.strip()]

    def stop(self):
        """No-op — AppleScript has no persistent resources to release."""
        pass
