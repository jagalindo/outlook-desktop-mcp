"""
AppleScript Bridge
==================
Runs Outlook automation via osascript subprocess calls on macOS.
Each call is stateless — no persistent COM-like objects.

Every tool builds an AppleScript string and passes it to bridge.run().
"""
import asyncio
import logging
import re

logger = logging.getLogger("outlook_desktop_mcp.applescript_bridge")

STARTUP_TIMEOUT = 10
SCRIPT_TIMEOUT = 30

# Outlook handles Apple events one at a time, so parallel tool calls only queue
# up inside Outlook while their osascript timeouts keep running. Serializing
# them here keeps each timeout honest.
MAX_CONCURRENT_SCRIPTS = 1

# AppleScript error numbers worth telling apart. osascript prints them at the
# end of stderr, e.g. "... Can't get message id 5. (-1728)".
ERROR_CODES = {
    -1728: "not_found",
    -1719: "not_found",       # invalid index
    -1743: "permission_denied",  # automation not allowed in System Settings
    -600: "outlook_not_running",
    -609: "outlook_not_running",
}

_ERROR_NUMBER = re.compile(r"\((-?\d+)\)\s*$")


class AppleScriptError(RuntimeError):
    """osascript exited with an error. `code` is a short name from ERROR_CODES,
    or "applescript_error" when the number is not one we recognize."""

    def __init__(self, message: str, number: int | None = None):
        super().__init__(message)
        self.number = number
        self.code = ERROR_CODES.get(number, "applescript_error")


class AppleScriptTimeout(RuntimeError):
    """osascript did not finish in time. Killing it does not recall an Apple
    event Outlook has already received, so the script may still complete."""

    code = "timeout"


def parse_error(stderr: str) -> AppleScriptError:
    """Build an AppleScriptError from osascript's stderr."""
    err = stderr.strip()
    match = _ERROR_NUMBER.search(err)
    number = int(match.group(1)) if match else None
    return AppleScriptError(f"AppleScript error: {err}", number)


class AppleScriptBridge:
    """Manages AppleScript execution for Outlook on macOS."""

    def __init__(self):
        self._version: str | None = None
        self._semaphore: asyncio.Semaphore | None = None
        self._semaphore_loop: asyncio.AbstractEventLoop | None = None

    def _limit(self) -> asyncio.Semaphore:
        # main() runs the startup check and the MCP server in two separate
        # event loops, so the semaphore is created per loop.
        loop = asyncio.get_running_loop()
        if self._semaphore is None or self._semaphore_loop is not loop:
            self._semaphore = asyncio.Semaphore(MAX_CONCURRENT_SCRIPTS)
            self._semaphore_loop = loop
        return self._semaphore

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

        Only the trailing newline osascript appends is removed. A full strip()
        would also eat the ASCII field and record delimiters (Python counts
        them as whitespace), dropping empty trailing fields from the output.

        Raises AppleScriptError on non-zero exit and AppleScriptTimeout on
        timeout.
        """
        async with self._limit():
            proc = await asyncio.create_subprocess_exec(
                "osascript", "-e", script,
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
                raise AppleScriptTimeout(
                    f"AppleScript timed out after {timeout}s"
                )

        if proc.returncode != 0:
            raise parse_error(stderr.decode("utf-8", errors="replace"))

        return stdout.decode("utf-8", errors="replace").rstrip("\r\n")

    def stop(self):
        """No-op — AppleScript has no persistent resources to release."""
        pass
