"""Tests for the osascript bridge, with the subprocess faked out so they run
on any OS."""
import asyncio

import pytest

from outlook_desktop_mcp import applescript_bridge as b
from outlook_desktop_mcp.utils.applescript_helpers import DELIM, RECORD_DELIM


class FakeProc:
    def __init__(self, stdout=b"", stderr=b"", returncode=0, hang=False):
        self._out, self._err, self.returncode, self._hang = stdout, stderr, returncode, hang
        self.killed = False

    async def communicate(self):
        if self._hang and not self.killed:
            await asyncio.sleep(10)
        return self._out, self._err

    def kill(self):
        self.killed = True


def patch_proc(monkeypatch, proc):
    async def fake_exec(*args, **kwargs):
        return proc
    monkeypatch.setattr(b.asyncio, "create_subprocess_exec", fake_exec)


def test_run_keeps_trailing_delimiters(monkeypatch):
    # Only osascript's newline is removed; the empty last field survives.
    out = f"1{DELIM}a{DELIM}{RECORD_DELIM}\n".encode()
    patch_proc(monkeypatch, FakeProc(stdout=out))
    assert asyncio.run(b.AppleScriptBridge().run("x")) == f"1{DELIM}a{DELIM}{RECORD_DELIM}"


@pytest.mark.parametrize("stderr, code, number", [
    ("execution error: Can't get message id 5. (-1728)\n", "not_found", -1728),
    ("execution error: Not authorized to send Apple events to Microsoft Outlook. (-1743)", "permission_denied", -1743),
    ("execution error: Application isn't running. (-600)", "outlook_not_running", -600),
    ("syntax error: Expected end of line. (-2741)", "applescript_error", -2741),
    ("something without a number", "applescript_error", None),
])
def test_errors_are_classified(monkeypatch, stderr, code, number):
    patch_proc(monkeypatch, FakeProc(stderr=stderr.encode(), returncode=1))
    with pytest.raises(b.AppleScriptError) as info:
        asyncio.run(b.AppleScriptBridge().run("x"))
    assert info.value.code == code and info.value.number == number


def test_timeout_kills_the_process(monkeypatch):
    proc = FakeProc(hang=True)
    patch_proc(monkeypatch, proc)
    with pytest.raises(b.AppleScriptTimeout):
        asyncio.run(b.AppleScriptBridge().run("x", timeout=0.05))
    assert proc.killed


def test_scripts_are_serialized(monkeypatch):
    running = 0
    peak = 0

    class SlowProc(FakeProc):
        async def communicate(self):
            nonlocal running, peak
            running += 1
            peak = max(peak, running)
            await asyncio.sleep(0.01)
            running -= 1
            return b"ok\n", b""

    async def fake_exec(*args, **kwargs):
        return SlowProc()
    monkeypatch.setattr(b.asyncio, "create_subprocess_exec", fake_exec)

    async def many():
        bridge = b.AppleScriptBridge()
        return await asyncio.gather(*(bridge.run("x") for _ in range(5)))

    assert asyncio.run(many()) == ["ok"] * 5
    assert peak == b.MAX_CONCURRENT_SCRIPTS


def test_semaphore_survives_a_new_event_loop(monkeypatch):
    # main() runs the startup check and the server in separate loops.
    patch_proc(monkeypatch, FakeProc(stdout=b"ok\n"))
    bridge = b.AppleScriptBridge()
    assert asyncio.run(bridge.run("x")) == "ok"
    assert asyncio.run(bridge.run("x")) == "ok"
