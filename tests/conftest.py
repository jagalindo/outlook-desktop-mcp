"""Shared fixtures for the unit tests. None of them needs Outlook: the macOS
server's AppleScript bridge is replaced by a fake that records each script
and returns a canned response."""
import asyncio
import json

import pytest

from outlook_desktop_mcp import server_mac
from outlook_desktop_mcp.utils.applescript_helpers import DELIM, RECORD_DELIM


class FakeBridge:
    """Stand-in for AppleScriptBridge. Records every script it is asked to
    run and returns `response` (a string, or a callable taking the script)."""

    def __init__(self):
        self.scripts: list[str] = []
        self.response = ""

    @property
    def last_script(self) -> str:
        return self.scripts[-1]

    @property
    def calls(self) -> int:
        return len(self.scripts)

    async def run(self, script, timeout=None):
        self.scripts.append(script)
        if isinstance(self.response, Exception):
            raise self.response
        if callable(self.response):
            return self.response(script)
        return self.response


@pytest.fixture
def fake(monkeypatch):
    bridge = FakeBridge()
    monkeypatch.setattr(server_mac, "bridge", bridge)
    return bridge


def call(coro):
    """Run a tool coroutine and decode its JSON result."""
    return json.loads(asyncio.run(coro))


def rec(*fields):
    """One list record, as the AppleScript side builds it."""
    return DELIM.join(fields) + RECORD_DELIM


def fields(*values):
    """One single-item response (no record terminator)."""
    return DELIM.join(values)
