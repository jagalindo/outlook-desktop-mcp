"""Tests for the pure AppleScript helpers."""
import pytest

from outlook_desktop_mcp.utils.applescript_helpers import escape, resolve_folder_ref, text_to_html


@pytest.mark.parametrize("raw, escaped", [
    ('say "hi"', 'say \\"hi\\"'),
    ("back\\slash", "back\\\\slash"),
    ("a\nb\tc\rd", "a\\nb\\tc\\rd"),
    ('" & (do shell script "x") & "', '\\" & (do shell script \\"x\\") & \\"'),
])
def test_escape(raw, escaped):
    assert escape(raw) == escaped


def test_text_to_html_keeps_line_breaks_and_escapes_markup():
    assert text_to_html("a < b\r\nc & d") == "<html><body>a &lt; b<br>c &amp; d</body></html>"
    assert text_to_html("x\ny", wrap=False) == "x<br>y"
    assert text_to_html("") == ""


@pytest.mark.parametrize("name, ref", [
    ("Inbox", "inbox"),
    ("sentmail", "sent items"),
    (" Trash ", "deleted items"),
    ("spam", "junk mail"),
    ('My "Projects"', 'mail folder "My \\"Projects\\""'),
])
def test_resolve_folder_ref(name, ref):
    assert resolve_folder_ref(name) == ref
