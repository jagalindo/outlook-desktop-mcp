import pytest

from outlook_desktop_mcp import __version__, entrypoint


def test_unsupported_platform_exits_with_a_clear_message(monkeypatch):
    monkeypatch.setattr(entrypoint.sys, "platform", "linux")
    with pytest.raises(SystemExit, match="Windows and macOS only"):
        entrypoint.main()


def test_version_is_exposed():
    assert __version__ and __version__ != "0.4.0"
