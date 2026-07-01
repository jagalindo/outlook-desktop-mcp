"""Platform-aware entry point for outlook-desktop-mcp.

Detects the OS and imports the correct server module:
- macOS → server_mac (AppleScript automation)
- Windows → server (COM automation)
"""
import sys


def main():
    if sys.platform == "darwin":
        from outlook_desktop_mcp.server_mac import main as _main
    elif sys.platform == "win32":
        from outlook_desktop_mcp.server import main as _main
    else:
        sys.exit(
            f"outlook-desktop-mcp: unsupported platform {sys.platform!r}. "
            "This server drives a locally running Outlook Desktop and only "
            "works on Windows (COM) or macOS (AppleScript)."
        )
    _main()
