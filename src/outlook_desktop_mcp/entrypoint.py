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
        # Importing the Windows server here used to fail with a confusing
        # "No module named 'pythoncom'".
        sys.exit(
            f"outlook-desktop-mcp supports Windows and macOS only "
            f"(this platform is {sys.platform!r}): it drives the locally "
            f"installed Outlook desktop app."
        )
    _main()
