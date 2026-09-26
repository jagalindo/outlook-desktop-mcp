from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("outlook-desktop-mcp")
except PackageNotFoundError:  # running from a source tree that is not installed
    __version__ = "0+unknown"
