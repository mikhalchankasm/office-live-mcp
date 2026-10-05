"""Точка входа собранного office-live-mcp.exe (PyInstaller): то же, что python -m office_live."""

import sys

from office_live.__main__ import main

if __name__ == "__main__":
    if sys.executable.lower().endswith("office-live-link.exe"):
        # The GUI executable has exactly one purpose; no accidental server or installer on double click.
        from office_live.protocol import open_link

        sys.exit(open_link(sys.argv[2:] if sys.argv[1:2] == ["open-link"] else []))
    sys.exit(main())
