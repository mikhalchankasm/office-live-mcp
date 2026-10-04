"""Точка входа собранного office-live-mcp.exe (PyInstaller): то же, что python -m office_live."""

import sys

from office_live.__main__ import main

if __name__ == "__main__":
    sys.exit(main())
