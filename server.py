"""Совместимая точка входа: python server.py == python -m office_live (stdio MCP)."""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from office_live.app import main  # noqa: E402

if __name__ == "__main__":
    main()
