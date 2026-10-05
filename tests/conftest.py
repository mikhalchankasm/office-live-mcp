"""Общие настройки pytest.

Юнит-тесты (tests/test_*.py) не требуют Office. Живые тесты (tests/live/test_*.py) работают с реальными Excel и Word
и запускаются только с OFFICE_LIVE_LIVE_TESTS=1:  set OFFICE_LIVE_LIVE_TESTS=1 && pytest tests/live
Они создают собственные черновые книги/документы и закрывают ТОЛЬКО их; чужие документы и сами приложения не трогаются.
"""

import os
import sys
import tempfile

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Журнал по умолчанию пишется в %LOCALAPPDATA%\office-live-mcp\journal: юнит-тесты не должны мусорить в профиле пользователя.
# Задаётся до импорта office_live: настройки читаются при импорте.
if os.environ.get("OFFICE_LIVE_LIVE_TESTS") != "1":
    os.environ["LOCALAPPDATA"] = tempfile.mkdtemp(prefix="office-live-localappdata-")


def pytest_configure(config):
    config.addinivalue_line("markers", "live: needs real Excel/Word (set OFFICE_LIVE_LIVE_TESTS=1)")


def pytest_collection_modifyitems(config, items):
    if os.environ.get("OFFICE_LIVE_LIVE_TESTS") == "1":
        return
    skip = pytest.mark.skip(reason="live Office tests: set OFFICE_LIVE_LIVE_TESTS=1")
    for item in items:
        if "live" in item.keywords:
            item.add_marker(skip)
