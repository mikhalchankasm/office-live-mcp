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
    for _key in list(os.environ):
        if _key.startswith("OFFICE_LIVE_") and _key not in {"OFFICE_LIVE_LIVE_TESTS", "OFFICE_LIVE_TEST_EXE"}:
            os.environ.pop(_key)
    _profile = tempfile.mkdtemp(prefix="office-live-profile-")
    for _key, _part in (("USERPROFILE", "home"), ("HOME", "home"), ("APPDATA", "roaming"),
                        ("LOCALAPPDATA", "local"), ("CODEX_HOME", "home/.codex")):
        os.environ[_key] = os.path.join(_profile, _part)


@pytest.fixture(autouse=True)
def isolated_profile(tmp_path, monkeypatch, request):
    if "live" in request.keywords:
        return
    for key, part in (("USERPROFILE", "home"), ("HOME", "home"), ("APPDATA", "roaming"),
                      ("LOCALAPPDATA", "local"), ("CODEX_HOME", "home/.codex")):
        monkeypatch.setenv(key, str(tmp_path / part))


@pytest.fixture(autouse=True)
def fake_registry(monkeypatch, request):
    if "live" in request.keywords:
        return
    import winreg

    from office_live import protocol
    from tests.registry_fake import FakeRegistry

    def forbidden(*args, **kwargs):
        raise AssertionError("Unit tests must never write to the native registry")

    for name in ("CreateKey", "CreateKeyEx", "SetValue", "SetValueEx", "DeleteKey", "DeleteKeyEx", "DeleteValue", "LoadKey", "SaveKey"):
        monkeypatch.setattr(winreg, name, forbidden)
    registry = FakeRegistry()
    monkeypatch.setattr(protocol, "Registry", lambda: registry)
    return registry


@pytest.fixture(autouse=True)
def no_native_window_mutations(monkeypatch, request):
    if "live" in request.keywords:
        return
    import ctypes

    from office_live import win32

    def forbidden(*args, **kwargs):
        raise AssertionError("Unit tests must never call native window APIs; use the fake Win32 layer")

    monkeypatch.setattr(win32, "_load_dll", forbidden)
    names = ("SetWindowPos", "ShowWindow", "ShowWindowAsync", "SetForegroundWindow", "SetWindowPlacement",
             "AllowSetForegroundWindow", "AttachThreadInput", "FlashWindowEx", "BringWindowToTop", "SetFocus", "SetActiveWindow")
    for name in names:
        monkeypatch.setattr(ctypes.windll.user32, name, forbidden)
    original = ctypes.WinDLL

    def guarded_dll(name, *args, **kwargs):
        dll = original(name, *args, **kwargs)
        if os.path.basename(str(name)).lower() in {"user32", "user32.dll"}:
            for entry in names:
                monkeypatch.setattr(dll, entry, forbidden)
        return dll

    monkeypatch.setattr(ctypes, "WinDLL", guarded_dll)


@pytest.fixture(autouse=True)
def no_real_office(monkeypatch, request):
    if "live" in request.keywords:
        return
    import win32com.client

    def forbidden(*args, **kwargs):
        raise AssertionError("Unit tests must use fake Office objects")

    for name in ("Dispatch", "DispatchEx", "GetActiveObject"):
        monkeypatch.setattr(win32com.client, name, forbidden)


def pytest_configure(config):
    config.addinivalue_line("markers", "live: needs real Excel/Word (set OFFICE_LIVE_LIVE_TESTS=1)")


def pytest_collection_modifyitems(config, items):
    if os.environ.get("OFFICE_LIVE_LIVE_TESTS") == "1":
        return
    skip = pytest.mark.skip(reason="live Office tests: set OFFICE_LIVE_LIVE_TESTS=1")
    for item in items:
        if "live" in item.keywords:
            item.add_marker(skip)
