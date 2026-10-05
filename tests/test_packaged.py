"""Проверки готового EXE из dist, только MCP и собственные временные файлы."""

import json
import os
import subprocess
import tempfile
import time
import zipfile
from pathlib import Path

import pytest

from office_live import __version__
from office_live.probe import StdioClient, smoke


@pytest.fixture
def packaged(monkeypatch):
    # Never combine the current source installer with a stale binary left in ignored dist/.
    archives = list((Path(__file__).resolve().parents[1] / "dist").glob(f"office-live-mcp-{__version__}-win64.zip"))
    if not archives:
        pytest.skip(f"Архив версии {__version__} ещё не собран; MCP исходников проверяется test_stdio")
    with tempfile.TemporaryDirectory(prefix="olexe-") as temporary:
        root = Path(temporary)
        for key, part in (("USERPROFILE", "home"), ("HOME", "home"), ("APPDATA", "roaming"),
                          ("LOCALAPPDATA", "local"), ("CODEX_HOME", "codex")):
            monkeypatch.setenv(key, str(root / part))
        with zipfile.ZipFile(max(archives, key=lambda p: p.stat().st_mtime)) as archive:
            archive.extractall(root / "archive")
        exe = next((root / "archive").glob("*/app/office-live-mcp.exe"))
        yield root, exe


def test_exe_mcp_and_utf8_diagnostics_in_ansi_environment(packaged):
    _, exe = packaged
    assert smoke([str(exe)], {"OFFICE_LIVE_MODE": "full", "PYTHONIOENCODING": "cp1251"}) > 0
    assert smoke([str(exe)], {"OFFICE_LIVE_MODE": "readonly", "PYTHONIOENCODING": "cp1251"}) > 0
    result = subprocess.run([str(exe), "serve"], input=b"", capture_output=True,
                            env={**os.environ, "OFFICE_LIVE_TOOLSETS": "unknown", "PYTHONIOENCODING": "cp1251"}, timeout=30)
    assert result.returncode == 2 and result.stdout == b""
    assert "Ошибка настройки" in result.stderr.decode("utf-8", "strict")


def test_installed_exe_in_use_and_self_uninstall_report(packaged, monkeypatch):
    from office_live import cli, install

    root, archive_exe = packaged
    target = root / "installed"
    # В CI Office отсутствует; эта проверка подменяет только сведения реестра, а не запуск EXE.
    with monkeypatch.context() as patch:
        import sys

        patch.setattr(sys, "executable", str(archive_exe))
        patch.setattr(install, "frozen", lambda: True)
        patch.setattr(install, "_office_installed", lambda progid: True)
        assert install.install_cmd(["--target", str(target), "--yes", "--clients", "none", "--no-links"]) == 0
        exe = target / "app" / install.EXE_NAME
        client = StdioClient([str(exe), "serve"])
        try:
            assert install.install_cmd(["--yes", "--no-links"]) == 1
            assert client.list_tools()
        finally:
            client.close()
    assert cli._server_command()[1]  # frozen-подмена уже снята
    result = subprocess.run([str(exe), "uninstall"], capture_output=True, timeout=30)
    assert result.returncode == 0, result.stdout.decode("utf-8")
    report = Path(os.environ["LOCALAPPDATA"]) / "office-live-mcp/uninstall-report.json"
    deadline = time.monotonic() + 30
    while time.monotonic() < deadline:
        if report.exists():
            try:
                data = json.loads(report.read_text(encoding="utf-8-sig"))
                if data.get("status") == "complete":
                    break
            except ValueError:
                pass
        time.sleep(0.1)
    else:
        pytest.fail("Фоновое удаление не завершилось за 30 секунд")
    assert data["errors"] == [] and data["remaining"] == []
    assert not target.exists()


def test_packaged_link_gui_rejects_invalid_links_quickly(packaged):
    import struct

    _, exe = packaged
    gui = exe.with_name("office-live-link.exe")
    data = gui.read_bytes()
    pe = struct.unpack_from("<I", data, 0x3C)[0]
    assert struct.unpack_from("<H", data, pe + 24 + 68)[0] == 2  # IMAGE_SUBSYSTEM_WINDOWS_GUI
    for args in ([], ["open-link", "javascript:bad"], ["setup", "--links", "--dry-run"]):
        result = subprocess.run([str(gui), *args], stdin=subprocess.DEVNULL, capture_output=True, timeout=15,
                                env={**os.environ, "OFFICE_LIVE_LINK_NO_DIALOG": "1"})
        assert result.returncode != 0
        assert b"Traceback" not in result.stderr
    result = subprocess.run([str(exe), "tools", "--json"], capture_output=True, timeout=30)
    assert any(t["name"] == "office_link" and t["kind"] == "read" for t in json.loads(result.stdout))
