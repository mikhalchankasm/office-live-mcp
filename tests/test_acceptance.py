"""Критерии 1–10: временный профиль, готовый архив при наличии, иначе исходники; без Office."""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import pytest

from office_live import cli, install, registrations as reg
from office_live.probe import smoke
from office_live.state import read_state

REPO = Path(__file__).resolve().parents[1]


@pytest.fixture
def acceptance(tmp_path, monkeypatch):
    # Короткий путь нужен и для распаковки вложенных библиотек PyInstaller.
    with tempfile.TemporaryDirectory(prefix="olacc-") as temporary:
        root = Path(temporary)
        for key, part in (("USERPROFILE", "home"), ("HOME", "home"), ("APPDATA", "roaming"),
                          ("LOCALAPPDATA", "local"), ("CODEX_HOME", "Особый Codex")):
            monkeypatch.setenv(key, str(root / part))
        for key in list(os.environ):
            if key.startswith("OFFICE_LIVE_") and key != "OFFICE_LIVE_TEST_EXE":
                monkeypatch.delenv(key)
        monkeypatch.setenv("OFFICE_LIVE_LIVE_TESTS", "0")
        monkeypatch.chdir(root)
        external = os.environ.get("OFFICE_LIVE_TEST_EXE")
        archives = list((REPO / "dist").glob("office-live-mcp-*-win64.zip"))
        if external:
            source_exe = Path(external)
        elif archives:
            with zipfile.ZipFile(max(archives, key=lambda p: p.stat().st_mtime)) as archive:
                archive.extractall(root / "archive")
            source_exe = next((root / "archive").glob("*/app/office-live-mcp.exe"))
        else:
            source_exe = None
        try:
            yield root, source_exe
        finally:
            monkeypatch.chdir(REPO)


def test_owner_acceptance_1_to_10(acceptance, monkeypatch, capsys):
    root, source_exe = acceptance
    target = root / "Программа с пробелами"
    # Проверяем пути до первого вызова установки.
    assert Path.home().is_relative_to(root)
    for client in cli.CLIENT_LABELS:
        if client != "vscode":
            assert Path(reg.target(client)["path"]).is_relative_to(root)

    def execute(command, *args, expected=0):
        # Environment isolation does NOT redirect HKCU. Every child install must skip the protocol.
        if args[0] == "install":
            args = (*args, "--no-links")
        if source_exe and args[0] == "install":
            # Проверка наличия Office всегда подменяется: даже HKCR на рабочей машине не читаем.
            # Копирование и MCP-проверки выполняются с настоящим EXE в изолированном профиле.
            with monkeypatch.context() as patch:
                patch.setattr(install, "frozen", lambda: True)
                patch.setattr(sys, "executable", str(source_exe))
                patch.setattr(install, "_office_installed", lambda progid: True)
                assert install.install_cmd(list(args[1:])) == expected
                return capsys.readouterr().out
        result = subprocess.run([*command, *args], env=os.environ.copy(), cwd=REPO, capture_output=True, timeout=120)
        output = result.stdout.decode("utf-8") + result.stderr.decode("utf-8")
        assert result.returncode == expected, output
        return output

    if source_exe:
        archive_command = [str(source_exe)]
        execute(archive_command, "install", "--target", str(target), "--clients", "none", "--yes")
        command = [str(target / "app" / install.EXE_NAME)]
        update = lambda: execute(archive_command, "install", "--yes")
        uninstall = lambda: execute(archive_command, "uninstall", "--target", str(target))
        backend = "готовый EXE"
    else:
        # Разработческий сценарий: одна копия исходников, запись указывает на её server.py.
        shutil.copytree(REPO / "office_live", target / "office_live", ignore=shutil.ignore_patterns("__pycache__"))
        shutil.copy2(REPO / "server.py", target / "server.py")
        command = [sys.executable, str(target / "server.py")]
        update = lambda: execute(command, "setup", "--clients", "codex,zcode", "--yes")
        uninstall = lambda: execute(command, "setup", "--remove", "--clients", "codex,zcode,cursor", "--yes")
        backend = "исходники (копирование/откат EXE отдельно проверены тестами установщика)"

    execute(command, "setup", "--clients", "codex,zcode", "--yes")
    execute(command, "setup", "--clients", "codex", "--full", "--yes")
    execute(command, "setup", "--clients", "zcode", "--readonly", "--yes")
    assert smoke(command, {"OFFICE_LIVE_MODE": "full"}) > smoke(command, {"OFFICE_LIVE_MODE": "readonly"}) > 0
    codex_path = Path(reg.target("codex")["path"])
    zcode_path = Path(reg.target("zcode")["path"])
    before = (codex_path.read_bytes(), zcode_path.read_bytes())
    update()  # 5. повторная установка
    update()  # 6. обновление с сохранением ограничений
    assert (codex_path.read_bytes(), zcode_path.read_bytes()) == before
    assert len(read_state()["registrations"]) == 2
    execute(command, "setup", "--remove", "--clients", "zcode", "--yes")
    assert reg.read_entry(reg.target("codex"))[2] and reg.read_entry(reg.target("zcode"))[2] is None
    custom = root / "нестандартный клиент.json"
    custom.write_text('{"mcpServers":{"foreign":{"command":"keep"}},"setting":42}', encoding="utf-8")
    execute(command, "setup", "--clients", "cursor", "--config", str(custom), "--yes")
    uninstall()
    assert not read_state()["registrations"]
    assert all(r.get("kind") != "protocol" for r in read_state()["registrations"])
    assert json.loads(custom.read_text(encoding="utf-8")) == {"mcpServers": {"foreign": {"command": "keep"}}, "setting": 42}
    if source_exe:
        assert not target.exists()
    # 10. Повреждённый файл и чужая запись не затираются; повтор после исправления восстанавливает учёт.
    retry_command = archive_command if source_exe else command
    custom.write_text("{broken", encoding="utf-8")
    execute(retry_command, "setup", "--clients", "cursor", "--config", str(custom), "--yes", expected=1)
    assert custom.read_text() == "{broken"
    custom.write_text('{"mcpServers":{"office-live":{"command":"foreign"}}}', encoding="utf-8")
    execute(retry_command, "setup", "--clients", "cursor", "--config", str(custom), "--yes", expected=1)
    assert "foreign" in custom.read_text()
    execute(retry_command, "setup", "--clients", "cursor", "--config", str(custom), "--force", "--yes")
    execute(retry_command, "setup", "--remove", "--clients", "cursor", "--config", str(custom), "--yes")
    assert not read_state()["registrations"]
    print(f"Критерии 1–10: {backend}; Office и клиентские приложения не запускались")
