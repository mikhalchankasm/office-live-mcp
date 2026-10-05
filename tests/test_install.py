"""Установщик автономной сборки (office-live-mcp.exe install/uninstall). Собранный exe и Office не нужны: «замороженный»
режим и запуск дочерних exe подменяются."""

import sys
import os
import subprocess
from types import SimpleNamespace
import json
import base64
from pathlib import Path

import pytest

from office_live import __main__ as entry
from office_live import cli, install

REAL_STARTS = install._starts


@pytest.fixture(autouse=True)
def isolated_install_profile(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "profile"))


@pytest.fixture
def bundle(tmp_path, monkeypatch):
    """Распакованный архив: папка с exe, из которой запущен установщик; вызовы дочерних exe записываются."""
    src = tmp_path / "unzipped" / "app"
    (src / "_internal").mkdir(parents=True)
    (src / install.EXE_NAME).write_text("exe v2", encoding="utf-8")
    (src / "_internal" / "lib.pyd").write_text("lib v2", encoding="utf-8")
    monkeypatch.setattr(install, "frozen", lambda: True)
    monkeypatch.setattr(sys, "executable", str(src / install.EXE_NAME))
    monkeypatch.setattr(install, "_office_installed", lambda progid: True)
    calls = []
    monkeypatch.setattr(install.subprocess, "call", lambda cmd, **kw: calls.append([Path(cmd[0]).name, *cmd[1:]]) or 0)
    monkeypatch.setattr(install, "_starts", lambda exe: calls.append([exe.name, "tools", "--json"]) or 0)
    return SimpleBundle(src, tmp_path / "root", calls)


class SimpleBundle:
    def __init__(self, src, root, calls):
        self.src, self.root, self.calls = src, root, calls
        self.app = root / "app"


def test_install_copies_the_program_checks_it_and_connects_agents(bundle):
    rc = install.install_cmd(["--target", str(bundle.root), "--yes", "--clients", "cursor", "--readonly"])
    assert rc == 0
    assert (bundle.app / install.EXE_NAME).read_text(encoding="utf-8") == "exe v2"
    assert (bundle.app / "_internal" / "lib.pyd").exists()
    # проверяется и подключается УСТАНОВЛЕННАЯ копия (её путь попадёт в конфиги агентов), параметры уходят в setup
    assert bundle.calls == [
        [install.EXE_NAME, "doctor"], [install.EXE_NAME, "tools", "--json"], [install.EXE_NAME, "setup", "--yes", "--clients", "cursor", "--readonly"],
    ]


def test_update_replaces_the_old_copy_completely(bundle):
    (bundle.app / "_internal").mkdir(parents=True)
    (bundle.app / install.EXE_NAME).write_text("exe v1", encoding="utf-8")
    (bundle.app / "_internal" / "stale.pyd").write_text("old", encoding="utf-8")
    install._write_marker(bundle.root)
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 0
    assert (bundle.app / install.EXE_NAME).read_text(encoding="utf-8") == "exe v2"
    assert not (bundle.app / "_internal" / "stale.pyd").exists()
    assert sorted(p.name for p in bundle.root.iterdir()) == ["app", install.MARKER]


def test_a_copy_in_use_is_left_intact_and_the_user_is_told_to_close_agents(bundle, monkeypatch, capsys):
    bundle.app.mkdir(parents=True)
    (bundle.app / install.EXE_NAME).write_text("exe v1", encoding="utf-8")
    install._write_marker(bundle.root)
    real_rename = Path.rename

    def locked(self, target):
        if self.name == "app":
            raise PermissionError("in use")
        return real_rename(self, target)

    monkeypatch.setattr(Path, "rename", locked)
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 1
    assert (bundle.app / install.EXE_NAME).read_text(encoding="utf-8") == "exe v1"
    assert sorted(p.name for p in bundle.root.iterdir()) == ["app", install.MARKER]
    assert "закройте" in capsys.readouterr().out and bundle.calls == []


def test_without_excel_and_word_nothing_is_installed(bundle, monkeypatch, capsys):
    monkeypatch.setattr(install, "_office_installed", lambda progid: False)
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 1
    assert not bundle.root.exists() and bundle.calls == []
    assert "Excel или Word" in capsys.readouterr().out


def test_one_missing_office_app_is_only_a_warning(bundle, monkeypatch, capsys):
    monkeypatch.setattr(install, "_office_installed", lambda progid: progid.startswith("Excel"))
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 0
    assert "Word не установлен" in capsys.readouterr().out.replace("  ", " ")


def test_running_the_installed_exe_again_does_not_copy_onto_itself(bundle, monkeypatch):
    bundle.app.mkdir(parents=True)
    (bundle.app / install.EXE_NAME).write_text("installed", encoding="utf-8")
    install._write_marker(bundle.root)
    monkeypatch.setattr(sys, "executable", str(bundle.app / install.EXE_NAME))
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 0
    assert (bundle.app / install.EXE_NAME).read_text(encoding="utf-8") == "installed"


def test_install_from_source_is_refused(monkeypatch):
    monkeypatch.setattr(install, "frozen", lambda: False)
    assert install.install_cmd(["--yes"]) == 2


def test_default_location_is_per_user_and_needs_no_admin_rights(monkeypatch, tmp_path):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    assert install.install_root() == tmp_path / "Programs" / "office-live-mcp"


def test_uninstall_disconnects_every_agent_and_deletes_the_folder(tmp_path, monkeypatch):
    root = tmp_path / "root"
    (root / "app").mkdir(parents=True)
    install._write_marker(root)
    seen = []
    monkeypatch.setattr(install, "setup_cmd", lambda args: seen.append(args) or 0)
    monkeypatch.setattr(install, "frozen", lambda: False)
    assert install.uninstall_cmd(["--target", str(root)]) == 0
    assert seen[0][:2] == ["--remove", "--yes"] and "claude-code" in seen[0][3] and "vscode" not in seen[0][3]
    assert not root.exists()
    (root / "app").mkdir(parents=True)
    install.uninstall_cmd(["--target", str(root), "--keep-files"])
    assert root.exists()


def test_uninstall_without_marker_keeps_every_file(tmp_path, monkeypatch, capsys):
    (tmp_path / "app").mkdir()
    important = tmp_path / "app" / "important.txt"
    important.write_text("keep")
    monkeypatch.setattr(install, "setup_cmd", lambda args: pytest.fail("must validate ownership before disconnecting"))
    assert install.uninstall_cmd(["--target", str(tmp_path)]) != 0
    assert important.read_text() == "keep" and install.MARKER in capsys.readouterr().out


@pytest.mark.parametrize("problem", ["invalid-marker", "wrong-product", "root-reparse", "app-reparse", "app-file"])
def test_uninstall_invalid_target_does_not_disconnect_or_change_files(tmp_path, monkeypatch, problem):
    import stat

    root = tmp_path / "install"
    app = root / "app"
    root.mkdir()
    if problem == "app-file":
        app.write_text("keep app file")
    elif problem != "app-missing":
        app.mkdir()
        (app / install.EXE_NAME).write_text("keep exe")
    install._write_marker(root)
    marker = root / install.MARKER
    if problem == "invalid-marker":
        marker.write_text("{broken")
    elif problem == "wrong-product":
        marker.write_text('{"product": "other"}')
    if problem in {"root-reparse", "app-reparse"}:
        target = root if problem == "root-reparse" else app
        lstat = Path.lstat

        def reparse(path, *args, **kwargs):
            result = lstat(path, *args, **kwargs)
            return SimpleNamespace(st_file_attributes=result.st_file_attributes | stat.FILE_ATTRIBUTE_REPARSE_POINT) if path == target else result

        monkeypatch.setattr(Path, "lstat", reparse)
    before = {p.relative_to(root): p.read_bytes() for p in root.rglob("*") if p.is_file()}
    monkeypatch.setattr(install, "setup_cmd", lambda args: pytest.fail("must validate target before disconnecting"))
    monkeypatch.setattr(install.subprocess, "Popen", lambda *a, **kw: pytest.fail("must not detach"))
    assert install.uninstall_cmd(["--target", str(root)]) == 1
    assert {p.relative_to(root): p.read_bytes() for p in root.rglob("*") if p.is_file()} == before


def test_uninstall_keep_files_disconnects_without_folder_checks(tmp_path, monkeypatch):
    calls = []
    monkeypatch.setattr(install, "setup_cmd", lambda args: calls.append(args) or 0)
    monkeypatch.setattr(install, "_owned", lambda root: pytest.fail("must not check ownership"))
    monkeypatch.setattr(install, "_plain_dir", lambda path: pytest.fail("must not check folders"))
    assert install.uninstall_cmd(["--target", str(tmp_path), "--keep-files"]) == 0
    assert len(calls) == 1 and calls[0][:2] == ["--remove", "--yes"]


def test_uninstall_missing_root_disconnects_and_reports_no_files(tmp_path, monkeypatch, capsys):
    root = tmp_path / "missing"
    calls = []
    monkeypatch.setattr(install, "setup_cmd", lambda args: calls.append(args) or 0)
    assert install.uninstall_cmd(["--target", str(root)]) == 0
    assert len(calls) == 1 and calls[0][:2] == ["--remove", "--yes"]
    assert not root.exists() and "файлов не было" in capsys.readouterr().out.lower()


def test_install_does_not_claim_an_existing_unmarked_app(bundle):
    bundle.app.mkdir(parents=True)
    (bundle.app / "important.txt").write_text("keep")
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 1
    assert (bundle.app / "important.txt").read_text() == "keep"
    assert not (bundle.root / install.MARKER).exists()


@pytest.mark.parametrize("name,exe,removed", [
    ("app.old-20261005010203", True, True), ("app.new-123", True, True),
    ("app.notes-important", True, False), ("app.old-2026100501020", True, False),
    ("app.old-202610050102030", True, False), ("app.new-abc", True, False),
    ("app.old-20261005010203", False, False), ("app.new-123", False, False),
])
def test_cleanup_requires_exact_name_and_executable(tmp_path, name, exe, removed):
    install._write_marker(tmp_path)
    folder = tmp_path / name
    folder.mkdir()
    (folder / (install.EXE_NAME if exe else "important.txt")).write_text("data")
    install._cleanup_leftovers(tmp_path)
    assert folder.exists() is not removed


def test_uninstall_preserves_unrelated_root_files(tmp_path, monkeypatch):
    install._write_marker(tmp_path)
    (tmp_path / "app").mkdir()
    (tmp_path / "notes.txt").write_text("keep")
    (tmp_path / "app.notes-important").mkdir()
    monkeypatch.setattr(install, "setup_cmd", lambda args: 0)
    monkeypatch.setattr(install, "frozen", lambda: False)
    assert install.uninstall_cmd(["--target", str(tmp_path)]) == 0
    assert sorted(p.name for p in tmp_path.iterdir()) == ["app.notes-important", "notes.txt"]


def test_detached_uninstall_uses_bounded_literal_powershell_paths(tmp_path, monkeypatch):
    root = tmp_path / "user's files & %TEMP%"
    (root / "app").mkdir(parents=True)
    install._write_marker(root)
    monkeypatch.setattr(install, "setup_cmd", lambda args: 0)
    monkeypatch.setattr(install, "frozen", lambda: True)
    monkeypatch.setattr(sys, "executable", str(root / "app" / install.EXE_NAME))
    calls = []
    monkeypatch.setattr(install.subprocess, "Popen", lambda cmd, **kw: calls.append(cmd))
    assert install.uninstall_cmd(["--target", str(root)]) == 0
    script = base64.b64decode(calls[0][-1]).decode("utf-16-le")
    assert calls[0][0] == "powershell" and "user''s files & %TEMP%" in script
    assert "Remove-Item -LiteralPath $installRoot" not in script and "rd /s" not in script
    assert "[IO.Directory]::Delete($installRoot)" in script
    assert root.exists()  # detached process is never actually launched by this test


def test_detached_script_deletes_only_installer_items_in_temp_dir(tmp_path):
    import subprocess

    root = tmp_path / "install's & files"
    root.mkdir()
    install._write_marker(root)
    for name in ("app", "app.old-20261005010203", "app.new-123", "app.notes-important"):
        (root / name).mkdir()
        (root / name / install.EXE_NAME).write_text("fixture")
    (root / "app.new-456").mkdir()  # no executable, not ours
    (root / "notes.txt").write_text("keep")
    assert root.resolve().is_relative_to(tmp_path.resolve())
    encoded = base64.b64encode(install._self_delete_script(root.resolve()).encode("utf-16-le")).decode("ascii")
    result = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded], capture_output=True)
    assert result.returncode == 0, result.stderr
    assert sorted(p.name for p in root.iterdir()) == ["app.new-456", "app.notes-important", "notes.txt"]


def test_marker_exists_before_copy_and_is_kept_on_interruption(bundle, monkeypatch):
    def interrupted(source, app):
        assert install._owned(app.parent)
        raise OSError("copy interrupted")

    monkeypatch.setattr(install, "_replace_app", interrupted)
    with pytest.raises(OSError, match="interrupted"):
        install.install_cmd(["--target", str(bundle.root), "--yes"])
    assert install._owned(bundle.root)


def test_failed_second_rename_restores_old_app(bundle, monkeypatch):
    bundle.app.mkdir(parents=True)
    install._write_marker(bundle.root)
    (bundle.app / install.EXE_NAME).write_text("old")
    rename = Path.rename

    def fail_staging(path, target):
        if path.name.startswith("app.new-"):
            raise PermissionError("staging locked")
        return rename(path, target)

    monkeypatch.setattr(Path, "rename", fail_staging)
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 1
    assert (bundle.app / install.EXE_NAME).read_text() == "old" and not bundle.calls


@pytest.mark.parametrize("failure", ["return", "raise"])
def test_update_that_does_not_start_rolls_back_before_setup(bundle, monkeypatch, capsys, failure):
    bundle.app.mkdir(parents=True)
    install._write_marker(bundle.root)
    (bundle.app / install.EXE_NAME).write_text("old")

    def run(cmd, **kw):
        assert cmd[1:] == ["tools", "--json"]
        assert len(list(bundle.root.glob("app.old-*"))) == 1  # прежняя копия ещё цела
        if failure == "raise":
            raise OSError("invalid exe")
        return SimpleNamespace(returncode=1, stdout=b"", stderr=b"boom")

    monkeypatch.setattr(install, "_starts", REAL_STARTS)
    monkeypatch.setattr(install.subprocess, "run", run)
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 1
    assert (bundle.app / install.EXE_NAME).read_text() == "old"
    assert "восстановлена" in capsys.readouterr().out
    assert not any(c[1:2] == ["setup"] for c in bundle.calls)


def test_doctor_failure_alone_does_not_roll_back_an_update(bundle, monkeypatch):
    """doctor «падает» на машине без Word (или без Excel), где программа рабочая: обновление не должно откатываться."""
    bundle.app.mkdir(parents=True)
    install._write_marker(bundle.root)
    (bundle.app / install.EXE_NAME).write_text("old")
    monkeypatch.setattr(install.subprocess, "call", lambda cmd, **kw: 1 if cmd[1] == "doctor" else 0)
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 0
    assert (bundle.app / install.EXE_NAME).read_text(encoding="utf-8") == "exe v2"


def test_first_install_keeps_copy_but_does_not_connect_when_it_does_not_start(bundle, monkeypatch, capsys):
    monkeypatch.setattr(install, "_starts", lambda exe: 1)
    assert install.install_cmd(["--target", str(bundle.root), "--yes", "--clients", "cursor", "--readonly"]) == 1
    assert bundle.calls == [[install.EXE_NAME, "doctor"]]
    assert (bundle.app / install.EXE_NAME).read_text(encoding="utf-8") == "exe v2"
    assert (bundle.app / "_internal" / "lib.pyd").exists()
    marker = json.loads((bundle.root / install.MARKER).read_text())
    assert marker["product"] == install.PRODUCT and marker["version"] and marker["created"]
    output = capsys.readouterr().out
    assert "Файлы сохранены" in output and "Повторить" in output
    assert "install --target" in output and "--clients cursor --readonly" in output


def test_uninstall_disconnect_failure_keeps_program_and_marker(tmp_path, monkeypatch):
    install._write_marker(tmp_path)
    (tmp_path / "app").mkdir()
    (tmp_path / "app" / install.EXE_NAME).write_text("exe")
    monkeypatch.setattr(install, "setup_cmd", lambda args: 1)
    monkeypatch.setattr(install.subprocess, "Popen", lambda *a, **kw: pytest.fail("must not detach"))
    assert install.uninstall_cmd(["--target", str(tmp_path)]) == 1
    assert (tmp_path / "app" / install.EXE_NAME).read_text() == "exe" and (tmp_path / install.MARKER).exists()


def test_frozen_server_entry_points_at_the_exe_without_arguments(monkeypatch):
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", r"C:\Users\u\AppData\Local\Programs\office-live-mcp\app\office-live-mcp.exe")
    assert cli._server_command() == (sys.executable, [])
    assert cli._json_server_entry(None)["args"] == []


class Tty:
    def __init__(self, tty):
        self.tty = tty

    def isatty(self):
        return self.tty


@pytest.mark.parametrize("tty,expected", [(True, ("install", ["--pause"])), (False, None)])
def test_double_click_installs_but_an_agent_with_pipes_gets_the_server(monkeypatch, tty, expected):
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "stdin", Tty(tty))
    ran = []
    monkeypatch.setattr(cli, "run", lambda cmd, args: ran.append((cmd, args)) or 0)
    import office_live.app as app

    monkeypatch.setattr(app, "main", lambda: ran.append("serve"))
    entry.main([])
    assert ran == [expected or "serve"]


def test_uninstall_of_an_owned_folder_without_app_still_cleans_up(tmp_path, monkeypatch):
    """Прерванная установка или частичное удаление: метка есть, app нет — остатки удаляются, а не «отказ навсегда»."""
    root = tmp_path / "install"
    root.mkdir()
    install._write_marker(root)
    (root / "user.txt").write_text("keep")
    monkeypatch.setattr(install, "setup_cmd", lambda args: 0)
    monkeypatch.setattr(install, "frozen", lambda: False)
    assert install.uninstall_cmd(["--target", str(root)]) == 0
    assert not (root / install.MARKER).exists() and (root / "user.txt").read_text() == "keep"


def test_a_hanging_doctor_does_not_stop_the_installation(bundle, monkeypatch, capsys):
    """doctor обращается к запущенному Office; зависший Excel/Word не должен навсегда остановить установку."""
    import subprocess

    seen = []

    def call(cmd, timeout=None):
        seen.append((cmd[1], timeout))
        if cmd[1] == "doctor":
            raise subprocess.TimeoutExpired(cmd, timeout)
        return 0

    monkeypatch.setattr(install.subprocess, "call", call)
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 0
    assert seen[0] == ("doctor", install.DOCTOR_TIMEOUT) and seen[-1][0] == "setup"  # проверка запуска и подключение состоялись
    assert "doctor не ответил" in capsys.readouterr().out


@pytest.mark.parametrize("length,warning", [(199, False), (200, False), (201, True), (220, True)])
def test_packaged_cmd_long_path_warns_but_continues_in_isolated_profile(tmp_path, length, warning):
    source = (Path(__file__).parents[1] / "packaging" / "install.cmd").read_text(encoding="ascii")
    invocation = '"app\\office-live-mcp.exe" install %*'
    assert invocation in source
    # В копии теста заменяем только запуск exe на безопасную заглушку cmd; Office/установщик не запускаются.
    source = source.replace(invocation, 'call "app\\probe.cmd" install %*')
    base = tmp_path / "with spaces & apostrophe's !"
    count = length - len(str(base / "app" / "_internal")) - 1
    assert count > 0
    root = base / ("x" * count)
    app = root / "app"
    app.mkdir(parents=True)
    assert len(str(app / "_internal")) == length
    (app / "office-live-mcp.exe").write_text("never executed", encoding="ascii")
    (root / "install.cmd").write_text(source, encoding="ascii")
    (app / "probe.cmd").write_text('@echo off\necho PROBE_CALLED %*\necho PROFILE=%LOCALAPPDATA%\nexit /b 7\n', encoding="ascii")
    profile = tmp_path / "profile"
    env = {**os.environ, "LOCALAPPDATA": str(profile), "OFFICE_LIVE_NO_PAUSE": "1"}
    result = subprocess.run(["cmd.exe", "/d", "/v:on", "/c", r".\install.cmd", "--yes"], cwd=root, env=env, capture_output=True, text=True, timeout=20)
    assert result.returncode == 7, result.stdout + result.stderr
    assert ("WARNING:" in result.stdout) is warning
    assert "PROBE_CALLED install --yes" in result.stdout and str(profile) in result.stdout
    if warning:
        assert "shorter folder" in result.stdout and "Downloads" in result.stdout
