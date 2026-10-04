"""Установщик автономной сборки (office-live-mcp.exe install/uninstall). Собранный exe и Office не нужны: «замороженный»
режим и запуск дочерних exe подменяются."""

import sys
from pathlib import Path

import pytest

from office_live import __main__ as entry
from office_live import cli, install


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
    monkeypatch.setattr(install.subprocess, "call", lambda cmd: calls.append([Path(cmd[0]).name, *cmd[1:]]) or 0)
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
    assert bundle.calls == [[install.EXE_NAME, "doctor"], [install.EXE_NAME, "setup", "--yes", "--clients", "cursor", "--readonly"]]


def test_update_replaces_the_old_copy_completely(bundle):
    (bundle.app / "_internal").mkdir(parents=True)
    (bundle.app / install.EXE_NAME).write_text("exe v1", encoding="utf-8")
    (bundle.app / "_internal" / "stale.pyd").write_text("old", encoding="utf-8")
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 0
    assert (bundle.app / install.EXE_NAME).read_text(encoding="utf-8") == "exe v2"
    assert not (bundle.app / "_internal" / "stale.pyd").exists()
    assert sorted(p.name for p in bundle.root.iterdir()) == ["app"]  # ни app.old-*, ни app.new-*


def test_a_copy_in_use_is_left_intact_and_the_user_is_told_to_close_agents(bundle, monkeypatch, capsys):
    bundle.app.mkdir(parents=True)
    (bundle.app / install.EXE_NAME).write_text("exe v1", encoding="utf-8")
    real_rename = Path.rename

    def locked(self, target):
        if self.name == "app":
            raise PermissionError("in use")
        return real_rename(self, target)

    monkeypatch.setattr(Path, "rename", locked)
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 1
    assert (bundle.app / install.EXE_NAME).read_text(encoding="utf-8") == "exe v1"
    assert sorted(p.name for p in bundle.root.iterdir()) == ["app"]  # временная копия убрана
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
    seen = []
    monkeypatch.setattr(install, "setup_cmd", lambda args: seen.append(args) or 0)
    monkeypatch.setattr(install, "frozen", lambda: False)
    assert install.uninstall_cmd(["--target", str(root)]) == 0
    assert seen[0][:2] == ["--remove", "--yes"] and "claude-code" in seen[0][3] and "vscode" not in seen[0][3]
    assert not root.exists()
    (root / "app").mkdir(parents=True)
    install.uninstall_cmd(["--target", str(root), "--keep-files"])
    assert root.exists()


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
