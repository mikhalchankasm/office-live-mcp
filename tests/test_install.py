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
    (src / "office-live-link.exe").write_text("GUI exe", encoding="utf-8")
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


def test_profile_paths_are_isolated_before_installer_runs(tmp_path):
    assert Path.home().is_relative_to(tmp_path)
    assert cli._appdata().is_relative_to(tmp_path)
    assert install.install_root().is_relative_to(tmp_path)
    for client in ("cursor", "zcode", "claude-desktop"):
        assert cli._json_target(client)[0].is_relative_to(tmp_path)
    assert Path(os.environ["CODEX_HOME"]).is_relative_to(tmp_path)


def test_update_without_target_reuses_custom_install(bundle):
    assert install.install_cmd(["--target", str(bundle.root), "--yes", "--clients", "none"]) == 0
    (bundle.app / install.EXE_NAME).write_text("old", encoding="utf-8")
    assert install.install_cmd(["--yes"]) == 0
    assert (bundle.app / install.EXE_NAME).read_text(encoding="utf-8") == "exe v2"
    assert not install.install_root().exists()


def test_update_does_not_discover_or_reset_clients(bundle):
    assert install.install_cmd(["--target", str(bundle.root), "--yes", "--clients", "codex"]) == 0
    bundle.calls.clear()
    assert install.install_cmd(["--yes", "--readonly"]) == 0
    assert not any(c[1] == "setup" for c in bundle.calls)
    assert install.install_cmd(["--yes", "--clients", "zcode", "--readonly"]) == 0
    assert bundle.calls[-1][1:] == ["setup", "--yes", "--clients", "zcode", "--readonly"]


def test_new_location_requires_explicit_flag(bundle, tmp_path):
    assert install.install_cmd(["--target", str(bundle.root), "--clients", "none", "--yes"]) == 0
    other = tmp_path / "other"
    assert install.install_cmd(["--target", str(other), "--yes"]) == 1
    assert not other.exists()
    assert install.install_cmd(["--target", str(other), "--new-location", "--clients", "none", "--yes"]) == 0


def test_invalid_staging_keeps_old_copy_before_switch(bundle, monkeypatch):
    assert install.install_cmd(["--target", str(bundle.root), "--clients", "none", "--yes"]) == 0
    (bundle.app / install.EXE_NAME).write_text("old")
    monkeypatch.setattr(install, "_starts", lambda exe: 1)
    assert install.install_cmd(["--yes"]) == 1
    assert (bundle.app / install.EXE_NAME).read_text() == "old"


def test_post_switch_failure_restores_old_copy_and_state(bundle, monkeypatch):
    from office_live.state import state_path

    assert install.install_cmd(["--target", str(bundle.root), "--clients", "none", "--yes"]) == 0
    original = state_path().read_bytes()
    (bundle.app / install.EXE_NAME).write_text("old")
    monkeypatch.setattr(install, "_starts", lambda exe: int(exe.parent == bundle.app))
    assert install.install_cmd(["--yes"]) == 1
    assert (bundle.app / install.EXE_NAME).read_text() == "old"
    assert state_path().read_bytes() == original


def test_open_executable_is_reported_without_killing_any_process(bundle, capsys):
    bundle.app.mkdir(parents=True)
    exe = bundle.app / install.EXE_NAME
    exe.write_text("old")
    install._write_marker(bundle.root)
    with exe.open("rb"):
        assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 1
    assert exe.read_text() == "old" and "закройте" in capsys.readouterr().out


def test_failed_staging_cleanup_refuses_a_new_junction(bundle, tmp_path, monkeypatch):
    outside = tmp_path / "outside"
    outside.mkdir()
    (outside / "keep.txt").write_text("keep")
    links = []

    def failed(exe):
        link = exe.parent / "unexpected-link"
        result = subprocess.run(["cmd", "/d", "/c", "mklink", "/J", str(link), str(outside)], capture_output=True, timeout=10)
        assert result.returncode == 0
        links.append(link)
        return 1

    monkeypatch.setattr(install, "_starts", failed)
    try:
        assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 1
        assert (outside / "keep.txt").read_text() == "keep"
        assert links[0].exists()
    finally:
        for link in links:
            os.rmdir(link)


def test_legacy_custom_update_adopts_only_existing_registration(bundle):
    from office_live import registrations as reg
    from office_live.state import read_state

    bundle.app.mkdir(parents=True)
    (bundle.app / install.EXE_NAME).write_text("old")
    install._write_marker(bundle.root)
    spec = reg.target("codex")
    path = Path(spec["path"])
    path.parent.mkdir(parents=True)
    old = {"command": str(bundle.app / install.EXE_NAME), "env": {"OFFICE_LIVE_MODE": "readonly"}, "startup_timeout_sec": 60}
    path.write_text(reg.render(spec, "", {}, old), encoding="utf-8")
    before = path.read_bytes()
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 0
    assert path.read_bytes() == before
    assert [r["client"] for r in read_state()["registrations"]] == ["codex", "protocol"]


def test_install_copies_the_program_checks_it_and_connects_agents(bundle):
    rc = install.install_cmd(["--target", str(bundle.root), "--yes", "--clients", "cursor", "--readonly"])
    assert rc == 0
    assert (bundle.app / install.EXE_NAME).read_text(encoding="utf-8") == "exe v2"
    assert (bundle.app / "_internal" / "lib.pyd").exists()
    # проверяется и подключается УСТАНОВЛЕННАЯ копия (её путь попадёт в конфиги агентов), параметры уходят в setup
    assert bundle.calls == [
        [install.EXE_NAME, "tools", "--json"], [install.EXE_NAME, "tools", "--json"], [install.EXE_NAME, "setup", "--yes", "--clients", "cursor", "--readonly"],
    ]


def test_update_replaces_the_old_copy_completely(bundle):
    (bundle.app / "_internal").mkdir(parents=True)
    (bundle.app / install.EXE_NAME).write_text("exe v1", encoding="utf-8")
    (bundle.app / "_internal" / "stale.pyd").write_text("old", encoding="utf-8")
    install._write_marker(bundle.root)
    install.write_manifest(bundle.app)
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
    assert "закройте" in capsys.readouterr().out and bundle.calls == [[install.EXE_NAME, "tools", "--json"]]


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
    (bundle.app / "office-live-link.exe").write_text("installed GUI", encoding="utf-8")
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
    from office_live import registrations as reg

    root = tmp_path / "root"
    (root / "app").mkdir(parents=True)
    (root / "app" / install.EXE_NAME).write_text("fixture")
    install._write_marker(root)
    monkeypatch.setattr(reg, "identity", lambda: (str(root / "app" / install.EXE_NAME), [], str(root)))
    monkeypatch.setattr(install, "frozen", lambda: False)
    assert cli.setup_cmd(["--clients", "codex,zcode,claude-code", "--yes"]) == 0
    assert install.uninstall_cmd(["--target", str(root)]) == 0
    assert all(reg.read_entry(reg.target(c))[2] is None for c in ("codex", "zcode", "claude-code"))
    assert not root.exists()


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
    from office_live import registrations as reg

    monkeypatch.setattr(reg, "identity", lambda: (str(tmp_path / "app" / install.EXE_NAME), [], str(tmp_path)))
    assert cli.setup_cmd(["--clients", "codex", "--yes"]) == 0
    monkeypatch.setattr(install, "_owned", lambda root: pytest.fail("must not check ownership"))
    assert install.uninstall_cmd(["--target", str(tmp_path), "--keep-files"]) == 0
    assert reg.read_entry(reg.target("codex"))[2] is None


def test_uninstall_missing_root_disconnects_and_reports_no_files(tmp_path, monkeypatch, capsys):
    from office_live import registrations as reg

    root = tmp_path / "missing"
    monkeypatch.setattr(reg, "identity", lambda: (str(root / "app" / install.EXE_NAME), [], str(root)))
    assert cli.setup_cmd(["--clients", "codex", "--yes"]) == 0
    assert install.uninstall_cmd(["--target", str(root)]) == 0
    assert reg.read_entry(reg.target("codex"))[2] is None
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
    root = tmp_path / "install"
    root.mkdir()
    install._write_marker(root)
    (root / "app").mkdir()
    (root / "notes.txt").write_text("keep")
    (root / "app.notes-important").mkdir()
    monkeypatch.setattr(install, "frozen", lambda: False)
    assert install.uninstall_cmd(["--target", str(root)]) == 0
    assert sorted(p.name for p in root.iterdir()) == ["app.notes-important", "notes.txt"]


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
    script = Path(calls[0][-1]).read_text(encoding="utf-8-sig")
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
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 1
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
    assert (bundle.app / install.EXE_NAME).read_text() == "old" and bundle.calls == [[install.EXE_NAME, "tools", "--json"]]


@pytest.mark.parametrize("failure", ["return", "raise"])
def test_update_that_does_not_start_rolls_back_before_setup(bundle, monkeypatch, capsys, failure):
    from office_live import probe

    bundle.app.mkdir(parents=True)
    install._write_marker(bundle.root)
    (bundle.app / install.EXE_NAME).write_text("old")

    def smoke(command, **kwargs):
        assert (bundle.app / install.EXE_NAME).read_text() == "old"
        assert not list(bundle.root.glob("app.old-*"))  # проверка до переключения
        if failure == "raise":
            raise OSError("invalid exe")
        raise probe.ProtocolError("invalid stdout")

    monkeypatch.setattr(install, "_starts", REAL_STARTS)
    monkeypatch.setattr(probe, "smoke", smoke)
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 1
    assert (bundle.app / install.EXE_NAME).read_text() == "old"
    assert "переключение не выполнялось" in capsys.readouterr().out
    assert not any(c[1:2] == ["setup"] for c in bundle.calls)


def test_doctor_failure_alone_does_not_roll_back_an_update(bundle, monkeypatch):
    """doctor «падает» на машине без Word (или без Excel), где программа рабочая: обновление не должно откатываться."""
    bundle.app.mkdir(parents=True)
    install._write_marker(bundle.root)
    (bundle.app / install.EXE_NAME).write_text("old")
    monkeypatch.setattr(install.subprocess, "call", lambda cmd, **kw: 1 if cmd[1] == "doctor" else 0)
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 0
    assert (bundle.app / install.EXE_NAME).read_text(encoding="utf-8") == "exe v2"


def test_first_install_does_not_switch_or_connect_when_staging_does_not_start(bundle, monkeypatch, capsys):
    monkeypatch.setattr(install, "_starts", lambda exe: 1)
    assert install.install_cmd(["--target", str(bundle.root), "--yes", "--clients", "cursor", "--readonly"]) == 1
    assert bundle.calls == [] and not bundle.app.exists()
    marker = json.loads((bundle.root / install.MARKER).read_text())
    assert marker["product"] == install.PRODUCT and marker["created"]
    assert "переключение не выполнялось" in capsys.readouterr().out


def test_uninstall_disconnect_failure_keeps_program_and_marker(tmp_path, monkeypatch):
    from office_live import registrations as reg

    root = tmp_path / "root"
    (root / "app").mkdir(parents=True)
    install._write_marker(root)
    (root / "app" / install.EXE_NAME).write_text("exe")
    monkeypatch.setattr(reg, "identity", lambda: (str(root / "app" / install.EXE_NAME), [], str(root)))
    assert cli.setup_cmd(["--clients", "codex", "--yes"]) == 0
    monkeypatch.setattr(reg, "remove", lambda *a: (False, "не удалось снять"))
    assert install.uninstall_cmd(["--target", str(root)]) == 1
    assert (root / "app" / install.EXE_NAME).read_text() == "exe" and (root / install.MARKER).exists()


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
    import office_live.app as app

    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "stdin", Tty(tty))
    ran = []
    monkeypatch.setattr(cli, "run", lambda cmd, args: ran.append((cmd, args)) or 0)

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


def test_install_does_not_call_doctor_or_office_tools(bundle, monkeypatch):
    seen = []

    def call(command, **kwargs):
        seen.append(command[1])
        assert command[1] != "doctor"
        return 0

    monkeypatch.setattr(install.subprocess, "call", call)
    assert install.install_cmd(["--target", str(bundle.root), "--yes"]) == 0
    assert seen == ["setup"]


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


def test_install_no_links_skips_registry_entirely_and_dry_run_shows_protocol(bundle, fake_registry, capsys):
    from office_live.state import read_state

    assert install.install_cmd(["--target", str(bundle.root), "--clients", "none", "--dry-run"]) == 0
    assert "HKCU\\Software\\Classes\\officelive" in capsys.readouterr().out
    assert "write" not in fake_registry.calls and not read_state()["registrations"]
    fake_registry.calls.clear()
    assert install.install_cmd(["--target", str(bundle.root), "--clients", "none", "--yes", "--no-links"]) == 0
    assert not fake_registry.calls and not read_state()["registrations"]


def test_rename_retries_transient_access_denied(tmp_path, monkeypatch):
    from office_live import install

    src, dst = tmp_path / "app.new-1", tmp_path / "app"
    src.mkdir()
    real, calls = Path.rename, []

    def flaky(self, target):
        calls.append(target)
        if len(calls) < 3:
            raise PermissionError(5, "Access is denied")
        return real(self, target)

    monkeypatch.setattr(Path, "rename", flaky)
    monkeypatch.setattr(install.time, "sleep", lambda s: None)
    install._rename(src, dst)
    assert dst.is_dir() and len(calls) == 3


def test_rename_gives_up_after_attempts(tmp_path, monkeypatch):
    from office_live import install

    src = tmp_path / "app.new-1"
    src.mkdir()
    monkeypatch.setattr(Path, "rename", lambda self, target: (_ for _ in ()).throw(PermissionError(5, "denied")))
    monkeypatch.setattr(install.time, "sleep", lambda s: None)
    with pytest.raises(PermissionError):
        install._rename(src, tmp_path / "app", attempts=3)
