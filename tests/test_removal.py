"""Отключение и удаление только временной установки."""

import json
import os
import subprocess
from pathlib import Path

import pytest

from office_live import cli, install, registrations as reg
from office_live.install_files import removal_plan, write_manifest
from office_live.state import read_state, remember_install


@pytest.fixture
def installed(tmp_path, monkeypatch):
    root = tmp_path / "установка с пробелами"
    app = root / "app"
    app.mkdir(parents=True)
    exe = app / install.EXE_NAME
    exe.write_bytes(b"fixture exe")
    (app / "lib.dll").write_bytes(b"fixture lib")
    write_manifest(app)
    install._write_marker(root)
    remember_install(root)
    monkeypatch.setattr(reg, "identity", lambda: (str(exe), [], str(root)))
    monkeypatch.setattr(install, "frozen", lambda: False)
    return root


def setup(client, *args):
    return cli.setup_cmd(["--yes", "--clients", client, *args])


def test_disconnect_one_keeps_program_and_other_client(installed):
    assert setup("codex,zcode") == 0
    assert setup("zcode", "--remove") == 0
    assert reg.read_entry(reg.target("codex"))[2]
    assert reg.read_entry(reg.target("zcode"))[2] is None
    assert (installed / "app" / install.EXE_NAME).exists()
    assert [r["client"] for r in read_state()["registrations"]] == ["codex"]


def test_uninstall_all_registered_scopes_keeps_foreign_files(installed, tmp_path):
    assert setup("codex,zcode") == 0
    assert setup("cursor", "--scope", "project", "--project", str(tmp_path / "проект")) == 0
    custom = tmp_path / "custom.json"
    custom.write_text('{"mcpServers":{"other":{"command":"keep"}},"theme":"dark"}')
    assert setup("claude-code", "--config", str(custom)) == 0
    (installed / "notes.txt").write_text("keep")
    (installed / "app/user.txt").write_text("keep")
    assert install.uninstall_cmd([]) == 0
    assert not read_state()["registrations"]
    assert (installed / "app/user.txt").read_text() == "keep"
    assert (installed / "notes.txt").exists()
    assert not (installed / "app" / install.EXE_NAME).exists()
    assert json.loads(custom.read_text()) == {"mcpServers": {"other": {"command": "keep"}}, "theme": "dark"}


def test_changed_dependent_registration_keeps_exe(installed, capsys):
    assert setup("cursor") == 0
    path = Path(reg.target("cursor")["path"])
    data = json.loads(path.read_text(encoding="utf-8"))
    data["mcpServers"]["office-live"]["userSetting"] = True
    path.write_text(json.dumps(data))
    before = path.read_bytes()
    assert install.uninstall_cmd([]) == 1
    assert path.read_bytes() == before and (installed / "app" / install.EXE_NAME).exists()
    assert "Файлы сохранены" in capsys.readouterr().out


def test_changed_foreign_command_is_left_but_no_longer_blocks_files(installed):
    assert setup("cursor") == 0
    path = Path(reg.target("cursor")["path"])
    path.write_text('{"mcpServers":{"office-live":{"command":"foreign.exe"}}}')
    assert install.uninstall_cmd([]) == 0
    assert "foreign.exe" in path.read_text()
    assert not installed.exists()


def test_config_permission_failure_keeps_exe_and_continues(installed, monkeypatch):
    assert setup("cursor,codex") == 0
    original = reg.remove

    def remove(spec, *args):
        if spec["client"] == "cursor":
            raise PermissionError("занят файл")
        return original(spec, *args)

    monkeypatch.setattr(reg, "remove", remove)
    assert install.uninstall_cmd([]) == 1
    assert reg.read_entry(reg.target("codex"))[2] is None
    assert (installed / "app" / install.EXE_NAME).exists()


def test_legacy_entry_without_state_is_removed_only_for_this_exe(installed):
    spec = reg.target("zcode")
    path = Path(spec["path"])
    path.parent.mkdir(parents=True)
    old = {"command": str(installed / "app" / install.EXE_NAME), "env": {"OFFICE_LIVE_MODE": "readonly"}}
    path.write_text(reg.render(spec, "", {}, old), encoding="utf-8")
    assert install.uninstall_cmd([]) == 0
    assert reg.read_entry(spec)[2] is None


def test_modified_program_file_is_preserved(installed):
    (installed / "app/lib.dll").write_text("user modified")
    assert install.uninstall_cmd([]) == 0
    assert (installed / "app/lib.dll").read_text() == "user modified"


def test_manifest_cannot_escape_installation(installed, tmp_path):
    outside = tmp_path / "keep.txt"
    outside.write_text("keep")
    (installed / "app/office-live-mcp.files.json").write_text('{"../../keep.txt":"hash"}')
    assert install.uninstall_cmd([]) == 1
    assert outside.read_text() == "keep"


def test_nested_junction_refuses_before_disconnect(installed, tmp_path):
    assert setup("codex") == 0
    outside = tmp_path / "outside"
    outside.mkdir()
    (outside / "keep.txt").write_text("keep")
    junction = installed / "app/link"
    # Только создание ссылки; удаление выполняется Path.unlink/rmdir, без рекурсивного shell.
    result = subprocess.run(["cmd", "/d", "/c", "mklink", "/J", str(junction), str(outside)], capture_output=True, timeout=10)
    assert result.returncode == 0
    try:
        assert install.uninstall_cmd([]) == 1
        assert reg.read_entry(reg.target("codex"))[2]
        assert (outside / "keep.txt").read_text() == "keep"
    finally:
        os.rmdir(junction)


def test_detached_cleanup_has_exact_report(installed, tmp_path):
    (installed / "app/user.txt").write_text("keep")
    report = tmp_path / "report.json"
    script = tmp_path / "cleanup.ps1"
    script.write_text(install._self_delete_script(installed, report), encoding="utf-8-sig")
    assert installed.resolve().is_relative_to(tmp_path.resolve())
    result = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(script)],
                            capture_output=True, timeout=20)
    assert result.returncode == 0, result.stderr.decode("utf-8", "replace")
    data = json.loads(report.read_text(encoding="utf-8-sig"))
    assert data["errors"] == [] and str(installed / "app/user.txt") in data["remaining"]
    assert not (installed / "app" / install.EXE_NAME).exists()


def test_removal_plan_is_read_only(installed):
    before = {str(p): p.read_bytes() for p in installed.rglob("*") if p.is_file()}
    own, keep = removal_plan(installed, installed / "app")
    assert len(own) == 3 and not keep
    assert before == {str(p): p.read_bytes() for p in installed.rglob("*") if p.is_file()}


def test_partial_file_failure_reports_actual_remaining_files(installed, monkeypatch, capsys):
    real = Path.unlink
    failed = installed / "app/lib.dll"

    def locked(path, *args, **kwargs):
        if path == failed:
            raise PermissionError("locked lib.dll")
        return real(path, *args, **kwargs)

    monkeypatch.setattr(Path, "unlink", locked)
    assert install.uninstall_cmd([]) == 1
    assert failed.exists() and (installed / install.MARKER).exists()
    output = capsys.readouterr().out
    assert "Удаление неполное" in output and str(failed) in output and "повторите uninstall" in output


def test_junction_above_install_root_is_allowed_but_inside_is_refused(tmp_path):
    import subprocess

    from office_live import install_files

    real = tmp_path / "real"
    (real / "inst" / "app").mkdir(parents=True)
    link = tmp_path / "redirected"
    subprocess.run(["cmd", "/c", "mklink", "/J", str(link), str(real)], check=True, capture_output=True)
    root = link / "inst"
    (root / "app" / "office-live-mcp.exe").write_bytes(b"exe")
    install_files.check_tree(root, root / "app")  # перенаправленный предок (как AppData) не мешает
    inner = root / "app" / "inner"
    (tmp_path / "elsewhere").mkdir()
    subprocess.run(["cmd", "/c", "mklink", "/J", str(inner), str(tmp_path / "elsewhere")], check=True, capture_output=True)
    with pytest.raises(ValueError, match="reparse"):
        install_files.check_tree(root, root / "app")
