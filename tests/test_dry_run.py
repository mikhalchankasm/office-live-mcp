"""Полный снимок изолированного профиля: dry-run не создаёт даже замков/копий/состояния."""

import sys
from pathlib import Path

from office_live import cli, install


def snapshot(root):
    return {str(p.relative_to(root)): (p.read_bytes() if p.is_file() else None) for p in root.rglob("*")}


def test_setup_dry_run_first_repeat_and_remove(tmp_path):
    before = snapshot(tmp_path)
    assert cli.setup_cmd(["--yes", "--clients", "codex,zcode", "--dry-run"]) == 0
    assert snapshot(tmp_path) == before
    assert cli.setup_cmd(["--yes", "--clients", "codex,zcode"]) == 0
    before = snapshot(tmp_path)
    assert cli.setup_cmd(["--yes", "--clients", "codex,zcode", "--readonly", "--dry-run"]) == 0
    assert cli.setup_cmd(["--yes", "--clients", "codex,zcode", "--remove", "--dry-run"]) == 0
    assert snapshot(tmp_path) == before


def test_install_and_uninstall_dry_run_snapshot(tmp_path, monkeypatch):
    src = tmp_path / "archive/app"
    src.mkdir(parents=True)
    (src / install.EXE_NAME).write_text("fixture")
    monkeypatch.setattr(install, "frozen", lambda: True)
    monkeypatch.setattr(sys, "executable", str(src / install.EXE_NAME))
    monkeypatch.setattr(install, "_office_installed", lambda name: True)
    target = tmp_path / "установка"
    before = snapshot(tmp_path)
    assert install.install_cmd(["--target", str(target), "--clients", "codex,zcode", "--dry-run"]) == 0
    assert snapshot(tmp_path) == before
    (target / "app").mkdir(parents=True)
    (target / "app" / install.EXE_NAME).write_text("fixture")
    install._write_marker(target)
    assert cli.setup_cmd(["--yes", "--clients", "cursor"], server=(str(target / "app" / install.EXE_NAME), [])) == 0
    before = snapshot(tmp_path)
    assert install.install_cmd(["--target", str(target), "--dry-run"]) == 0
    assert install.uninstall_cmd(["--target", str(target), "--dry-run"]) == 0
    assert snapshot(tmp_path) == before


def test_project_custom_dry_run_creates_no_directories(tmp_path):
    before = snapshot(tmp_path)
    assert cli.setup_cmd(["--clients", "cursor", "--scope", "project", "--project", str(tmp_path / "new"), "--dry-run"]) == 0
    assert cli.setup_cmd(["--clients", "codex", "--config", str(tmp_path / "custom/file.toml"), "--dry-run"]) == 0
    assert snapshot(tmp_path) == before


def test_all_profile_environment_variables_are_temporary(tmp_path):
    import os

    for key in ("USERPROFILE", "HOME", "APPDATA", "LOCALAPPDATA", "CODEX_HOME"):
        assert Path(os.environ[key]).is_relative_to(tmp_path)
