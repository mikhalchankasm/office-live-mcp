"""Контракт Inno/CLI: временные профили, fake EXE/MCP/реестр, без запуска установщика."""

import json
import sys
from pathlib import Path

import pytest

from office_live import __version__, cli, install, installer, registrations as reg
from office_live.install_files import MANIFEST
from office_live.state import read_state, remember_install, state_path, write_state


@pytest.fixture
def copied(tmp_path, monkeypatch):
    root = tmp_path / "Установлено с пробелами"
    app = root / "app"
    (app / "_internal").mkdir(parents=True)
    (app / install.EXE_NAME).write_bytes(b"fake exe")
    (app / "_internal/lib.dll").write_bytes(b"fake library")
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(app / install.EXE_NAME))
    monkeypatch.setattr(installer, "smoke", lambda command: 114)
    monkeypatch.setattr(cli, "_detect_clients", lambda: {"cursor": True, "codex": True})
    return root


def post(capsys, *args, expected=0):
    rc = cli.run("postinstall", ["--json", "--no-links", *args])
    output = capsys.readouterr()
    result = json.loads(output.out)
    assert rc == expected, result
    assert result["ok"] is (expected == 0)
    return result


def connect(client, readonly=False):
    reg.configure(reg.target(client), {"OFFICE_LIVE_MODE": "readonly"} if readonly else {})


def test_clients_json_is_read_only_and_root_specific(copied, capsys, tmp_path):
    connect("codex", True)
    connect("zcode")
    foreign = tmp_path / "other"
    reg.configure(reg.target("cursor"), {}, server=(str(foreign / "app" / install.EXE_NAME), []))
    capsys.readouterr()
    before = {p: p.read_bytes() for p in tmp_path.rglob("*") if p.is_file()}
    assert cli.run("clients", ["--json"]) == 0
    data = json.loads(capsys.readouterr().out)
    clients = {c["id"]: c for c in data["clients"]}
    assert clients["codex"]["registered"] and clients["codex"]["readonly"]
    assert clients["zcode"]["registered"] and not clients["zcode"]["readonly"]
    assert clients["cursor"]["detected"] and not clients["cursor"]["registered"]
    assert data["access"] == "mixed"
    assert before == {p: p.read_bytes() for p in tmp_path.rglob("*") if p.is_file()}


@pytest.mark.parametrize("readonly,access", [(True, "readonly"), (False, "full")])
def test_clients_common_mode(copied, readonly, access):
    connect("codex", readonly)
    connect("zcode", readonly)
    assert installer.clients_info()["access"] == access


def test_clients_with_no_state_does_not_create_it(copied, capsys):
    assert cli.run("clients", ["--json"]) == 0
    data = json.loads(capsys.readouterr().out)
    assert data["access"] == "full" and not any(c["registered"] for c in data["clients"])
    assert not state_path().exists()


def test_keep_preserves_settings_and_does_not_add_detected(copied, capsys):
    connect("codex", True)
    path = Path(reg.target("codex")["path"])
    before = path.read_bytes()
    capsys.readouterr()
    result = post(capsys, "--clients", "keep", "--access", "full")
    assert result["kept"] == ["codex"] and not result["added"] and not result["removed"]
    assert path.read_bytes() == before
    assert not Path(reg.target("cursor")["path"]).exists()
    assert result["smoke_tools"] == 114
    assert json.loads((copied / install.MARKER).read_text())["version"] == __version__
    assert (copied / "app" / MANIFEST).exists() and Path(read_state()["root"]) == copied


def test_explicit_selection_adds_removes_and_keeps_existing_access(copied, capsys):
    connect("codex", True)
    connect("zcode")
    capsys.readouterr()
    result = post(capsys, "--clients", "codex,cursor", "--access", "full")
    assert result["added"] == ["cursor"] and result["removed"] == ["zcode"] and result["kept"] == ["codex"]
    assert reg.read_entry(reg.target("codex"))[2]["env"]["OFFICE_LIVE_MODE"] == "readonly"
    assert "env" not in reg.read_entry(reg.target("cursor"))[2]
    assert reg.read_entry(reg.target("zcode"))[2] is None


def test_none_disconnects_shown_clients_but_preserves_vscode(copied, capsys, tmp_path):
    connect("codex")
    spec = reg.target("vscode", "project", str(tmp_path / "project"))
    reg.configure(spec, {})
    capsys.readouterr()
    result = post(capsys, "--clients", "none")
    assert result["removed"] == ["codex"] and result["kept"] == ["vscode"]
    assert reg.read_entry(spec)[2]


def test_detected_adds_readonly_and_keeps_undetected(copied, capsys):
    connect("zcode")
    capsys.readouterr()
    result = post(capsys, "--clients", "detected", "--access", "readonly")
    assert set(result["added"]) == {"codex", "cursor"} and result["kept"] == ["zcode"]
    assert all(reg.read_entry(reg.target(c))[2]["env"]["OFFICE_LIVE_MODE"] == "readonly" for c in result["added"])


def test_adopts_legacy_before_selection_so_access_is_not_reset(copied, capsys):
    spec = reg.target("codex")
    path = Path(spec["path"])
    path.parent.mkdir(parents=True)
    entry = {"command": sys.executable, "args": [], "env": {"OFFICE_LIVE_MODE": "readonly"}}
    path.write_text(reg.render(spec, "", {}, entry), encoding="utf-8")
    before = path.read_bytes()
    assert cli.run("clients", ["--json"]) == 0
    info = json.loads(capsys.readouterr().out)
    assert next(c for c in info["clients"] if c["id"] == "codex")["registered"]
    assert info["access"] == "readonly" and not state_path().exists() and path.read_bytes() == before
    result = post(capsys, "--clients", "codex", "--access", "full")
    assert result["kept"] == ["codex"] and path.read_bytes() == before


def test_smoke_failure_leaves_profiles_and_metadata_untouched(copied, capsys, monkeypatch):
    def fail(command):
        raise RuntimeError("MCP failure")

    monkeypatch.setattr(installer, "smoke", fail)
    result = post(capsys, "--clients", "codex", expected=1)
    assert result["problems"] == ["MCP failure"] and result["smoke_tools"] == 0
    assert not state_path().exists() and not (copied / install.MARKER).exists()
    assert not (copied / "app" / MANIFEST).exists() and not Path(reg.target("codex")["path"]).exists()


def test_postinstall_rejects_source_execution(capsys):
    result = post(capsys, expected=2)
    assert "frozen" in result["problems"][0] and not state_path().exists()


def test_postinstall_rejects_wrong_frozen_layout(copied, capsys, monkeypatch):
    monkeypatch.setattr(sys, "executable", str(copied / "office-live-mcp.exe"))
    post(capsys, expected=2)
    assert not (copied / install.MARKER).exists()


def test_bad_selection_is_rejected_before_smoke_or_writes(copied, capsys, monkeypatch):
    monkeypatch.setattr(installer, "smoke", lambda command: pytest.fail("must not run"))
    post(capsys, "--clients", "vscode", expected=1)
    assert not (copied / install.MARKER).exists()


def test_modified_registration_reports_problem_without_overwriting(copied, capsys):
    connect("codex")
    path = Path(reg.target("codex")["path"])
    path.write_text(path.read_text() + "timeout = 99\n")
    before = path.read_bytes()
    capsys.readouterr()
    result = post(capsys, "--clients", "none", expected=1)
    assert str(path) in result["problems"][0] and result["removed"] == []
    assert path.read_bytes() == before


def test_protocol_failure_is_json_and_does_not_block_other_clients(copied, capsys, monkeypatch):
    from office_live import protocol

    monkeypatch.setattr(protocol, "configure", lambda **kw: False)
    assert cli.run("postinstall", ["--json", "--clients", "codex"]) == 1
    result = json.loads(capsys.readouterr().out)
    assert result["added"] == ["codex"] and "officelive://" in result["problems"][0]


@pytest.mark.parametrize("args,tail", [([], []), (["--yes"], ["/SILENT"]), (["--dry-run"], None)])
def test_uninstall_delegates_without_lock_or_file_removal(copied, monkeypatch, args, tail):
    uninstaller = copied / "unins000.exe"
    uninstaller.write_bytes(b"fake uninstaller")
    calls = []
    monkeypatch.setattr(install.subprocess, "Popen", lambda command, **kw: calls.append(command))
    monkeypatch.setattr(install, "file_lock", lambda *args: pytest.fail("delegation must not hold maintenance"))
    assert install.uninstall_cmd(args) == 0
    assert calls == ([] if tail is None else [[str(uninstaller), *tail]])
    assert (copied / "app" / install.EXE_NAME).exists() and uninstaller.exists()


def test_keep_files_forgets_only_own_root_and_returns_remaining_problems(copied, capsys):
    remember_install(copied)
    connect("codex")
    path = Path(reg.target("codex")["path"])
    path.write_text("[broken")
    capsys.readouterr()
    assert install.uninstall_cmd(["--target", str(copied), "--keep-files", "--yes", "--json"]) == 1
    result = json.loads(capsys.readouterr().out)
    assert str(path) in result["problems"][0] and not result["ok"]
    assert read_state()["root"] == "" and read_state()["registrations"]
    assert state_path().with_suffix(".json.bak").exists()
    assert (copied / "app" / install.EXE_NAME).exists()


def test_keep_files_does_not_forget_another_installation(copied, tmp_path):
    other = tmp_path / "other"
    remember_install(other)
    assert install.uninstall_cmd(["--target", str(copied), "--keep-files"]) == 0
    assert Path(read_state()["root"]) == other


def test_keep_files_dry_run_does_not_change_state(copied):
    remember_install(copied)
    before = state_path().read_bytes()
    assert install.uninstall_cmd(["--keep-files", "--target", str(copied), "--dry-run"]) == 0
    assert state_path().read_bytes() == before


def test_zip_install_warns_about_inno_but_target_still_works(copied, monkeypatch, capsys):
    install._write_marker(copied)
    (copied / "unins000.exe").write_bytes(b"fake uninstaller")
    monkeypatch.setattr(install, "_office_installed", lambda progid: True)
    assert install.install_cmd(["--target", str(copied), "--dry-run", "--no-links", "--clients", "none", "--yes"]) == 0
    assert "setup.exe" in capsys.readouterr().out


@pytest.mark.parametrize("command,args", [("setup", ["--clients", "none", "--yes"]), ("doctor", [])])
def test_pause_waits_once_and_preserves_exit_code(monkeypatch, command, args):
    prompts = []
    monkeypatch.setattr("builtins.input", lambda prompt: prompts.append(prompt) or "")
    monkeypatch.setattr(cli, "doctor", lambda: 1)
    assert cli.run(command, [*args, "--pause"]) == (1 if command == "doctor" else 0)
    assert len(prompts) == 1 and "Enter" in prompts[0]


def test_pause_handles_eof_and_errors(monkeypatch):
    def eof(prompt):
        raise EOFError

    monkeypatch.setattr("builtins.input", eof)
    assert cli.setup_cmd(["--clients", "invalid", "--yes", "--pause"]) == 2


def test_clients_corrupt_ledger_returns_json_failure(copied, capsys):
    write_state({"schema": 123, "root": "", "registrations": []})
    assert cli.run("clients", ["--json"]) == 1
    assert json.loads(capsys.readouterr().out)["problems"]
