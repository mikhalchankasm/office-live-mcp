"""Учёт регистраций: тестовые профили, проекты и конфликты; клиентские приложения не запускаются."""

import json
from pathlib import Path

import pytest

from office_live import cli, registrations as reg
from office_live.state import file_lock, read_state, state_path


def setup(client, *args):
    return cli.setup_cmd(["--yes", "--clients", client, *args])


def entry(client, **kwargs):
    return reg.read_entry(reg.target(client, **kwargs))[2]


def test_every_default_path_uses_isolated_environment(tmp_path):
    for client in cli.CLIENT_LABELS:
        if client != "vscode":
            assert Path(reg.target(client)["path"]).is_relative_to(tmp_path)
    assert state_path().is_relative_to(tmp_path)


def test_two_clients_one_copy_independent_access_and_late_setup():
    assert setup("codex", "--full") == 0
    codex = entry("codex")
    assert setup("zcode", "--readonly") == 0
    assert entry("codex") == codex
    assert entry("zcode")["command"] == codex["command"]
    assert entry("zcode")["env"]["OFFICE_LIVE_MODE"] == "readonly"
    assert setup("codex,zcode") == 0
    assert entry("codex") == codex and len(read_state()["registrations"]) == 2
    assert entry("zcode")["env"]["OFFICE_LIVE_MODE"] == "readonly"


@pytest.mark.parametrize("client", list(cli.CLIENT_LABELS))
def test_preserves_additional_parameters_and_env(client, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    spec = reg.target(client)
    path = Path(spec["path"])
    path.parent.mkdir(parents=True, exist_ok=True)
    expected = {**cli._json_server_entry({"OFFICE_LIVE_MODE": "readonly", "OTHER": "keep"}, spec["typed"]),
                "timeout": 37, "enabled": False, "extra": {"items": ["a", "b"]}}
    path.write_text(reg.render(spec, "", {}, expected), encoding="utf-8")
    assert setup(client) == 0
    assert reg.read_entry(spec)[2] == expected
    record = read_state()["registrations"][0]
    assert record["fingerprint"] == reg.fingerprint(expected) and record["status"] == "confirmed"
    assert record["time"] and record["command"]


@pytest.mark.parametrize("client,relative", [("claude-code", ".mcp.json"), ("cursor", ".cursor/mcp.json"),
                                            ("codex", ".codex/config.toml"), ("zcode", ".zcode/config.json"), ("vscode", ".vscode/mcp.json")])
def test_official_project_paths(client, relative, tmp_path):
    project = tmp_path / "Проект с пробелами"
    assert setup(client, "--scope", "project", "--project", str(project)) == 0
    assert (project / relative).exists()
    assert read_state()["registrations"][0]["scope"] == "project"


def test_codex_home_and_custom_path(tmp_path, monkeypatch):
    monkeypatch.setenv("CODEX_HOME", str(tmp_path / "Особый Codex"))
    assert setup("codex") == 0
    assert (tmp_path / "Особый Codex/config.toml").exists()
    path = tmp_path / "нестандартный.json"
    assert setup("zcode", "--config", str(path)) == 0
    assert read_state()["registrations"][-1]["scope"] == "custom"


def test_scope_refusals_do_not_guess(tmp_path):
    assert setup("claude-desktop", "--scope", "project", "--project", str(tmp_path)) == 1
    assert setup("cursor", "--scope", "project") == 2
    assert setup("codex,zcode", "--config", str(tmp_path / "config")) == 2
    assert not state_path().exists()


def test_foreign_entry_requires_force_even_with_reset(tmp_path):
    spec = reg.target("cursor")
    path = Path(spec["path"])
    path.parent.mkdir(parents=True)
    old = '{"mcpServers":{"office-live":{"command":"foreign.exe"},"other":{"command":"keep"}}}'
    path.write_text(old, encoding="utf-8")
    assert setup("cursor", "--reset") == 1 and path.read_text(encoding="utf-8") == old
    assert setup("cursor", "--force") == 0
    assert json.loads(path.read_text())["mcpServers"]["other"] == {"command": "keep"}


def test_same_python_other_script_is_foreign():
    spec = reg.target("cursor")
    path = Path(spec["path"])
    path.parent.mkdir(parents=True)
    path.write_text(json.dumps({"mcpServers": {"office-live": {"command": cli._server_command()[0], "args": ["different.py"]}}}))
    assert setup("cursor") == 1


@pytest.mark.parametrize("client,content", [("cursor", "{broken"), ("zcode", '{"mcp":42}'), ("codex", 'model = "unterminated')])
def test_corrupt_config_survives_force_and_reset(client, content):
    path = Path(reg.target(client)["path"])
    path.parent.mkdir(parents=True)
    path.write_text(content, encoding="utf-8")
    assert setup(client, "--force", "--reset") == 1
    assert path.read_text(encoding="utf-8") == content


def test_partial_failure_records_successful_clients():
    path = Path(reg.target("zcode")["path"])
    path.parent.mkdir(parents=True)
    path.write_text("{broken")
    assert setup("codex,zcode,cursor") == 1
    assert {r["client"] for r in read_state()["registrations"]} == {"codex", "cursor"}


@pytest.mark.parametrize("always", [False, True])
def test_concurrent_change_is_re_read_once(monkeypatch, always):
    assert setup("cursor") == 0
    path = Path(reg.target("cursor")["path"])
    backup = cli._backup
    calls = []

    def changed(p):
        result = backup(p)
        if not calls or always:
            data = json.loads(p.read_text(encoding="utf-8"))
            data["user-change"] = len(calls) + 1
            p.write_text(json.dumps(data), encoding="utf-8")
        calls.append(p)
        return result

    monkeypatch.setattr(cli, "_backup", changed)
    assert setup("cursor", "--readonly") == (1 if always else 0)
    assert json.loads(path.read_text())["user-change"] == (2 if always else 1)


def test_lock_is_bounded_and_released(tmp_path):
    path = tmp_path / "config.json"
    with file_lock(path):
        with pytest.raises(OSError, match="занят"):
            with file_lock(path, timeout=0.1):
                pytest.fail("second writer acquired lock")
    with file_lock(path, timeout=0.1):
        pass


def test_state_write_failure_never_leaves_untracked_config(monkeypatch):
    monkeypatch.setattr(reg, "write_state", lambda data: (_ for _ in ()).throw(OSError("disk full")))
    assert setup("codex") == 1
    assert not Path(reg.target("codex")["path"]).exists()


def test_readback_failure_retains_pending_registration(monkeypatch):
    real = cli._atomic_write
    path = Path(reg.target("cursor")["path"])

    def change(p, text, **kwargs):
        real(p, text, **kwargs)
        if p == path:
            p.write_text('{"mcpServers":{}}')

    monkeypatch.setattr(cli, "_atomic_write", change)
    assert setup("cursor") == 1
    assert read_state()["registrations"][0]["status"] == "pending"


def test_long_unicode_custom_path(tmp_path):
    path = tmp_path
    for i in range(5):
        path /= f"Каталог с пробелами {i} " + "x" * 24
    path /= "клиент.json"
    assert len(str(path)) > 300
    assert setup("cursor", "--config", str(path), "--readonly") == 0
    assert entry("cursor", config=str(path))["env"]["OFFICE_LIVE_MODE"] == "readonly"
    assert setup("cursor", "--config", str(path), "--remove") == 0


def test_corrupt_state_does_not_write_config():
    state_path().parent.mkdir(parents=True)
    state_path().write_text('{"schema":1,"root":"","registrations":[{}]}')
    assert setup("cursor", "--dry-run") == 1
    assert setup("cursor") == 1
    assert not Path(reg.target("cursor")["path"]).exists()


def test_force_changes_http_transport_to_stdio():
    path = Path(reg.target("cursor")["path"])
    path.parent.mkdir(parents=True)
    path.write_text('{"mcpServers":{"office-live":{"type":"http","url":"https://example.test","timeout":20}}}')
    assert setup("cursor", "--force") == 0
    result = entry("cursor")
    assert result["type"] == "stdio" and "url" not in result and result["timeout"] == 20
