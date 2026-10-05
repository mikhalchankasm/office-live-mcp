"""Мастер подключения `setup` и `config`: пишем только во временный «домашний» каталог — реальные конфиги агентов не трогаем."""

import json
from pathlib import Path

import pytest

from office_live import cli


@pytest.fixture
def fake_home(tmp_path, monkeypatch):
    home = tmp_path / "home"
    appdata = home / "AppData" / "Roaming"
    appdata.mkdir(parents=True)
    monkeypatch.setattr(Path, "home", classmethod(lambda cls: home))
    monkeypatch.setenv("APPDATA", str(appdata))
    monkeypatch.setattr(cli.shutil, "which", lambda name: None)  # claude CLI «не установлен» — тест не должен его вызывать
    return home


def _servers(path: Path, keys: list[str]) -> dict:
    node = json.loads(path.read_text(encoding="utf-8"))
    for k in keys:
        node = node[k]
    return node


def test_parse_selection():
    detected = {"claude-code": True, "cursor": True, "zcode": False}
    assert cli._parse_selection("", detected) == ["claude-code", "cursor"]
    assert cli._parse_selection("0", detected) == []
    assert cli._parse_selection("none", detected) == []
    assert cli._parse_selection("2, 3", detected) == ["claude-desktop", "cursor"]
    assert cli._parse_selection("cursor zcode", detected) == ["cursor", "zcode"]
    assert cli._parse_selection("2,2,2", detected) == ["claude-desktop"]  # без дублей
    assert "vscode" not in cli._parse_selection("all", detected)  # VS Code — только по явному выбору
    with pytest.raises(ValueError):
        cli._parse_selection("banana", detected)


def test_setup_writes_entries_and_keeps_other_servers(fake_home):
    cursor = fake_home / ".cursor" / "mcp.json"
    cursor.parent.mkdir(parents=True)
    cursor.write_text(json.dumps({"mcpServers": {"notion": {"url": "https://example.test/mcp"}}}), encoding="utf-8")
    zcode = fake_home / ".zcode" / "cli" / "config.json"
    zcode.parent.mkdir(parents=True)
    zcode.write_text(json.dumps({"mcp": {"servers": {"other": {"command": "x"}}}, "plugins": {"a": 1}}), encoding="utf-8")

    assert cli.run("setup", ["--yes", "--clients", "cursor,zcode,claude-desktop"]) == 0

    s = _servers(cursor, ["mcpServers"])
    assert set(s) == {"notion", "office-live"} and s["office-live"]["command"].lower().endswith("python.exe")
    assert s["office-live"]["args"][0].endswith("server.py") and "env" not in s["office-live"]
    z = json.loads(zcode.read_text(encoding="utf-8"))
    assert set(z["mcp"]["servers"]) == {"other", "office-live"} and z["plugins"] == {"a": 1}  # чужое не тронуто
    desktop = fake_home / "AppData" / "Roaming" / "Claude" / "claude_desktop_config.json"
    assert "office-live" in _servers(desktop, ["mcpServers"])
    assert list(cursor.parent.glob("mcp.json.bak-*"))  # датированная резервная копия


def test_setup_is_idempotent_and_applies_readonly(fake_home):
    for _ in range(2):
        assert cli.run("setup", ["--yes", "--clients", "cursor", "--readonly"]) == 0
    s = _servers(fake_home / ".cursor" / "mcp.json", ["mcpServers"])
    assert list(s) == ["office-live"] and s["office-live"]["env"] == {"OFFICE_LIVE_MODE": "readonly"}


def test_reinstall_keeps_restrictions_unless_told_otherwise(fake_home):
    """Регрессия аудита: повторный `setup` без флагов молча снимал readonly/allowed-dirs."""
    path = fake_home / ".cursor" / "mcp.json"
    docs = r"D:\Docs"
    assert cli.run("setup", ["--yes", "--clients", "cursor", "--readonly", "--allowed-dirs", docs]) == 0
    assert cli.run("setup", ["--yes", "--clients", "cursor"]) == 0  # без флагов — ничего не теряем
    env = _servers(path, ["mcpServers"])["office-live"]["env"]
    assert env == {"OFFICE_LIVE_MODE": "readonly", "OFFICE_LIVE_ALLOWED_DIRS": docs}
    assert cli.run("setup", ["--yes", "--clients", "cursor", "--full"]) == 0  # явное расширение прав
    assert _servers(path, ["mcpServers"])["office-live"]["env"] == {"OFFICE_LIVE_ALLOWED_DIRS": docs}
    assert cli.run("setup", ["--yes", "--clients", "cursor", "--reset"]) == 0  # явный сброс
    assert "env" not in _servers(path, ["mcpServers"])["office-live"]


def test_config_rejects_contradictions_and_foreign_env(fake_home):
    assert cli.run("config", ["cursor", "--readonly", "--full"]) == 2
    assert cli.run("config", ["cursor", "--env", "PATH=x"]) == 2
    assert cli.run("config", ["cursor", "--env", "OFFICE_LIVE_AUTOSAVE=allow", "--quiet"]) == 0


def test_codex_env_survives_reinstall_and_headers_are_parsed_robustly(fake_home):
    cfg = fake_home / ".codex" / "config.toml"
    cfg.parent.mkdir(parents=True)
    cfg.write_text(
        '# my settings\nmodel = "x"\n\n'
        '[ mcp_servers."office-live" ]   # managed by hand\ncommand = "old"\nargs = ["a", "b"]\n'
        '[mcp_servers."office-live".env]\nOFFICE_LIVE_MODE = "readonly"\n\n'
        '[mcp_servers.other]\ncommand = "keep"\n',
        encoding="utf-8",
    )
    assert cli.run("setup", ["--yes", "--clients", "codex"]) == 0
    text = cfg.read_text(encoding="utf-8")
    assert text.count("[mcp_servers.office-live]") == 1 and text.count("[mcp_servers.office-live.env]") == 1  # по одному разу
    assert 'command = "old"' not in text
    assert 'OFFICE_LIVE_MODE = "readonly"' in text and "# my settings" in text and 'command = "keep"' in text
    assert list(cfg.parent.glob("config.toml.bak-*"))


def test_strip_toml_tables_does_not_touch_neighbours():
    text = '[mcp_servers.office-live-other]\ncommand = "x"\n[mcp_servers.office-live]\ncommand = "y"\n[mcp_servers.z]\nq = [1, 2]\n'
    out = cli._strip_toml_tables(text)
    assert "office-live-other" in out and 'command = "y"' not in out and "[mcp_servers.z]" in out


def test_json_write_is_atomic_and_keeps_every_backup(fake_home, monkeypatch):
    path = fake_home / ".cursor" / "mcp.json"
    path.parent.mkdir(parents=True)
    path.write_text(json.dumps({"mcpServers": {"keep": {"command": "x"}}}), encoding="utf-8")
    stamps = iter(["20260101-000001", "20260101-000002"])
    monkeypatch.setattr(cli.time, "strftime", lambda fmt: next(stamps))
    cli.run("setup", ["--yes", "--clients", "cursor"])
    cli.run("setup", ["--yes", "--clients", "cursor", "--readonly"])
    assert len(list(path.parent.glob("mcp.json.bak-*"))) == 2  # первая копия не затёрта второй
    assert not list(path.parent.glob("*.tmp"))

    def boom(src, dst):
        raise OSError("disk full")

    monkeypatch.setattr(cli.os, "replace", boom)
    monkeypatch.setattr(cli.time, "strftime", lambda fmt: "20260101-000003")
    before = path.read_text(encoding="utf-8")
    with pytest.raises(OSError):
        cli._write_json_with_backup(path, {"x": 1})
    assert path.read_text(encoding="utf-8") == before and not list(path.parent.glob("*.tmp"))


def test_claude_code_restores_previous_entry_when_add_fails(fake_home, monkeypatch):
    claude_json = fake_home / ".claude.json"
    old = {"type": "stdio", "command": "python", "args": ["server.py"], "env": {"OFFICE_LIVE_MODE": "readonly"}}
    claude_json.write_text(json.dumps({"mcpServers": {"office-live": old}}), encoding="utf-8")
    monkeypatch.setattr(cli.shutil, "which", lambda name: "claude.exe")
    calls = []

    def fake_call(parts, **kw):
        calls.append(parts)
        return 1 if parts[:3] == ["claude", "mcp", "add"] else 0  # обычное добавление «падает»

    monkeypatch.setattr(cli.subprocess, "call", fake_call)
    assert cli.config_cmd(["claude-code", "--write", "--quiet"]) == 1
    assert calls[0][:3] == ["claude", "mcp", "remove"]
    add = next(c for c in calls if c[:3] == ["claude", "mcp", "add"])
    assert "OFFICE_LIVE_MODE=readonly" in add  # прежние ограничения перенесены в новую запись
    assert calls[-1][:3] == ["claude", "mcp", "add-json"] and json.loads(calls[-1][-1]) == old  # и прежняя запись возвращена


def test_setup_codex_replaces_existing_block(fake_home):
    cfg = fake_home / ".codex" / "config.toml"
    cfg.parent.mkdir(parents=True)
    cfg.write_text('model = "x"\n\n[mcp_servers.office-live]\ncommand = "old"\nargs = ["old.py"]\n[mcp_servers.office-live.env]\nA = "1"\n\n[mcp_servers.other]\ncommand = "keep"\n', encoding="utf-8")
    assert cli.run("setup", ["--yes", "--clients", "codex", "--readonly"]) == 0
    text = cfg.read_text(encoding="utf-8")
    assert text.count("[mcp_servers.office-live]") == 1 and 'command = "old"' not in text
    assert 'model = "x"' in text and "[mcp_servers.other]" in text and 'command = "keep"' in text
    assert 'OFFICE_LIVE_MODE = "readonly"' in text


def test_setup_remove_undoes_everything(fake_home):
    cli.run("setup", ["--yes", "--clients", "cursor,codex,claude-desktop"])
    assert cli.run("setup", ["--remove", "--yes", "--clients", "cursor,codex,claude-desktop"]) == 0
    assert "office-live" not in _servers(fake_home / ".cursor" / "mcp.json", ["mcpServers"])
    assert "office-live" not in (fake_home / ".codex" / "config.toml").read_text(encoding="utf-8")
    assert cli._remove_client("cursor") == (True, "записи не было")  # повторное удаление безопасно


def test_remove_reports_failures_and_continues_other_clients(fake_home, monkeypatch, capsys):
    seen = []

    def remove(client):
        seen.append(client)
        if client == "cursor":
            raise PermissionError("locked config")
        return client != "codex", "не удалено" if client == "codex" else "удалено"

    monkeypatch.setattr(cli, "_remove_client", remove)
    assert cli.setup_cmd(["--remove", "--yes", "--clients", "cursor,codex,zcode"]) == 1
    assert seen == ["cursor", "codex", "zcode"]
    assert "Cursor, Codex CLI" in capsys.readouterr().out


def test_remove_unsafe_toml_reports_structured_failure(fake_home, monkeypatch):
    path = fake_home / ".codex" / "config.toml"
    path.parent.mkdir()
    path.write_text('[mcp_servers.office-live]\ncommand="old"\n')
    monkeypatch.setattr(cli, "_toml_rewrite_problem", lambda *a: "unsafe TOML")
    assert cli._remove_client("codex")[0] is False
    assert cli.setup_cmd(["--remove", "--yes", "--clients", "codex"]) == 1
    assert 'command="old"' in path.read_text()


@pytest.mark.parametrize("content", [
    '[mcp_servers]\noffice-live = { command = "old" }\nother = { command = "keep" }\n',
    'mcp_servers.office-live.command = "old"\nmcp_servers.other.command = "keep"\n',
])
def test_remove_codex_inline_and_dotted_entries_reports_failure_without_changes(fake_home, content):
    path = fake_home / ".codex" / "config.toml"
    path.parent.mkdir()
    path.write_text(content, encoding="utf-8")
    before = path.read_bytes()
    ok, message = cli._remove_client("codex")
    assert ok is False and message.startswith("не удалено:") and "вручную" in message
    assert cli.setup_cmd(["--remove", "--yes", "--clients", "codex"]) == 1
    assert path.read_bytes() == before and not list(path.parent.glob("*.bak-*"))


@pytest.mark.parametrize("content", [
    '[mcp_servers.office-live]\ncommand = "old"\n',
    '[mcp_servers]\noffice-live = { command = "old" }\n',
    'mcp_servers.office-live.command = "old"\n',
    '# office-live was configured by hand\n',
])
def test_remove_codex_without_parser_refuses_any_mention(fake_home, monkeypatch, content):
    path = fake_home / ".codex" / "config.toml"
    path.parent.mkdir()
    path.write_text(content, encoding="utf-8")
    before = path.read_bytes()
    monkeypatch.setattr(cli, "_toml_module", lambda: None)
    ok, message = cli._remove_client("codex")
    assert ok is False and message.startswith("не удалено:") and "вручную" in message
    assert path.read_bytes() == before and not list(path.parent.glob("*.bak-*"))


@pytest.mark.parametrize("parser", [True, False])
def test_remove_codex_without_entry_is_successful(fake_home, monkeypatch, parser):
    path = fake_home / ".codex" / "config.toml"
    path.parent.mkdir()
    content = 'mcp_servers.other.command = "keep"\n'
    path.write_text(content, encoding="utf-8")
    if not parser:
        monkeypatch.setattr(cli, "_toml_module", lambda: None)
    assert cli._remove_client("codex") == (True, "записи не было")
    assert path.read_text(encoding="utf-8") == content and not list(path.parent.glob("*.bak-*"))


def test_remove_codex_parser_distinguishes_text_mention_from_entry(fake_home):
    path = fake_home / ".codex" / "config.toml"
    path.parent.mkdir()
    content = 'developer_instructions = """\n[mcp_servers.office-live]\ncommand = "example"\n"""\n'
    path.write_text(content, encoding="utf-8")
    assert cli._remove_client("codex") == (True, "записи не было")
    assert path.read_text(encoding="utf-8") == content


@pytest.mark.parametrize("content", ['mcp_servers.office-live.command = [', 'mcp_servers = 42'])
def test_remove_codex_invalid_config_reports_structured_failure(fake_home, content):
    path = fake_home / ".codex" / "config.toml"
    path.parent.mkdir()
    path.write_text(content, encoding="utf-8")
    ok, message = cli._remove_client("codex")
    assert ok is False and message.startswith("не удалено:") and "вручную" in message
    assert path.read_text(encoding="utf-8") == content


@pytest.mark.parametrize("cli_available", [False, True])
def test_remove_claude_failure_is_not_reported_as_absent(fake_home, monkeypatch, cli_available):
    (fake_home / ".claude.json").write_text(json.dumps({"mcpServers": {cli.SERVER_NAME: {"command": "old"}}}))
    monkeypatch.setattr(cli.shutil, "which", lambda name: "claude" if cli_available else None)
    monkeypatch.setattr(cli.subprocess, "call", lambda *a, **kw: 1)
    assert cli._remove_client("claude-code")[0] is False
    assert cli.setup_cmd(["--remove", "--yes", "--clients", "claude-code"]) == 1


@pytest.mark.parametrize("content", ['{bad json', '{"mcpServers": []}'])
def test_remove_malformed_json_keeps_config_and_returns_failure(fake_home, content):
    path = fake_home / ".cursor" / "mcp.json"
    path.parent.mkdir()
    path.write_text(content)
    assert cli.setup_cmd(["--remove", "--yes", "--clients", "cursor"]) == 1
    assert path.read_text() == content


def test_setup_with_no_clients_changes_nothing(fake_home):
    assert cli.run("setup", ["--yes", "--clients", "none"]) == 0
    assert not (fake_home / ".cursor").exists() and not (fake_home / ".codex").exists()


def test_setup_rejects_unknown_client_and_survives_broken_json(fake_home):
    assert cli.run("setup", ["--yes", "--clients", "banana"]) == 2
    broken = fake_home / ".cursor" / "mcp.json"
    broken.parent.mkdir(parents=True)
    broken.write_text("{not json", encoding="utf-8")
    assert cli.run("setup", ["--yes", "--clients", "cursor"]) == 1  # файл клиента не затираем молча
    assert broken.read_text(encoding="utf-8") == "{not json"


def test_detect_clients(fake_home):
    assert cli._detect_clients()["cursor"] is False
    (fake_home / ".cursor").mkdir()
    (fake_home / ".zcode" / "cli").mkdir(parents=True)
    d = cli._detect_clients()
    assert d["cursor"] and d["zcode"] and not d["codex"] and not d["vscode"]


def test_vscode_entry_has_stdio_type(fake_home, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    assert cli.run("setup", ["--yes", "--clients", "vscode"]) == 0
    entry = _servers(tmp_path / ".vscode" / "mcp.json", ["servers"])["office-live"]
    assert entry["type"] == "stdio"


def test_tools_json_survives_redirection_in_an_ansi_code_page(tmp_path):
    """CI: `tools --json > tools.json` падал с UnicodeEncodeError на кириллице в описаниях (вывод шёл в cp1252)."""
    import os
    import subprocess
    import sys

    env = {k: v for k, v in os.environ.items() if k not in ("PYTHONIOENCODING", "PYTHONUTF8")}
    env["PYTHONIOENCODING"] = "cp1252"  # как у перенаправленного вывода на английской Windows
    root = Path(__file__).resolve().parent.parent
    out = subprocess.run([sys.executable, "-m", "office_live", "tools", "--json"], cwd=root, env=env, capture_output=True, timeout=120)
    assert out.returncode == 0, out.stderr.decode("utf-8", "replace")[-500:]
    tools = json.loads(out.stdout.decode("utf-8"))
    assert any("срезы" in t["description"] for t in tools)
