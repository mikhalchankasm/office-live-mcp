"""Записи клиентов: адрес, принадлежность установке и проверяемая запись под блокировкой."""

import copy
import hashlib
import json
import os
import re
import sys
from contextlib import nullcontext
from datetime import date, datetime, timezone
from pathlib import Path

from . import cli
from .state import file_lock, read_state, state_path, write_state


def codex_home() -> Path:
    return Path(os.environ.get("CODEX_HOME") or Path.home() / ".codex")


def target(client, scope="user", project="", config="") -> dict:
    if config and (project or scope != "user"):
        raise ValueError("--config нельзя сочетать с --scope/--project")
    if project and scope not in {"project", "local"}:
        raise ValueError("--project требует --scope project")
    keys, typed = ["mcpServers"], False
    if scope == "project":
        names = {"claude-code": ".mcp.json", "cursor": ".cursor/mcp.json", "vscode": ".vscode/mcp.json",
                 "codex": ".codex/config.toml", "zcode": ".zcode/config.json"}
        if client not in names:
            raise ValueError(f"{client}: проектная конфигурация не поддерживается; используйте --config PATH")
        if not project:
            raise ValueError("Для --scope project укажите --project DIR")
        path = Path(project) / names[client]
    elif scope == "local":
        if client != "claude-code":
            raise ValueError("--scope local поддерживается только для старых конфигураций Claude Code")
        path = Path.home() / ".claude.json"
        keys = ["projects", str(Path(project or Path.cwd()).resolve()).replace("\\", "/"), "mcpServers"]
    elif client == "claude-code":
        path = Path.home() / ".claude.json"
    elif client == "codex":
        path = codex_home() / "config.toml"
    elif client == "generic":
        if not config:
            raise ValueError("Для generic укажите --path/--config PATH")
        path = Path(config)
    else:
        path, keys, typed = cli._json_target(client)
    if client == "vscode":
        keys, typed = ["servers"], True
        if not config:
            scope = "project"  # совместимость с прежним setup --clients vscode
    elif client == "codex":
        keys = ["mcp_servers"]
    elif client == "zcode":
        keys = ["mcp", "servers"]
    if config:
        path, scope = Path(config), "custom"
    return {"client": client, "scope": "project" if scope == "local" else scope, "path": str(path.absolute()),
            "keys": keys, "typed": typed, "name": cli.SERVER_NAME}


def fingerprint(entry) -> str:
    return hashlib.sha256(json.dumps(entry, sort_keys=True, ensure_ascii=False, separators=(",", ":"), default=str).encode("utf-8")).hexdigest()


def same_path(a, b) -> bool:
    if os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b)):
        return True
    # Короткие имена 8.3 (C:\Users\RUNNER~1\…): frozen sys.executable бывает в той форме, в которой exe запустили.
    try:
        return os.path.normcase(os.path.realpath(a)) == os.path.normcase(os.path.realpath(b))
    except (OSError, ValueError):
        return False


def identity() -> tuple[str, list[str], str]:
    command, args = cli._server_command()
    root = str(Path(command).parent.parent) if getattr(sys, "frozen", False) else str(Path(__file__).resolve().parent.parent)
    return command, args, root


def belongs(entry, command, args) -> bool:
    if not isinstance(entry, dict) or not isinstance(entry.get("command"), str) or not same_path(entry["command"], command):
        return False
    # Один Python может запускать разные серверы; одного совпадения python.exe недостаточно.
    return not args or entry.get("args", []) == args


def _bytes(path):
    return path.read_bytes() if path.exists() else None


def read_entry(spec, raw=None):
    path = Path(spec["path"])
    raw = _bytes(path) if raw is None else raw
    text = (raw or b"").decode("utf-8-sig")
    if spec["client"] == "codex":
        parser = cli._toml_module()
        if parser is None:
            raise ValueError("Нет парсера TOML: установите tomli для Python 3.10; исправьте конфигурацию вручную")
        data = parser.loads(text)
    else:
        data = json.loads(text) if text.strip() else {}
    node = data
    for key in spec["keys"]:
        if not isinstance(node, dict):
            raise ValueError(f"Конфигурация {path}: ожидался объект/таблица {key}; исправьте вручную")
        node = node.get(key, {})
    if not isinstance(node, dict):
        raise ValueError(f"Список серверов в {path} не является объектом/таблицей; исправьте вручную")
    entry = node.get(spec["name"])
    if entry is not None and (not isinstance(entry, dict) or not isinstance(entry.get("env", {}), dict)):
        raise ValueError(f"Неверная запись {spec['name']} в {path}; исправьте вручную")
    return text, data, entry


def _toml_value(value):
    if isinstance(value, str):
        return json.dumps(value, ensure_ascii=False)
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return repr(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, list):
        return "[" + ", ".join(_toml_value(v) for v in value) + "]"
    if isinstance(value, dict):
        return "{ " + ", ".join(f"{json.dumps(k)} = {_toml_value(v)}" for k, v in value.items()) + " }"
    raise ValueError(f"Неподдерживаемое значение TOML: {type(value).__name__}")


def _toml_key(key):
    return key if re.fullmatch(r"[A-Za-z0-9_-]+", key) else json.dumps(key, ensure_ascii=False)


def render(spec, text, data, entry):
    if spec["client"] == "codex":
        result = cli._strip_toml_tables(text)
        if entry is not None:
            result += f"\n[mcp_servers.{cli.SERVER_NAME}]\n"
            result += "".join(f"{_toml_key(k)} = {_toml_value(v)}\n" for k, v in entry.items() if k != "env")
            if entry.get("env"):
                result += f"[mcp_servers.{cli.SERVER_NAME}.env]\n"
                result += "".join(f"{_toml_key(k)} = {_toml_value(v)}\n" for k, v in entry["env"].items())
        problem = cli._toml_rewrite_problem(text, result, entry.get("env", {}) if entry is not None else None)
        if problem:
            raise ValueError(f"Отказ от записи {spec['path']}: {problem}; исправьте вручную")
        _, _, parsed = read_entry(spec, result.encode("utf-8"))
        if parsed != entry:
            raise ValueError("Результат TOML не совпадает с новой записью; исправьте вручную")
        return result
    data = copy.deepcopy(data)
    node = data
    for key in spec["keys"]:
        node = node.setdefault(key, {})
    if entry is None:
        node.pop(spec["name"], None)
    else:
        node[spec["name"]] = entry
    return json.dumps(data, ensure_ascii=False, indent=2)


def record_key(spec):
    if spec.get("kind") == "protocol":
        return "protocol", spec["path"], spec["name"]
    return os.path.normcase(os.path.abspath(spec["path"])), tuple(spec["keys"]), spec["name"]


def _upsert(state, record):
    state["registrations"] = [r for r in state["registrations"] if record_key(r) != record_key(record)] + [record]


def configure(spec, explicit, reset=False, full=False, force=False, dry_run=False, server=None):
    read_state()  # предварительная проверка нужна и для dry-run
    path = Path(spec["path"])
    command, args, root = identity()
    if server is not None:
        command, args = server
        root = str(Path(command).parent.parent)
    with nullcontext() if dry_run else file_lock(path):
        for _attempt in range(2):
            before = _bytes(path)
            text, data, previous = read_entry(spec, before)
            if previous is not None and not belongs(previous, command, args) and not force:
                raise ValueError(f"Конфликт {path}: запись {spec['name']} запускает {previous.get('command')!r}. Для замены нужен --force")
            entry = {} if reset else copy.deepcopy(previous or {})
            env = cli._merge_env(entry.get("env", {}), explicit, reset, full)
            entry.update(command=command, args=args)
            entry.pop("url", None)
            if spec["typed"] or "type" in entry:
                entry["type"] = "stdio"
            if env:
                entry["env"] = env
            else:
                entry.pop("env", None)
            result = render(spec, text, data, entry)
            if dry_run:
                print(f"ПЛАН: {'изменить' if previous else 'создать'} {path}: {spec['name']} → {command}; учесть в {state_path()}")
                return
            with file_lock(state_path()):
                state = read_state()
                if _bytes(path) != before:
                    continue
                if result.encode("utf-8") != before:
                    cli._backup(path)
                if _bytes(path) != before:
                    continue
                record = {**spec, "command": command, "args": args, "root": root, "fingerprint": fingerprint(entry),
                          "time": datetime.now(timezone.utc).isoformat(), "status": "pending"}
                # Намерение сохраняется первым: при прерывании запись клиента не потеряется из учёта.
                _upsert(state, record)
                write_state(state)
                if result.encode("utf-8") != before:
                    cli._atomic_write(path, result, expected=before)
                _, _, actual = read_entry(spec)
                if actual != entry:
                    raise ValueError(f"Контрольное чтение {path} не совпало; запись изменена другим процессом. Файл сохранён, регистрация pending.")
                record["status"] = "confirmed"
                write_state(state)
            print(f"Запись клиента записана и прочитана обратно: {path} ({spec['scope']}). Нужен перезапуск клиента.")
            return
        raise ValueError(f"Файл {path} изменён параллельно дважды; ничего не перезаписано. Повторите команду.")


def remove(spec, command=None, args=None, dry_run=False) -> tuple[bool, str]:
    path = Path(spec["path"])
    command, args = (command, args or []) if command else identity()[:2]
    with nullcontext() if dry_run else file_lock(path):
        before = _bytes(path)
        text, data, entry = read_entry(spec, before)
        if entry is not None:
            if not belongs(entry, command, args):
                return False, f"не удалено: чужая команда в {path}; запись {spec['name']} оставлена, проверьте вручную"
            if spec.get("fingerprint") and fingerprint(entry) != spec["fingerprint"]:
                return False, f"не удалено: пользователь изменил {path}, запись {spec['name']}; уберите её вручную"
        result = render(spec, text, data, None) if entry is not None else text
        if dry_run:
            return True, f"ПЛАН: {'удалить запись ' + spec['name'] if entry is not None else 'записи не было'} в {path}; обновить {state_path()}"
        with file_lock(state_path()):
            state = read_state()
            if entry is not None:
                cli._backup(path)
                cli._atomic_write(path, result, expected=before)
                if read_entry(spec)[2] is not None:
                    raise ValueError(f"Запись осталась в {path}; удалите вручную")
            remaining = [r for r in state["registrations"] if record_key(r) != record_key(spec)]
            if remaining != state["registrations"]:
                state["registrations"] = remaining
                write_state(state)
        return True, "удалено" if entry is not None else "записи не было"


def remove_client(client, scope=None, project="", config="", dry_run=False):
    command, args, root = identity()
    specs = [r for r in read_state()["registrations"] if r["client"] == client and same_path(r["root"], root)]
    if config or scope or project:
        requested = target(client, scope or "user", project, config)
        specs = [r for r in specs if record_key(r) == record_key(requested)] or [requested]
    if not specs:
        specs = [target(client)]
    results = []
    for spec in specs:
        try:
            ok, message = remove(spec, command, args, dry_run)
        except (ValueError, OSError) as exc:
            ok, message = False, f"не удалено: {spec['path']}: {exc}; исправьте вручную"
        results.append((ok, message))
    return all(ok for ok, _ in results), "; ".join(message for _, message in results)


def adopt_legacy(root: Path, problems: list[str] | None = None):
    """Учитываем существующие пользовательские записи старой установки, не переписывая конфиги."""
    command = str(root / "app" / "office-live-mcp.exe")
    for client in cli.CLIENT_LABELS:
        if client == "vscode":
            continue
        spec = target(client)
        path = Path(spec["path"])
        if not path.exists():
            continue
        try:
            with file_lock(path), file_lock(state_path()):
                state = read_state()
                if any(record_key(r) == record_key(spec) for r in state["registrations"]):
                    continue
                _, _, entry = read_entry(spec)
                if belongs(entry, command, []):
                    _upsert(state, {**spec, "command": command, "args": [], "root": str(root),
                                    "fingerprint": fingerprint(entry), "time": datetime.now(timezone.utc).isoformat(), "status": "confirmed"})
                    write_state(state)
                    print(f"Учтена прежняя регистрация: {path}; настройки не изменены.")
        except (OSError, ValueError) as exc:
            message = f"Не удалось учесть старую запись в {path}: {exc}. Проверьте и повторите setup --clients {client}."
            print(message)
            if problems is not None:
                problems.append(message)
