"""Служебные команды: setup (мастер подключения), doctor (диагностика), tools (каталог), config (запись для одного агента)."""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

SERVER_NAME = "office-live"


# ------------------------------------------------------------------ doctor


def _office_installed(progid: str) -> bool:
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + "\\CLSID"):
            return True
    except OSError:
        return False


def doctor() -> int:
    from . import __version__, config

    ok = True

    def line(flag: bool | None, text: str):
        mark = {True: "[ OK ]", False: "[FAIL]", None: "[ -- ]"}[flag]
        print(f"{mark} {text}")

    print(f"office-live-mcp {__version__}  |  Python {sys.version.split()[0]}  |  {sys.platform}")
    line(sys.platform == "win32", "Windows (COM automation is Windows-only)")
    ok &= sys.platform == "win32"
    for mod in ("mcp", "win32com", "pythoncom", "olefile"):
        try:
            m = __import__(mod)
            line(True, f"module {mod} {getattr(m, '__version__', '')}".rstrip())
        except ImportError:
            line(mod == "olefile" and None, f"module {mod} is missing" + (" (optional: needed to read macro source)" if mod == "olefile" else ""))
            ok &= mod == "olefile"
    if sys.platform == "win32":
        for label, progid in (("Excel", "Excel.Application"), ("Word", "Word.Application")):
            installed = _office_installed(progid)
            line(installed, f"{label} registered in Windows")
            ok &= installed
        try:
            from . import com

            for kind in ("excel", "word"):
                label = com.APPS[kind]["label"]
                try:
                    import pythoncom

                    pythoncom.CoInitialize()
                    try:
                        apps = com.apps(kind)
                        names = []
                        for a in apps:
                            coll = a.Workbooks if kind == "excel" else a.Documents
                            from .xl_common import is_undo_workbook

                            names += [coll(i).Name for i in range(1, int(coll.Count) + 1) if kind != "excel" or not is_undo_workbook(coll(i))]
                        line(True, f"{label} running: {len(apps)} instance(s), open files: {names[:8]}")
                    finally:
                        pythoncom.CoUninitialize()
                except Exception as exc:  # noqa: BLE001
                    line(None, f"{label} is not running ({str(exc)[:70]}) - tools will start it when asked to open/create something")
        except Exception as exc:  # noqa: BLE001
            line(False, f"COM check failed: {exc}")
            ok = False
    from .app import CATALOG

    registered = sum(1 for i in CATALOG.values() if i.registered)
    cfg = config.SETTINGS
    print(f"\nMode: {cfg.mode}  |  tools registered: {registered}/{len(CATALOG)}  |  toolsets: {sorted(cfg.groups)}")
    print(f"Allowed dirs: {[str(d) for d in cfg.allowed_dirs] or 'any'}  |  eval: {'ON' if cfg.allow_eval else 'off'}  |  audit log: {cfg.audit_log or 'off'}  |  strict target: {cfg.strict_target}")
    print("\nResult:", "everything needed is in place" if ok else "problems found (see [FAIL] above)")
    return 0 if ok else 1


# ------------------------------------------------------------------ tools


def tools(args) -> int:
    from .app import CATALOG

    p = argparse.ArgumentParser(prog="office-live-mcp tools")
    p.add_argument("--markdown", action="store_true", help="markdown table grouped by toolset")
    p.add_argument("--json", action="store_true")
    ns = p.parse_args(args)
    items = list(CATALOG.values())
    if ns.json:
        print(json.dumps([{"name": i.name, "group": i.group, "kind": i.kind, "read_actions": list(i.read_actions), "registered": i.registered, "description": i.doc.split("\n")[0]} for i in items], ensure_ascii=False, indent=1))
        return 0
    groups: dict[str, list] = {}
    for i in items:
        groups.setdefault(i.group, []).append(i)
    for g, lst in groups.items():
        if ns.markdown:
            print(f"\n### {g} ({len(lst)})\n\n| Tool | Kind | What it does |\n|---|---|---|")
            for i in lst:
                first = i.doc.split("\n")[0].replace("|", "\\|")
                kind = i.kind + (f" (readonly: {', '.join(i.read_actions)})" if i.read_actions else "")
                print(f"| `{i.name}` | {kind} | {first} |")
        else:
            print(f"\n[{g}] {len(lst)} tools")
            for i in lst:
                kind = i.kind + ("*" if i.read_actions else "")
                print(f"  {i.name:<34} {kind:<12} {'' if i.registered else '(disabled) '}{i.doc.split(chr(10))[0][:90]}")
    print(f"\nTotal: {len(items)} tools, {sum(1 for i in items if i.registered)} registered with the current settings.", file=sys.stderr)
    return 0


# ------------------------------------------------------------------ config


def _server_command() -> tuple[str, list[str]]:
    if getattr(sys, "frozen", False):  # собранный office-live-mcp.exe: без аргументов он и есть stdio-сервер
        return sys.executable, []
    py = sys.executable
    root = Path(__file__).resolve().parent.parent
    script = root / "server.py"
    if script.exists():
        return py, [str(script)]
    return py, ["-m", "office_live"]


def _self_cmd() -> str:
    """Как запускать служебные команды в подсказках: путь к exe у собранной программы, иначе python -m office_live."""
    return f'"{sys.executable}"' if getattr(sys, "frozen", False) else "python -m office_live"


def _json_server_entry(env: dict | None, extra_type: bool = False) -> dict:
    cmd, args = _server_command()
    entry = {"command": cmd, "args": args}
    if extra_type:
        entry["type"] = "stdio"
    if env:
        entry["env"] = env
    return entry


def _read_json(path: Path) -> dict:
    if path.exists() and path.read_text(encoding="utf-8").strip():
        return json.loads(path.read_text(encoding="utf-8"))
    return {}


def _backup(path: Path) -> Path | None:
    """Копия рядом с файлом: <имя>.bak-ГГГГММДД-ЧЧММСС[-N]. Прежняя копия не затирается, даже если две записи случились в одну секунду."""
    if not path.exists():
        return None
    stamp = time.strftime("%Y%m%d-%H%M%S")
    target = path.with_name(f"{path.name}.bak-{stamp}")
    n = 1
    while target.exists():
        n += 1
        target = path.with_name(f"{path.name}.bak-{stamp}-{n}")
    shutil.copy2(path, target)
    return target


def _atomic_write(path: Path, text: str) -> None:
    """Запись через временный файл и os.replace: при сбое настройки клиента не остаются обрезанными."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=str(path.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="") as f:
            f.write(text)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def _write_json_with_backup(path: Path, data: dict) -> None:
    _backup(path)
    _atomic_write(path, json.dumps(data, ensure_ascii=False, indent=2))


def _toml_escape(s: str) -> str:
    return s.replace("\\", "\\\\").replace('"', '\\"')


CLIENT_LABELS = {
    "claude-code": "Claude Code",
    "claude-desktop": "Claude Desktop",
    "cursor": "Cursor",
    "codex": "Codex CLI",
    "zcode": "ZCode",
    "vscode": "VS Code (в текущей папке)",
}
_JSON_TARGETS = {  # клиент -> (путь относительно home/appdata, ключи до словаря серверов, нужен ли "type": "stdio")
    "claude-desktop": ("appdata", ("Claude", "claude_desktop_config.json"), ["mcpServers"], False),
    "cursor": ("home", (".cursor", "mcp.json"), ["mcpServers"], False),
    "zcode": ("home", (".zcode", "cli", "config.json"), ["mcp", "servers"], False),
}


def _appdata() -> Path:
    return Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))


def _json_target(client: str) -> tuple[Path | None, list[str], bool]:
    if client == "vscode":
        return Path.cwd() / ".vscode" / "mcp.json", ["servers"], True
    base, parts, keys, typed = _JSON_TARGETS[client]
    root = _appdata() if base == "appdata" else Path.home()
    return root.joinpath(*parts), keys, typed


_TOML_HEADER = re.compile(r"^\s*\[\[?[^\[\]=\n]*\]\]?\s*(?:#.*)?$")
_NAME_RE = re.escape(SERVER_NAME)
_OURS = re.compile(
    rf"""^\s*\[\s*(?:"mcp_servers"|'mcp_servers'|mcp_servers)\s*\.\s*(?:"{_NAME_RE}"|'{_NAME_RE}'|{_NAME_RE})\s*(?:\.[^\]]*)?\]\s*(?:#.*)?$"""
)


def _toml_lines(text: str):
    """(строка, начинается ли она вне многострочной строки и вне незакрытого массива/inline-таблицы).

    Заголовком таблицы может быть только такая строка: '[mcp_servers.office-live]' внутри строки в тройных кавычках (пример в
    developer_instructions) или строка '[1, 2]' внутри многострочного массива — это значения, а не таблицы.
    """
    in_ml, depth = None, 0
    for line in text.splitlines(keepends=True):
        yield line, in_ml is None and depth == 0
        i, n = 0, len(line)
        while i < n:
            if in_ml:
                if in_ml == '"""' and line[i] == "\\":
                    i += 2
                elif line.startswith(in_ml, i):
                    in_ml, i = None, i + 3
                else:
                    i += 1
                continue
            c = line[i]
            if c == "#":
                break
            if line.startswith('"""', i) or line.startswith("'''", i):
                in_ml, i = line[i:i + 3], i + 3
            elif c in "\"'":
                j = i + 1
                while j < n and line[j] != c:
                    j += 2 if (c == '"' and line[j] == "\\") else 1
                i = j + 1
            else:
                if c in "[{":
                    depth += 1
                elif c in "]}":
                    depth = max(depth - 1, 0)
                i += 1


def _has_our_table(text: str) -> bool:
    return any(start and _OURS.match(line) for line, start in _toml_lines(text))


def _strip_toml_tables(text: str) -> str:
    """Удаляет таблицы [mcp_servers.office-live] и все вложенные ([...env], кавычки, пробелы, комментарий после заголовка)."""
    out, skipping = [], False
    for line, start in _toml_lines(text):
        if start and _TOML_HEADER.match(line):
            skipping = bool(_OURS.match(line))
        if not skipping:
            out.append(line)
    return "".join(out).rstrip("\n") + ("\n" if out else "")


class UnreadableConfig(ValueError):
    """Существующую запись office-live не удалось прочитать: молча затирать её настройки нельзя."""


def _toml_env(text: str) -> dict:
    """env уже установленной записи office-live в config.toml. Нет записи — {}; не удалось прочитать — UnreadableConfig."""
    if not _has_our_table(text) and "office-live" not in text:
        return {}
    toml = _toml_module()
    if toml is None:  # самодельный разбор TOML уже терял ограничения на допустимом синтаксисе — лучше честный отказ
        raise UnreadableConfig("this Python has no TOML parser (Python 3.10 needs the 'tomli' package: pip install tomli)")
    try:
        node = toml.loads(text).get("mcp_servers", {}).get(SERVER_NAME, {})
    except toml.TOMLDecodeError as exc:
        raise UnreadableConfig(f"invalid TOML: {exc}") from None
    if not isinstance(node, dict) or not isinstance(node.get("env") or {}, dict):
        raise UnreadableConfig(f"unexpected structure of the {SERVER_NAME} entry")
    return {str(k): str(v) for k, v in (node.get("env") or {}).items()}


def _toml_rewrite_problem(old_text: str, new_text: str, env: dict | None) -> str:
    """Проверка построчной замены: результат разбирается, всё вне записи office-live не изменилось, а сама запись —
    ровно новая с этим env (env=None — записи быть не должно, это удаление). Пустая строка — замена безопасна.

    Без парсера TOML проверить нечем (Python 3.10 без tomli) — тогда замена идёт без проверки, как и раньше.
    """
    toml = _toml_module()
    if toml is None:
        return ""
    try:
        new = toml.loads(new_text)
    except toml.TOMLDecodeError as exc:
        return f"the result would not be valid TOML ({exc})"
    try:
        old = toml.loads(old_text) if old_text else {}
    except toml.TOMLDecodeError as exc:
        return f"the existing file is not valid TOML ({exc}); fix it first"
    servers = new.get("mcp_servers")
    node = servers.get(SERVER_NAME) if isinstance(servers, dict) else None
    if env is None:
        if node is not None:
            return f"the {SERVER_NAME} entry is written in a form this installer cannot remove line by line"
    else:
        got = {str(k): str(v) for k, v in (node.get("env") or {}).items()} if isinstance(node, dict) and isinstance(node.get("env") or {}, dict) else None
        if got != {str(k): str(v) for k, v in env.items()}:
            return "the existing entry is written in a form this installer cannot replace line by line"
    if _without_our_entry(old) != _without_our_entry(new):
        return "the rewrite would also change other settings in the file"
    return ""


def _without_our_entry(doc: dict) -> dict:
    doc = dict(doc)
    servers = doc.get("mcp_servers")
    if isinstance(servers, dict):
        servers = {k: v for k, v in servers.items() if k != SERVER_NAME}
        if servers:
            doc["mcp_servers"] = servers
        else:
            doc.pop("mcp_servers")
    return doc


def _toml_module():
    """tomllib (Python 3.11+) или его бэкпорт tomli (зависимость пакета на 3.10); None — разбирать нечем."""
    try:
        import tomllib

        return tomllib
    except ImportError:
        pass
    try:
        import tomli

        return tomli
    except ImportError:
        return None


def _existing_env(client: str, path: Path | None, keys: list[str], scope: str = "user") -> dict:
    """Переменные окружения уже подключённой записи — чтобы повторная установка не снимала ограничения молча.

    Нет записи — {}. Запись есть, но прочитать её настройки нельзя — UnreadableConfig (а не пустой словарь).
    """
    try:
        if client == "codex":
            return _toml_env(path.read_text(encoding="utf-8")) if path and path.exists() else {}
        if client == "claude-code":
            entry = _claude_code_entry(scope)
            return dict((entry or {}).get("env") or {})
        if not path or not path.exists():
            return {}
        node = _read_json(path)
        for k in keys:
            node = node.get(k, {}) if isinstance(node, dict) else {}
        return {str(k): str(v) for k, v in ((node.get(SERVER_NAME) or {}).get("env") or {}).items()}
    except UnreadableConfig:
        raise
    except (OSError, ValueError, AttributeError) as exc:  # битый JSON, нечитаемый файл, неожиданная структура
        raise UnreadableConfig(f"{type(exc).__name__}: {exc}") from None


def _claude_code_entry(scope: str) -> dict | None:
    """Запись office-live из конфигурации Claude Code для области scope (читаем файлы, claude CLI не вызываем)."""
    try:
        if scope == "project":
            data = _read_json(Path.cwd() / ".mcp.json")
            return (data.get("mcpServers") or {}).get(SERVER_NAME)
        data = _read_json(Path.home() / ".claude.json")
        if scope == "user":
            return (data.get("mcpServers") or {}).get(SERVER_NAME)
        here = os.path.normcase(os.path.normpath(os.getcwd()))
        for key, value in (data.get("projects") or {}).items():
            if os.path.normcase(os.path.normpath(key)) == here:
                return ((value or {}).get("mcpServers") or {}).get(SERVER_NAME)
        return None
    except (OSError, ValueError, AttributeError) as exc:
        raise UnreadableConfig(f"{type(exc).__name__}: {exc}") from None


def _merge_env(previous: dict, explicit: dict, reset: bool, full: bool) -> dict:
    """Прежние настройки + явно заданные сейчас. reset — начать с нуля; full — явно вернуть режим full."""
    env = {} if reset else dict(previous)
    env.update(explicit)
    if full:
        env.pop("OFFICE_LIVE_MODE", None)
    return env


def config_cmd(args) -> int:
    p = argparse.ArgumentParser(prog="office-live-mcp config", description="Print (or with --write, install) the MCP server entry for an agent.")
    p.add_argument("client", choices=["claude-code", "claude-desktop", "cursor", "vscode", "codex", "zcode", "generic"])
    p.add_argument("--write", action="store_true", help="modify the client's config file (a dated .bak copy is kept); an existing office-live entry is replaced, its OFFICE_LIVE_* settings are kept")
    p.add_argument("--readonly", action="store_true", help="start the server in read-only mode")
    p.add_argument("--full", action="store_true", help="explicitly return to full mode (removes a previously set OFFICE_LIVE_MODE)")
    p.add_argument("--toolsets", default="", help="OFFICE_LIVE_TOOLSETS value, e.g. 'core' or 'excel_core,excel_format'")
    p.add_argument("--allowed-dirs", default="", help="OFFICE_LIVE_ALLOWED_DIRS value (';'-separated)")
    p.add_argument("--env", action="append", default=[], metavar="OFFICE_LIVE_X=value", help="any other OFFICE_LIVE_* setting (repeatable), e.g. OFFICE_LIVE_AUTOSAVE=allow")
    p.add_argument("--reset", action="store_true", help="forget the settings of the existing entry instead of keeping them")
    p.add_argument("--path", default="", help="override the config file location")
    p.add_argument("--scope", default="user", choices=["user", "local", "project"], help="claude-code only: user = all projects (default), local = this project only, project = shared .mcp.json")
    p.add_argument("--quiet", action="store_true", help="do not print the configuration snippet")
    ns = p.parse_args(args)
    if ns.readonly and ns.full:
        print("--readonly and --full contradict each other.", file=sys.stderr)
        return 2
    say = (lambda *a, **k: None) if ns.quiet else print
    explicit: dict = {}
    if ns.readonly:
        explicit["OFFICE_LIVE_MODE"] = "readonly"
    if ns.toolsets:
        explicit["OFFICE_LIVE_TOOLSETS"] = ns.toolsets
    if ns.allowed_dirs:
        explicit["OFFICE_LIVE_ALLOWED_DIRS"] = ns.allowed_dirs
    for item in ns.env:
        key, sep, value = item.partition("=")
        if not sep or not key.startswith("OFFICE_LIVE_"):
            print(f"--env expects OFFICE_LIVE_NAME=value, got {item!r}.", file=sys.stderr)
            return 2
        explicit[key] = value
    cmd, a = _server_command()

    if ns.client == "generic":
        path, keys, typed = None, ["mcpServers"], False
    elif ns.client in ("claude-code", "codex"):
        path, keys, typed = (Path(ns.path) if ns.path else Path.home() / ".codex" / "config.toml") if ns.client == "codex" else None, [], False
    else:
        path, keys, typed = _json_target(ns.client)
    if ns.path and ns.client != "claude-code":
        path = Path(ns.path)  # для generic путь задаётся только здесь
    try:
        previous = _existing_env(ns.client, path, keys, ns.scope) if ns.write else {}
    except UnreadableConfig as exc:
        if not ns.reset:
            print(
                f"Cannot read the settings of the existing {SERVER_NAME} entry ({exc}). Nothing was changed. "
                "Fix the file, or pass --reset to replace the entry and drop its settings.",
                file=sys.stderr,
            )
            return 1
        previous = {}
    env = _merge_env(previous, explicit, ns.reset, ns.full)
    kept = {k: v for k, v in env.items() if k in previous and k not in explicit}
    if kept and ns.write:
        print(f"Kept the existing settings of the entry: {kept} (use --reset to drop them).")

    if ns.client == "claude-code":
        parts = ["claude", "mcp", "add", "--scope", ns.scope, SERVER_NAME] + [x for k, v in env.items() for x in ("-e", f"{k}={v}")] + ["--", cmd] + a
        say("Run:\n  " + subprocess.list2cmdline(parts))
        if ns.write:
            if not shutil.which("claude"):
                print("The 'claude' CLI was not found on PATH.", file=sys.stderr)
                return 1
            try:
                old = _claude_code_entry(ns.scope)
            except UnreadableConfig:
                old = None
            subprocess.call(["claude", "mcp", "remove", SERVER_NAME, "-s", ns.scope], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)  # повторная установка заменяет запись
            rc = subprocess.call(parts, stdout=subprocess.DEVNULL if ns.quiet else None)
            if rc != 0 and old:
                restored = subprocess.call(["claude", "mcp", "add-json", "--scope", ns.scope, SERVER_NAME, json.dumps(old)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                print("The new entry could not be added; the previous one was " + ("restored." if restored == 0 else "NOT restored - re-add it manually."), file=sys.stderr)
            return rc
        return 0

    if ns.client == "codex":
        block = f'[mcp_servers.{SERVER_NAME}]\ncommand = "{_toml_escape(cmd)}"\nargs = [{", ".join(chr(34) + _toml_escape(x) + chr(34) for x in a)}]\n'
        if env:
            block += f"[mcp_servers.{SERVER_NAME}.env]\n" + "".join(f'{k} = "{_toml_escape(v)}"\n' for k, v in env.items())
        say(block)
        if ns.write:
            text = path.read_text(encoding="utf-8") if path.exists() else ""
            base = _strip_toml_tables(text) if text else ""
            new_text = base + ("\n" if base else "") + block
            problem = _toml_rewrite_problem(text, new_text, env)
            if problem:
                print(f"Refusing to rewrite {path}: {problem}. Nothing was changed; edit the {SERVER_NAME} entry by hand.", file=sys.stderr)
                return 1
            _backup(path)
            _atomic_write(path, new_text)
            print(f"Written to {path}")
        return 0

    entry = _json_server_entry(env, typed)
    snippet: dict = {}
    node = snippet
    for k in keys[:-1]:
        node = node.setdefault(k, {})
    node.setdefault(keys[-1], {})[SERVER_NAME] = entry
    say(json.dumps(snippet, ensure_ascii=False, indent=2))
    if ns.write:
        if ns.client == "generic" and not ns.path:
            print("Pass --path with --write for the generic client.", file=sys.stderr)
            return 1
        data = _read_json(path)
        node = data
        for k in keys[:-1]:
            node = node.setdefault(k, {})
        node.setdefault(keys[-1], {})[SERVER_NAME] = entry
        _write_json_with_backup(path, data)
        print(f"Written to {path}")
    return 0


# ------------------------------------------------------------------ setup (мастер подключения)


def _detect_clients() -> dict[str, bool]:
    return {
        "claude-code": shutil.which("claude") is not None,
        "claude-desktop": (_appdata() / "Claude").exists(),
        "cursor": (Path.home() / ".cursor").exists(),
        "codex": (Path.home() / ".codex").exists(),
        "zcode": (Path.home() / ".zcode" / "cli").exists(),
        "vscode": False,  # конфиг VS Code лежит в папке проекта — только по явному запросу
    }


def _remove_client(client: str) -> tuple[bool, str]:
    """Убирает запись office-live. Статус успеха не зависит от текста сообщения."""
    if client == "claude-code":
        path = Path.home() / ".claude.json"
        data = _read_json(path)
        if SERVER_NAME not in data.get("mcpServers", {}):
            return True, "записи не было"
        if not shutil.which("claude"):
            return False, "не удалено: нет CLI claude"
        rc = subprocess.call(["claude", "mcp", "remove", SERVER_NAME, "-s", "user"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return (True, "удалено") if rc == 0 else (False, f"не удалено: CLI claude завершился с кодом {rc}")
    if client == "codex":
        path = Path.home() / ".codex" / "config.toml"
        text = path.read_text(encoding="utf-8") if path.exists() else ""
        if not _has_our_table(text):
            return True, "записи не было"
        new_text = _strip_toml_tables(text)
        problem = _toml_rewrite_problem(text, new_text, None)
        if problem:
            return False, f"не удалено: {problem}; уберите запись {SERVER_NAME} вручную"
        _backup(path)
        _atomic_write(path, new_text)
        return True, "удалено"
    path, keys, _ = _json_target(client)
    if not path or not path.exists():
        return True, "записи не было"
    data = _read_json(path)
    node = data
    for k in keys:
        if not isinstance(node, dict):
            return False, "не удалено: конфигурация серверов не является объектом JSON"
        node = node.get(k)
        if node is None:
            return True, "записи не было"
    if not isinstance(node, dict):
        return False, "не удалено: список серверов не является объектом JSON"
    if SERVER_NAME not in node:
        return True, "записи не было"
    del node[SERVER_NAME]
    _write_json_with_backup(path, data)
    return True, "удалено"


def _ask(prompt: str, default: str = "") -> str:
    try:
        answer = input(prompt).strip()
    except EOFError:
        return default
    return answer or default


def _parse_selection(text: str, detected: dict[str, bool]) -> list[str]:
    order = [k for k in CLIENT_LABELS if k != "vscode"] + ["vscode"]
    text = text.strip().lower()
    if text in ("", "detected"):
        return [k for k in order if detected.get(k)]
    if text in ("0", "none", "no", "-"):
        return []
    if text == "all":
        return [k for k in order if k != "vscode"]
    chosen = []
    for token in text.replace(";", ",").replace(" ", ",").split(","):
        token = token.strip()
        if not token:
            continue
        if token.isdigit() and 1 <= int(token) <= len(order):
            chosen.append(order[int(token) - 1])
        elif token in CLIENT_LABELS:
            chosen.append(token)
        else:
            raise ValueError(f"не понял «{token}»")
    return list(dict.fromkeys(chosen))


def setup_cmd(args) -> int:
    p = argparse.ArgumentParser(prog="office-live-mcp setup", description="Подключить сервер к ИИ-агентам на этом компьютере (диалог).")
    p.add_argument("--clients", default=None, help="claude-code,claude-desktop,cursor,codex,zcode,vscode | detected | all | none")
    p.add_argument("--readonly", action="store_true", help="режим «только чтение»")
    p.add_argument("--full", action="store_true", help="явно вернуть полный режим (снять ранее заданный «только чтение»)")
    p.add_argument("--reset", action="store_true", help="не переносить настройки уже подключённой записи (по умолчанию они сохраняются)")
    p.add_argument("--toolsets", default="", help="OFFICE_LIVE_TOOLSETS, например core")
    p.add_argument("--allowed-dirs", default="", help="OFFICE_LIVE_ALLOWED_DIRS, каталоги через ';'")
    p.add_argument("--yes", action="store_true", help="без вопросов: берём --clients (по умолчанию — обнаруженные агенты)")
    p.add_argument("--remove", action="store_true", help="отключить сервер от агентов (удалить записи office-live)")
    ns = p.parse_args(args)
    interactive = not ns.yes and sys.stdin.isatty()
    detected = _detect_clients()
    order = [k for k in CLIENT_LABELS if k != "vscode"] + ["vscode"]

    print("\n=== Office Live MCP: подключение к агентам ===")
    if not ns.remove and interactive:
        print("\nКакой доступ дать агентам?")
        print("  1 - полный: читать и менять открытые Excel/Word (рекомендуется)")
        print("  2 - только чтение: агент смотрит документы, но ничего не меняет")
        choice = _ask("Выбор [1]: ", "1")
        if choice == "2":
            ns.readonly = True
        elif choice == "1":
            ns.full = True
    if ns.clients is None:
        if interactive:
            print("\nГде подключить?" + (" (✓ — найдено на этом компьютере)" if any(detected.values()) else ""))
            for i, key in enumerate(order, start=1):
                print(f"  {i} - {CLIENT_LABELS[key]}" + ("  ✓" if detected.get(key) else ""))
            print("  0 - никуда (только установка)")
            default = "detected" if any(detected.values()) else "0"
            while True:
                try:
                    selected = _parse_selection(_ask("Номера через запятую, Enter — все отмеченные ✓: ", default), detected)
                    break
                except ValueError as exc:
                    print(f"  {exc}, попробуйте ещё раз")
        else:
            selected = _parse_selection("detected", detected)
    else:
        try:
            selected = _parse_selection(ns.clients, detected)
        except ValueError as exc:
            print(f"Ошибка в --clients: {exc}", file=sys.stderr)
            return 2

    if not selected:
        print(f"\nНи к какому агенту не подключаю. Подключить позже: {_self_cmd()} setup")
        return 0
    failed = 0
    still_connected = []
    print()
    for client in selected:
        label = CLIENT_LABELS[client]
        if ns.remove:
            try:
                ok, message = _remove_client(client)
            except Exception as exc:  # noqa: BLE001 — ошибка одного клиента не мешает отключить остальных
                ok, message = False, f"не удалено: {exc}"
            print(f"  {label}: {message}")
            if not ok:
                failed += 1
                still_connected.append(label)
            continue
        extra = ["--quiet"]
        if ns.readonly:
            extra.append("--readonly")
        if ns.full:
            extra.append("--full")
        if ns.reset:
            extra.append("--reset")
        if ns.toolsets:
            extra += ["--toolsets", ns.toolsets]
        if ns.allowed_dirs:
            extra += ["--allowed-dirs", ns.allowed_dirs]
        try:
            rc = config_cmd([client, "--write"] + extra)
        except Exception as exc:  # noqa: BLE001 — неправильный JSON у клиента и т. п.
            print(f"  ✘ {label}: {exc}")
            failed += 1
            continue
        print(f"  {'✔' if rc == 0 else '✘'} {label}: " + ("подключено" if rc == 0 else "не удалось (см. сообщение выше)"))
        failed += rc != 0
    if ns.remove:
        if failed:
            print("\nНе подтверждено отключение; всё ещё могут быть подключены: " + ", ".join(still_connected))
        else:
            print("\nГотово. Перезапустите агентов.")
    elif failed:
        print(f"\nНе удалось подключить: {failed}. Проверка окружения: {_self_cmd()} doctor")
    else:
        mode = "только чтение" if ns.readonly else "полный доступ"
        print(f"\nГотово ({mode}). Перезапустите агента и попросите, например: «покажи, какие книги открыты в Excel».")
        print(f"Проверить окружение: {_self_cmd()} doctor   |   отключить: {_self_cmd()} setup --remove")
    return 1 if failed else 0


def _utf8_output() -> None:
    """Вывод в файл/канал по умолчанию идёт в кодировке ANSI (cp1252 в CI, cp1251 у пользователя) и падает на
    кириллице в описаниях инструментов; консоль Windows печатает Юникод и так. Поэтому вывод CLI — всегда UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError, OSError):
            pass


def run(cmd: str, args: list[str]) -> int:
    _utf8_output()
    if cmd == "doctor":
        return doctor()
    if cmd == "tools":
        return tools(args)
    if cmd == "config":
        return config_cmd(args)
    if cmd == "setup":
        return setup_cmd(args)
    if cmd in ("install", "uninstall"):
        from .install import install_cmd, uninstall_cmd

        return install_cmd(args) if cmd == "install" else uninstall_cmd(args)
    print(f"Usage: {_self_cmd()} [serve | install | uninstall | setup | doctor | tools [--markdown|--json] | config <client> [--write]]", file=sys.stderr)
    return 2
