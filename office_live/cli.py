"""Служебные команды: setup (мастер подключения), doctor (диагностика), tools (каталог), config (запись для одного агента)."""

import argparse
import json
import os
import shutil
import subprocess
import sys
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
                            names += [coll(i).Name for i in range(1, int(coll.Count) + 1)]
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
        print(json.dumps([{"name": i.name, "group": i.group, "kind": i.kind, "registered": i.registered, "description": i.doc.split("\n")[0]} for i in items], ensure_ascii=False, indent=1))
        return 0
    groups: dict[str, list] = {}
    for i in items:
        groups.setdefault(i.group, []).append(i)
    for g, lst in groups.items():
        if ns.markdown:
            print(f"\n### {g} ({len(lst)})\n\n| Tool | Kind | What it does |\n|---|---|---|")
            for i in lst:
                first = i.doc.split("\n")[0].replace("|", "\\|")
                print(f"| `{i.name}` | {i.kind} | {first} |")
        else:
            print(f"\n[{g}] {len(lst)} tools")
            for i in lst:
                print(f"  {i.name:<34} {i.kind:<11} {'' if i.registered else '(disabled) '}{i.doc.split(chr(10))[0][:90]}")
    print(f"\nTotal: {len(items)} tools, {sum(1 for i in items if i.registered)} registered with the current settings.", file=sys.stderr)
    return 0


# ------------------------------------------------------------------ config


def _server_command() -> tuple[str, list[str]]:
    py = sys.executable
    root = Path(__file__).resolve().parent.parent
    script = root / "server.py"
    if script.exists():
        return py, [str(script)]
    return py, ["-m", "office_live"]


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


def _write_json_with_backup(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        shutil.copy2(path, path.with_suffix(path.suffix + ".bak"))
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


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


def _strip_toml_tables(text: str) -> str:
    """Удаляет таблицы [mcp_servers.office-live] и [mcp_servers.office-live.env] из TOML (построчно: массивы в значениях не мешают)."""
    out, skipping = [], False
    for line in text.splitlines(keepends=True):
        stripped = line.strip()
        if stripped.startswith("["):
            skipping = stripped in (f"[mcp_servers.{SERVER_NAME}]", f"[mcp_servers.{SERVER_NAME}.env]")
        if not skipping:
            out.append(line)
    return "".join(out).rstrip("\n") + ("\n" if out else "")


def config_cmd(args) -> int:
    p = argparse.ArgumentParser(prog="office-live-mcp config", description="Print (or with --write, install) the MCP server entry for an agent.")
    p.add_argument("client", choices=["claude-code", "claude-desktop", "cursor", "vscode", "codex", "zcode", "generic"])
    p.add_argument("--write", action="store_true", help="modify the client's config file (a .bak copy is kept); an existing office-live entry is replaced")
    p.add_argument("--readonly", action="store_true", help="start the server in read-only mode")
    p.add_argument("--toolsets", default="", help="OFFICE_LIVE_TOOLSETS value, e.g. 'core' or 'excel_core,excel_format'")
    p.add_argument("--allowed-dirs", default="", help="OFFICE_LIVE_ALLOWED_DIRS value (';'-separated)")
    p.add_argument("--path", default="", help="override the config file location")
    p.add_argument("--scope", default="user", choices=["user", "local", "project"], help="claude-code only: user = all projects (default), local = this project only, project = shared .mcp.json")
    p.add_argument("--quiet", action="store_true", help="do not print the configuration snippet")
    ns = p.parse_args(args)
    say = (lambda *a, **k: None) if ns.quiet else print
    env = {}
    if ns.readonly:
        env["OFFICE_LIVE_MODE"] = "readonly"
    if ns.toolsets:
        env["OFFICE_LIVE_TOOLSETS"] = ns.toolsets
    if ns.allowed_dirs:
        env["OFFICE_LIVE_ALLOWED_DIRS"] = ns.allowed_dirs
    cmd, a = _server_command()

    if ns.client == "claude-code":
        parts = ["claude", "mcp", "add", "--scope", ns.scope, SERVER_NAME] + [x for k, v in env.items() for x in ("-e", f"{k}={v}")] + ["--", cmd] + a
        say("Run:\n  " + subprocess.list2cmdline(parts))
        if ns.write:
            if not shutil.which("claude"):
                print("The 'claude' CLI was not found on PATH.", file=sys.stderr)
                return 1
            subprocess.call(["claude", "mcp", "remove", SERVER_NAME, "-s", ns.scope], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)  # повторная установка заменяет запись
            return subprocess.call(parts, stdout=subprocess.DEVNULL if ns.quiet else None)
        return 0

    if ns.client == "codex":
        block = f'[mcp_servers.{SERVER_NAME}]\ncommand = "{_toml_escape(cmd)}"\nargs = [{", ".join(chr(34) + _toml_escape(x) + chr(34) for x in a)}]\n'
        if env:
            block += f"[mcp_servers.{SERVER_NAME}.env]\n" + "".join(f'{k} = "{_toml_escape(v)}"\n' for k, v in env.items())
        say(block)
        if ns.write:
            path = Path(ns.path) if ns.path else Path.home() / ".codex" / "config.toml"
            text = path.read_text(encoding="utf-8") if path.exists() else ""
            path.parent.mkdir(parents=True, exist_ok=True)
            if path.exists():
                shutil.copy2(path, path.with_suffix(".toml.bak"))
            base = _strip_toml_tables(text) if text else ""
            path.write_text(base + ("\n" if base else "") + block, encoding="utf-8")
            print(f"Written to {path}")
        return 0

    if ns.client == "generic":
        path, keys, typed = None, ["mcpServers"], False
    else:
        path, keys, typed = _json_target(ns.client)
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
        path = Path(ns.path) if ns.path else path
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


def _remove_client(client: str) -> str:
    """Убирает запись office-live из конфигурации клиента. Возвращает короткий статус."""
    if client == "claude-code":
        if not shutil.which("claude"):
            return "нет CLI claude"
        rc = subprocess.call(["claude", "mcp", "remove", SERVER_NAME, "-s", "user"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return "удалено" if rc == 0 else "записи не было"
    if client == "codex":
        path = Path.home() / ".codex" / "config.toml"
        if not path.exists() or f"[mcp_servers.{SERVER_NAME}]" not in path.read_text(encoding="utf-8"):
            return "записи не было"
        shutil.copy2(path, path.with_suffix(".toml.bak"))
        path.write_text(_strip_toml_tables(path.read_text(encoding="utf-8")), encoding="utf-8")
        return "удалено"
    path, keys, _ = _json_target(client)
    if not path or not path.exists():
        return "записи не было"
    data = _read_json(path)
    node = data
    for k in keys:
        node = node.get(k) if isinstance(node, dict) else None
        if node is None:
            return "записи не было"
    if SERVER_NAME not in node:
        return "записи не было"
    del node[SERVER_NAME]
    _write_json_with_backup(path, data)
    return "удалено"


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
        if _ask("Выбор [1]: ", "1") == "2":
            ns.readonly = True
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
        print("\nНи к какому агенту не подключаю. Подключить позже: python -m office_live setup")
        return 0
    failed = 0
    print()
    for client in selected:
        label = CLIENT_LABELS[client]
        if ns.remove:
            print(f"  {label}: {_remove_client(client)}")
            continue
        extra = ["--quiet"]
        if ns.readonly:
            extra.append("--readonly")
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
        print("\nГотово. Перезапустите агентов.")
    elif failed:
        print(f"\nНе удалось подключить: {failed}. Проверка окружения: python -m office_live doctor")
    else:
        mode = "только чтение" if ns.readonly else "полный доступ"
        print(f"\nГотово ({mode}). Перезапустите агента и попросите, например: «покажи, какие книги открыты в Excel».")
        print("Проверить окружение: python -m office_live doctor   |   отключить: python -m office_live setup --remove")
    return 1 if failed else 0


def run(cmd: str, args: list[str]) -> int:
    if cmd == "doctor":
        return doctor()
    if cmd == "tools":
        return tools(args)
    if cmd == "config":
        return config_cmd(args)
    if cmd == "setup":
        return setup_cmd(args)
    print("Usage: python -m office_live [serve | setup | doctor | tools [--markdown|--json] | config <client> [--write]]", file=sys.stderr)
    return 2
