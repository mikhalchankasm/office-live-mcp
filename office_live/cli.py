"""Служебные команды: setup (мастер подключения), doctor (диагностика), tools (каталог), config (запись для одного агента)."""

import argparse
import json
import os
import re
import shutil
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


_UNCHECKED = object()


def _atomic_write(path: Path, text: str, expected=_UNCHECKED) -> None:
    """Запись через временный файл и os.replace: при сбое настройки клиента не остаются обрезанными."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=str(path.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="") as f:
            f.write(text)
        if expected is not _UNCHECKED and (path.read_bytes() if path.exists() else None) != expected:
            raise OSError(f"Файл {path} изменён параллельно; запись отменена, повторите команду")
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


def _strip_toml_tables(text: str) -> str:
    """Удаляет таблицы [mcp_servers.office-live] и все вложенные ([...env], кавычки, пробелы, комментарий после заголовка)."""
    out, skipping = [], False
    for line, start in _toml_lines(text):
        if start and _TOML_HEADER.match(line):
            skipping = bool(_OURS.match(line))
        if not skipping:
            out.append(line)
    return "".join(out).rstrip("\n") + ("\n" if out else "")


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


def _merge_env(previous: dict, explicit: dict, reset: bool, full: bool) -> dict:
    """Прежние настройки + явно заданные сейчас. reset — начать с нуля; full — явно вернуть режим full."""
    env = {} if reset else dict(previous)
    env.update(explicit)
    if full:
        env.pop("OFFICE_LIVE_MODE", None)
    return env


def config_cmd(args, server=None) -> int:
    from . import registrations as reg

    p = argparse.ArgumentParser(prog="office-live-mcp config", description="Показать или записать конфигурацию одного клиента.")
    p.add_argument("client", choices=[*CLIENT_LABELS, "generic"])
    p.add_argument("--write", action="store_true", help="записать конфигурацию с резервной копией и контрольным чтением")
    p.add_argument("--readonly", action="store_true", help="только чтение")
    p.add_argument("--full", action="store_true", help="снять ограничение только чтения")
    p.add_argument("--toolsets", default="", help="наборы OFFICE_LIVE_TOOLSETS")
    p.add_argument("--allowed-dirs", default="", help="разрешённые каталоги через ';'")
    p.add_argument("--env", action="append", default=[], help="OFFICE_LIVE_NAME=value, можно повторять")
    p.add_argument("--reset", action="store_true", help="явно сбросить параметры своей записи")
    p.add_argument("--path", "--config", default="", help="нестандартный файл конфигурации")
    p.add_argument("--scope", default=None, choices=["user", "local", "project"], help="область; local — совместимость Claude Code")
    p.add_argument("--project", default="", help="каталог проекта для --scope project")
    p.add_argument("--force", action="store_true", help="разрешить замену чужой записи office-live")
    p.add_argument("--dry-run", action="store_true", help="показать план, ничего не записывать")
    p.add_argument("--quiet", action="store_true", help="не выводить фрагмент конфигурации")
    ns = p.parse_args(args)
    if ns.readonly and ns.full:
        print("--readonly и --full противоречат друг другу.", file=sys.stderr)
        return 2
    explicit = {}
    for key, value in (("OFFICE_LIVE_MODE", "readonly" if ns.readonly else ""),
                       ("OFFICE_LIVE_TOOLSETS", ns.toolsets), ("OFFICE_LIVE_ALLOWED_DIRS", ns.allowed_dirs)):
        if value:
            explicit[key] = value
    for item in ns.env:
        key, sep, value = item.partition("=")
        if not sep or not key.startswith("OFFICE_LIVE_"):
            print(f"--env ожидает OFFICE_LIVE_NAME=value: {item!r}", file=sys.stderr)
            return 2
        explicit[key] = value
    try:
        if ns.client == "vscode" and ns.scope == "user" and not ns.path:
            raise ValueError("VS Code: пользовательский профиль выбирается через --config PATH; либо используйте --scope project --project DIR")
        # generic без записи служит только примером JSON.
        spec = reg.target(ns.client, ns.scope or "user", ns.project or (str(Path.cwd()) if ns.scope in {"project", "local"} else ""),
                          ns.path or ("mcp.json" if ns.client == "generic" and not ns.write else ""))
        if ns.write or ns.dry_run:
            reg.configure(spec, explicit, ns.reset, ns.full, ns.force, ns.dry_run, server)
        elif not ns.quiet:
            entry = _json_server_entry(explicit, spec["typed"])
            print(reg.render(spec, "", {}, entry))
        return 0
    except (OSError, ValueError) as exc:
        print(f"Не удалось записать конфигурацию: {exc}", file=sys.stderr)
        return 1


# ------------------------------------------------------------------ setup (мастер подключения)


def _detect_clients() -> dict[str, bool]:
    return {
        "claude-code": shutil.which("claude") is not None,
        "claude-desktop": (_appdata() / "Claude").exists(),
        "cursor": (Path.home() / ".cursor").exists(),
        "codex": Path(os.environ.get("CODEX_HOME") or Path.home() / ".codex").exists(),
        "zcode": (Path.home() / ".zcode" / "cli").exists(),
        "vscode": False,  # конфиг VS Code лежит в папке проекта — только по явному запросу
    }


def _remove_client(client: str, **kwargs) -> tuple[bool, str]:
    from .registrations import remove_client

    return remove_client(client, **kwargs)


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


def setup_cmd(args, server=None) -> int:
    p = argparse.ArgumentParser(prog="office-live-mcp setup", description="Подключить сервер к ИИ-агентам на этом компьютере (диалог).")
    p.add_argument("--clients", default=None, help="claude-code,claude-desktop,cursor,codex,zcode,vscode | detected | all | none")
    p.add_argument("--readonly", action="store_true", help="режим «только чтение»")
    p.add_argument("--full", action="store_true", help="явно вернуть полный режим (снять ранее заданный «только чтение»)")
    p.add_argument("--reset", action="store_true", help="не переносить настройки уже подключённой записи (по умолчанию они сохраняются)")
    p.add_argument("--toolsets", default="", help="OFFICE_LIVE_TOOLSETS, например core")
    p.add_argument("--allowed-dirs", default="", help="OFFICE_LIVE_ALLOWED_DIRS, каталоги через ';'")
    p.add_argument("--yes", action="store_true", help="без вопросов: берём --clients (по умолчанию — обнаруженные агенты)")
    p.add_argument("--remove", action="store_true", help="отключить сервер от агентов (удалить записи office-live)")
    p.add_argument("--config", default="", help="нестандартный файл; нужен ровно один клиент")
    p.add_argument("--scope", default=None, choices=["user", "project"], help="область регистрации (по умолчанию user; VS Code — текущий проект)")
    p.add_argument("--project", default="", help="каталог проекта для --scope project")
    p.add_argument("--force", action="store_true", help="разрешить замену чужой записи office-live")
    p.add_argument("--dry-run", action="store_true", help="показать план без записи файлов и состояния")
    p.add_argument("--env", action="append", default=[], help="OFFICE_LIVE_NAME=value, можно повторять")
    ns = p.parse_args(args)
    if ns.readonly and ns.full:
        print("--readonly и --full противоречат друг другу.", file=sys.stderr)
        return 2
    if ns.scope == "project" and not ns.project:
        print("Для --scope project укажите --project DIR.", file=sys.stderr)
        return 2
    interactive = not ns.yes and not ns.dry_run and sys.stdin.isatty()
    detected = _detect_clients()
    order = [k for k in CLIENT_LABELS if k != "vscode"] + ["vscode"]

    print("\n=== Office Live MCP: подключение к агентам ===")
    if not ns.remove and interactive and not (ns.readonly or ns.full):
        print("\nКакой доступ дать агентам?")
        print("  1 - полный: читать и менять открытые Excel/Word (рекомендуется)")
        print("  2 - только чтение: агент смотрит документы, но ничего не меняет")
        print("  Enter — сохранить прежние настройки; для новых записей полный доступ")
        choice = _ask("Выбор [сохранить]: ", "")
        if choice == "2":
            ns.readonly = True
        elif choice == "1":
            ns.full = True
    if ns.clients is None:
        if ns.remove:
            from . import registrations as reg
            from .state import read_state

            try:
                root = reg.identity()[2]
                selected = list(dict.fromkeys(r["client"] for r in read_state()["registrations"]
                                              if reg.same_path(r["root"], root)))
                selected = list(dict.fromkeys([*selected, *_parse_selection("detected", detected)]))
            except (OSError, ValueError) as exc:
                print(str(exc), file=sys.stderr)
                return 1
        elif interactive:
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

    if ns.config and len(selected) != 1:
        print("--config требует ровно одного клиента в --clients.", file=sys.stderr)
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
                ok, message = _remove_client(client, scope=ns.scope, project=ns.project, config=ns.config, dry_run=ns.dry_run)
            except Exception as exc:  # noqa: BLE001 — ошибка одного клиента не мешает отключить остальных
                ok, message = False, f"не удалено: {exc}"
            print(f"  {label}: {message}")
            if not ok:
                failed += 1
                still_connected.append(label)
            continue
        extra = ["--quiet"]
        for flag, value in (("--scope", ns.scope), ("--project", ns.project), ("--config", ns.config)):
            if value:
                extra += [flag, value]
        if ns.force:
            extra.append("--force")
        if ns.dry_run:
            extra.append("--dry-run")
        for item in ns.env:
            extra += ["--env", item]
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
            rc = config_cmd([client, "--write"] + extra, server=server)
        except Exception as exc:  # noqa: BLE001 — неправильный JSON у клиента и т. п.
            print(f"  ✘ {label}: {exc}")
            failed += 1
            continue
        print(f"  {'✔' if rc == 0 else '✘'} {label}: " + ("план проверен" if rc == 0 and ns.dry_run else "конфигурация проверена" if rc == 0 else "не удалось (см. сообщение выше)"))
        failed += rc != 0
    if ns.remove:
        if failed:
            print("\nНе подтверждено отключение; всё ещё могут быть подключены: " + ", ".join(still_connected))
        else:
            print("\nГотово. Перезапустите агентов.")
    elif failed:
        print(f"\nНе удалось подключить: {failed}. Проверка окружения: {_self_cmd()} doctor")
    else:
        mode = "только чтение" if ns.readonly else "полный доступ" if ns.full else "настройки сохранены; новые записи — полный доступ"
        print(f"\nГотово ({mode}). Перезапустите агента и попросите, например: «покажи, какие книги открыты в Excel».")
        print(f"Проверить окружение: {_self_cmd()} doctor   |   отключить: {_self_cmd()} setup --remove")
    return 1 if failed else 0


def _utf8_output() -> None:
    """Вывод в файл/канал по умолчанию идёт в кодировке ANSI (cp1252 в CI, cp1251 у пользователя) и падает на
    кириллице в описаниях инструментов; консоль Windows печатает Юникод и так. Поэтому вывод CLI — всегда UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="backslashreplace" if stream is sys.stderr else "strict")
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
    print(f"Использование: {_self_cmd()} [serve | install | uninstall | setup | doctor | tools [--markdown|--json] | config <client> [--write]]", file=sys.stderr)
    return 2
