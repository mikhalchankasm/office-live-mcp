"""Служебные команды: doctor (диагностика), tools (каталог инструментов), config (подключение к агентам)."""

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


def config_cmd(args) -> int:
    p = argparse.ArgumentParser(prog="office-live-mcp config", description="Print (or with --write, install) the MCP server entry for an agent.")
    p.add_argument("client", choices=["claude-code", "claude-desktop", "cursor", "vscode", "codex", "zcode", "generic"])
    p.add_argument("--write", action="store_true", help="modify the client's config file (a .bak copy is kept)")
    p.add_argument("--readonly", action="store_true", help="start the server in read-only mode")
    p.add_argument("--toolsets", default="", help="OFFICE_LIVE_TOOLSETS value, e.g. 'core' or 'excel_core,excel_format'")
    p.add_argument("--allowed-dirs", default="", help="OFFICE_LIVE_ALLOWED_DIRS value (';'-separated)")
    p.add_argument("--path", default="", help="override the config file location")
    p.add_argument("--scope", default="user", choices=["user", "local", "project"], help="claude-code only: user = all projects (default), local = this project only, project = shared .mcp.json")
    ns = p.parse_args(args)
    env = {}
    if ns.readonly:
        env["OFFICE_LIVE_MODE"] = "readonly"
    if ns.toolsets:
        env["OFFICE_LIVE_TOOLSETS"] = ns.toolsets
    if ns.allowed_dirs:
        env["OFFICE_LIVE_ALLOWED_DIRS"] = ns.allowed_dirs
    cmd, a = _server_command()
    home = Path.home()
    appdata = Path(os.environ.get("APPDATA", home / "AppData" / "Roaming"))

    if ns.client == "claude-code":
        parts = ["claude", "mcp", "add", "--scope", ns.scope, SERVER_NAME] + [x for k, v in env.items() for x in ("-e", f"{k}={v}")] + ["--", cmd] + a
        print("Run:\n  " + subprocess.list2cmdline(parts))
        if ns.write:
            if not shutil.which("claude"):
                print("The 'claude' CLI was not found on PATH.", file=sys.stderr)
                return 1
            return subprocess.call(parts)
        return 0

    if ns.client == "codex":
        block = f'[mcp_servers.{SERVER_NAME}]\ncommand = "{_toml_escape(cmd)}"\nargs = [{", ".join(chr(34) + _toml_escape(x) + chr(34) for x in a)}]\n'
        if env:
            block += f"[mcp_servers.{SERVER_NAME}.env]\n" + "".join(f'{k} = "{_toml_escape(v)}"\n' for k, v in env.items())
        print(block)
        if ns.write:
            path = Path(ns.path) if ns.path else home / ".codex" / "config.toml"
            text = path.read_text(encoding="utf-8") if path.exists() else ""
            if f"[mcp_servers.{SERVER_NAME}]" in text:
                print(f"{path} already has [mcp_servers.{SERVER_NAME}] - edit it by hand.", file=sys.stderr)
                return 1
            path.parent.mkdir(parents=True, exist_ok=True)
            if path.exists():
                shutil.copy2(path, path.with_suffix(".toml.bak"))
            path.write_text(text + ("\n" if text and not text.endswith("\n") else "") + "\n" + block, encoding="utf-8")
            print(f"Written to {path}")
        return 0

    targets = {
        "claude-desktop": (appdata / "Claude" / "claude_desktop_config.json", ["mcpServers"], False),
        "cursor": (home / ".cursor" / "mcp.json", ["mcpServers"], False),
        "vscode": (Path.cwd() / ".vscode" / "mcp.json", ["servers"], True),
        "zcode": (home / ".zcode" / "cli" / "config.json", ["mcp", "servers"], False),
        "generic": (None, ["mcpServers"], False),
    }
    path, keys, typed = targets[ns.client]
    entry = _json_server_entry(env, typed)
    snippet: dict = {}
    node = snippet
    for k in keys[:-1]:
        node = node.setdefault(k, {})
    node.setdefault(keys[-1], {})[SERVER_NAME] = entry
    print(json.dumps(snippet, ensure_ascii=False, indent=2))
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


def run(cmd: str, args: list[str]) -> int:
    if cmd == "doctor":
        return doctor()
    if cmd == "tools":
        return tools(args)
    if cmd == "config":
        return config_cmd(args)
    print("Usage: python -m office_live [serve | doctor | tools [--markdown|--json] | config <client> [--write]]", file=sys.stderr)
    return 2
