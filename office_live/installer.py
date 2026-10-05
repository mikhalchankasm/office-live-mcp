"""JSON-контракт установщика Inno: учёт клиентов и настройка уже скопированной сборки."""

import argparse
import json
import subprocess
import sys
from contextlib import nullcontext, redirect_stdout
from pathlib import Path

from . import __version__, cli, install, registrations as reg
from .install_files import check_path, check_tree, write_manifest
from .probe import smoke
from .state import file_lock, read_state, remember_install, state_path

# Проектные/произвольные клиенты, не показанные мастером, никогда не снимаются его списком.
WIZARD_CLIENTS = tuple(c for c in cli.CLIENT_LABELS if c != "vscode")


def _records(root):
    return [r for r in read_state()["registrations"] if r.get("kind") != "protocol" and reg.same_path(r["root"], root)]


def clients_info() -> dict:
    command, args, root = reg.identity()
    records = _records(root)
    # Старый ZIP мог ещё не иметь ledger. Показываем свои пользовательские записи без записи state,
    # иначе мастер снимет их при postinstall/adopt_legacy как неотмеченные.
    known = {reg.record_key(r) for r in records}
    for client in WIZARD_CLIENTS:
        spec = reg.target(client)
        if reg.record_key(spec) not in known and Path(spec["path"]).is_file():
            entry = reg.read_entry(spec)[2]
            if reg.belongs(entry, command, args):
                records.append({**spec, "command": command, "args": args, "root": root})
    detected = cli._detect_clients()
    clients, modes = [], set()
    for client, label in cli.CLIENT_LABELS.items():
        own = [r for r in records if r["client"] == client]
        access = set()
        for record in own:
            entry = reg.read_entry(record)[2]
            if reg.belongs(entry, record["command"], record.get("args", [])):
                access.add("readonly" if entry.get("env", {}).get("OFFICE_LIVE_MODE") == "readonly" else "full")
        modes.update(access)
        clients.append({"id": client, "label": label, "detected": bool(detected.get(client)),
                        "registered": bool(own), "readonly": access == {"readonly"}})
    return {"clients": clients, "access": "mixed" if len(modes) > 1 else next(iter(modes), "full")}


def clients_cmd(args) -> int:
    parser = argparse.ArgumentParser(prog="office-live-mcp clients")
    parser.add_argument("--json", action="store_true")
    ns = parser.parse_args(args)
    try:
        data = clients_info()
    except (OSError, ValueError) as exc:
        if ns.json:
            print(json.dumps({"clients": [], "access": "full", "problems": [str(exc)]}, ensure_ascii=True))
        else:
            print(str(exc), file=sys.stderr)
        return 1
    if ns.json:
        print(json.dumps(data, ensure_ascii=True))
    else:
        for client in data["clients"]:
            print(f"{client['id']}: {client['label']}; detected={client['detected']}; registered={client['registered']}; readonly={client['readonly']}")
        print(f"Access: {data['access']}")
    return 0


def _selection(value):
    if value in {"keep", "detected", "none"}:
        return value
    selected = list(dict.fromkeys(c.strip() for c in value.split(",")))
    if not selected or any(c not in WIZARD_CLIENTS for c in selected):
        raise ValueError("--clients: " + ",".join(WIZARD_CLIENTS) + " | detected | none | keep")
    return selected


def postinstall_cmd(args) -> int:
    parser = argparse.ArgumentParser(prog="office-live-mcp postinstall")
    parser.add_argument("--clients", default="keep")
    parser.add_argument("--access", choices=["full", "readonly"], default="full")
    parser.add_argument("--no-links", action="store_true")
    parser.add_argument("--json", action="store_true")
    ns = parser.parse_args(args)
    result = {"ok": False, "smoke_tools": 0, "added": [], "removed": [], "kept": [], "problems": []}
    rc = 1
    # stdout остаётся ровно одним JSON; диагностический вывод помощников идёт в лог Inno через stderr.
    with redirect_stdout(sys.stderr) if ns.json else nullcontext():
        try:
            selected = _selection(ns.clients)
            exe = Path(sys.executable).absolute()
            if not install.frozen() or exe.name.lower() != install.EXE_NAME or exe.parent.name.lower() != "app":
                rc = 2
                raise ValueError("postinstall работает только из frozen office-live-mcp.exe в <корень>\\app")
            with file_lock(state_path().with_name("maintenance")):
                _postinstall(exe, ns, selected, result)
            rc = 1 if result["problems"] else 0
        except (OSError, ValueError, RuntimeError, subprocess.SubprocessError) as exc:
            result["problems"].append(str(exc))
    result["ok"] = rc == 0
    if ns.json:
        print(json.dumps(result, ensure_ascii=True))
    else:
        print(f"{'OK' if result['ok'] else 'FAIL'}: MCP tools={result['smoke_tools']}; "
              f"added={result['added']}; removed={result['removed']}; kept={result['kept']}")
        for problem in result["problems"]:
            print(problem)
    return rc


def _postinstall(exe, ns, selected, result):
    root = exe.parent.parent
    check_tree(root, exe.parent)
    check_path(root / install.MARKER, root)
    if (root / install.MARKER).exists() and not install._owned(root):
        raise ValueError(f"Недействительная метка {root / install.MARKER}")
    read_state()  # Повреждённый ledger не заменяется пустым.
    result["smoke_tools"] = smoke([str(exe)])
    install._write_marker(root)
    marker = json.loads((root / install.MARKER).read_text(encoding="utf-8"))
    marker["version"] = __version__
    cli._atomic_write(root / install.MARKER, json.dumps(marker))
    write_manifest(exe.parent)
    remember_install(root)
    reg.adopt_legacy(root, problems=result["problems"])
    install._cleanup_leftovers(root)
    if not ns.no_links:
        from .protocol import configure

        try:
            if not configure(server=(str(exe), [])):
                result["problems"].append("officelive://: регистрация не подтверждена; повторите setup --links")
        except (OSError, ValueError) as exc:
            result["problems"].append(f"officelive://: {exc}")
    records = _records(root)
    registered = {r["client"] for r in records}
    if selected == "keep":
        result["kept"] = sorted(registered)
        return
    if selected == "detected":
        # Обнаружение добавляет клиентов, но не отключает установленные в нестандартном месте.
        detected = cli._detect_clients()
        selected = [c for c in WIZARD_CLIENTS if detected.get(c) or c in registered]
    elif selected == "none":
        selected = []
    result["kept"] = sorted(c for c in registered if c in selected or c not in WIZARD_CLIENTS)
    for client in WIZARD_CLIENTS:
        own = [r for r in records if r["client"] == client]
        try:
            if client in selected and not own:
                reg.configure(reg.target(client), {"OFFICE_LIVE_MODE": "readonly"} if ns.access == "readonly" else {}, server=(str(exe), []))
                result["added"].append(client)
            elif client not in selected and own:
                failures = []
                for record in own:
                    try:
                        ok, message = reg.remove(record, str(exe), [])
                        if not ok:
                            failures.append(message)
                    except (OSError, ValueError) as exc:
                        failures.append(f"{record['path']}: {exc}")
                result["problems"].extend(f"{client}: {message}" for message in failures)
                if not failures:
                    result["removed"].append(client)
        except (OSError, ValueError) as exc:
            result["problems"].append(f"{client}: {exc}")
