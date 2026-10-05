"""Per-user URI registration. All native registry access lives in Registry; tests replace it."""

import os
import sys
from contextlib import nullcontext
from datetime import datetime, timezone
from pathlib import Path

from .state import file_lock, read_state, state_path, write_state

KEY = r"Software\Classes\officelive"
PATH = "HKCU\\" + KEY
GUI_EXE = "office-live-link.exe"


class Registry:
    def read(self):
        import winreg

        def walk(key):
            values, children = {}, {}
            subkeys, count, _ = winreg.QueryInfoKey(key)
            for i in range(count):
                name, value, kind = winreg.EnumValue(key, i)
                values[name] = (value, kind)
            for i in range(subkeys):
                name = winreg.EnumKey(key, i)
                with winreg.OpenKey(key, name) as child:
                    children[name] = walk(child)
            return {"values": values, "children": children}

        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, KEY) as key:
                return walk(key)
        except FileNotFoundError:
            return None

    def write(self, command):
        import winreg

        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, KEY, 0, winreg.KEY_WRITE) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "URL:Office Live link")
            winreg.SetValueEx(key, "URL Protocol", 0, winreg.REG_SZ, "")
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, KEY + r"\shell\open\command", 0, winreg.KEY_WRITE) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)

    def delete(self):
        import winreg

        # Only the known four keys, never a recursive delete of arbitrary registry content.
        for suffix in (r"\shell\open\command", r"\shell\open", r"\shell", ""):
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, KEY + suffix)


def tree(command):
    def node(values=None, children=None):
        return {"values": {k: (v, 1) for k, v in (values or {}).items()}, "children": children or {}}

    return node({"": "URL:Office Live link", "URL Protocol": ""},
                {"shell": node(children={"open": node(children={"command": node({"": command})})})})


def command_of(snapshot):
    try:
        value, kind = snapshot["children"]["shell"]["children"]["open"]["children"]["command"]["values"][""]
        return value if kind == 1 else None
    except (KeyError, TypeError):
        return None


def identity(server=None):
    from .registrations import identity as server_identity

    executable, args, root = server_identity()
    if server is not None:
        executable, args = server
        root = str(Path(executable).parent.parent)
    if not args:
        launcher = Path(executable).with_name(GUI_EXE)
        command = f'"{launcher}" open-link "%1"'
    else:
        launcher = Path(executable).with_name("pythonw.exe")
        # -m must work from any browser working directory: installed package/venv required.
        command = f'"{launcher}" -m office_live open-link "%1"'
    return command, root, launcher


def records(state):
    return [r for r in state["registrations"] if r.get("kind") == "protocol"]


def configure(enabled=True, *, dry_run=False, server=None, allowed_dirs=None):
    from . import config

    command, root, launcher = identity(server)
    registry = Registry()
    with nullcontext() if dry_run else file_lock(state_path()):
        state = read_state()
        previous = records(state)
        snapshot = registry.read()
        actual = command_of(snapshot)
        if not enabled:
            if snapshot is not None and actual != command:
                print(f"Схема оставлена: {PATH} указывает на другую установку; проверьте вручную.")
                return False
            if snapshot is not None and snapshot != tree(command):
                print(f"Схема оставлена: {PATH} изменён пользователем; проверьте вручную.")
                return False
            print(f"{'ПЛАН: ' if dry_run else ''}Удалить свою схему {PATH}; обновить {state_path()}")
            if not dry_run:
                if snapshot is not None:
                    if registry.read() != snapshot:
                        raise OSError("Регистрация схемы изменилась параллельно; удаление отменено")
                    registry.delete()
                    if registry.read() is not None:
                        raise OSError("Схема осталась после удаления")
                remaining = [r for r in state["registrations"] if not (r.get("kind") == "protocol" and r["command"] == command)]
                if remaining != state["registrations"]:
                    state["registrations"] = remaining
                    write_state(state)
            return True
        # Updating/moving our registered copy is allowed; foreign or modified trees are preserved.
        owned = actual == command or any(r["command"] == actual for r in previous)
        if snapshot is not None and (not owned or snapshot != tree(actual)):
            raise ValueError(f"Конфликт {PATH}: чужая или изменённая регистрация оставлена")
        if not dry_run and not launcher.is_file():
            raise ValueError(f"Не найден обработчик ссылок: {launcher}")
        prior = next((r for r in previous if r["command"] == actual), None)
        dirs = allowed_dirs if allowed_dirs is not None else prior.get("allowed_dirs", []) if prior else [str(p) for p in config.SETTINGS.allowed_dirs]
        # Normalize with the same parser used by the server; no shell expansion during link dispatch.
        dirs = [str(p) for p in config.load({"OFFICE_LIVE_ALLOWED_DIRS": ";".join(dirs)}).allowed_dirs]
        print(f"{'ПЛАН: ' if dry_run else ''}{PATH} → {command}; разрешённые папки: {dirs or 'любые локальные'}; учесть в {state_path()}")
        if dry_run:
            return True
        record = {"kind": "protocol", "client": "protocol", "scope": "user", "path": PATH, "keys": [KEY], "name": "officelive",
                  "command": command, "root": root, "fingerprint": "", "time": datetime.now(timezone.utc).isoformat(),
                  "status": "pending", "allowed_dirs": dirs}
        state["registrations"] = [r for r in state["registrations"] if r.get("kind") != "protocol"] + [record]
        write_state(state)
        if registry.read() != snapshot:
            raise OSError("Регистрация схемы изменилась параллельно; запись отменена, повторите setup --links")
        registry.write(command)
        if registry.read() != tree(command):
            raise OSError("Контрольное чтение схемы не совпало; регистрация pending")
        record["status"] = "confirmed"
        write_state(state)
        return True


def allowed_settings():
    """Explorer does not inherit a client's env: also enforce the installation's saved policy."""
    from . import config

    command, _, _ = identity()
    # The GUI entry point must identify the same installation as the console server.
    record = next((r for r in records(read_state()) if r["command"] == command), None)
    return config.load({"OFFICE_LIVE_ALLOWED_DIRS": ";".join(record["allowed_dirs"])}) if record else config.SETTINGS


def error_dialog(message):
    if os.environ.get("OFFICE_LIVE_LINK_NO_DIALOG") == "1":
        return
    import ctypes

    ctypes.windll.user32.MessageBoxW(None, message, "Office Live", 0x10)


def open_link(args):
    # Parse before importing the COM tool layer. Invalid links finish quickly even without Office.
    from . import links

    try:
        if len(args) != 1:
            raise ValueError("Expected exactly one officelive:// link.")
        app, params = links.parse(args[0])
        from . import com, safety
        from .navigation import identity as file_identity, resolve

        policy = allowed_settings()

        def navigate():
            application, obj, sheet, rng = resolve(app, params)
            if not safety.doc_allowed(file_identity(app, obj) if obj.Path else "", policy):
                raise ValueError("The linked file is outside the link handler's allowed directories.")
            if app == "excel":
                from .excel_core import select_resolved_range
                from .xl_common import suspend_events

                suspend_events(application, force=True)
                if rng is not None:
                    select_resolved_range(application, obj, sheet, rng)
                else:
                    obj.Activate()
                    if sheet is not None:
                        sheet.Activate()
                hwnd = int(application.Hwnd)
            else:
                from .word_core import select_resolved_range

                if rng is not None:
                    select_resolved_range(application, obj, rng)
                else:
                    obj.Activate()
                hwnd = int(application.ActiveWindow.Hwnd)
            foreground(hwnd)

        # Convert errors ourselves to avoid run_com's developer traceback diagnostics.
        def guarded():
            from .errors import ToolError

            try:
                navigate()
            except Exception as exc:  # noqa: BLE001
                raise ToolError("Could not navigate to the open Office location: " + str(exc)[:240]) from None

        com.run_com(guarded, (), {}, tool="open-link", kind="ui")
        return 0
    except Exception as exc:  # noqa: BLE001 — shell entry point must show a short error, never a traceback
        message = "Office Live: " + (str(exc).splitlines() or ["Could not navigate to the linked location."])[0][:320]
        error_dialog(message)
        if sys.stderr is not None:
            print(message, file=sys.stderr)
        return 1


def foreground(hwnd):
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    user32.IsIconic.argtypes = [wintypes.HWND]
    user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.SetForegroundWindow.argtypes = [wintypes.HWND]
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, 9)  # SW_RESTORE
    user32.SetForegroundWindow(hwnd)  # Windows can deny focus stealing; Office activation still applies.
