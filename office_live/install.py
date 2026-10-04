"""Установка и удаление автономной сборки office-live-mcp.exe (Python и зависимости внутри).

install: проверяет Windows и Office, копирует папку программы в %LOCALAPPDATA%\\Programs\\office-live-mcp\\app (права
администратора не нужны), проверяет установленную копию (doctor) и подключает её к агентам (setup). Путь в конфигах
агентов постоянный, поэтому обновление — тот же install из нового архива; настройки агентов сохраняются.
"""

import argparse
import os
import platform
import shutil
import subprocess
import sys
import time
from pathlib import Path

from . import __version__
from .cli import CLIENT_LABELS, _office_installed, setup_cmd

EXE_NAME = "office-live-mcp.exe"


def frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def install_root() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / "Programs" / "office-live-mcp"


def _step(n: int, total: int, text: str) -> None:
    print(f"\n[{n}/{total}] {text}", flush=True)  # иначе при выводе в канал шаги перемешаются с выводом дочерних exe


def _same(a: Path, b: Path) -> bool:
    try:
        return os.path.normcase(str(a.resolve())) == os.path.normcase(str(b.resolve()))
    except OSError:
        return False


def _check_windows() -> tuple[bool, str]:
    if sys.platform != "win32":
        return False, "программа работает только в Windows"
    build = sys.getwindowsversion().build
    if build < 10240:
        return False, f"нужна Windows 10 или 11 (здесь сборка {build})"
    return True, f"Windows {platform.release()} (сборка {build}), {platform.machine()}"


def _replace_app(source: Path, app: Path) -> str:
    """Копирует программу в app. Пустая строка — успех, иначе причина отказа (ничего не испорчено)."""
    staging = app.with_name(f"app.new-{os.getpid()}")
    shutil.rmtree(staging, ignore_errors=True)
    shutil.copytree(source, staging)
    old = None
    if app.exists():
        old = app.with_name(f"app.old-{time.strftime('%Y%m%d%H%M%S')}")
        try:
            app.rename(old)  # папку с запущенным exe Windows переименовать не даст — значит, агент сейчас использует сервер
        except OSError:
            shutil.rmtree(staging, ignore_errors=True)
            return (
                "установленная копия сейчас используется: закройте ИИ-агентов, к которым подключён Office Live "
                "(Claude, Cursor, Codex…), и запустите установку ещё раз"
            )
    staging.rename(app)
    if old is not None:
        shutil.rmtree(old, ignore_errors=True)
    return ""


def _cleanup_leftovers(root: Path) -> None:
    """Остатки прерванных обновлений (app.old-*, app.new-*), которые тогда были заняты."""
    for d in root.glob("app.*-*"):
        shutil.rmtree(d, ignore_errors=True)


def install_cmd(args: list[str]) -> int:
    p = argparse.ArgumentParser(
        prog="office-live-mcp install",
        description="Установить Office Live MCP на этот компьютер и подключить к ИИ-агентам. Остальные параметры передаются в setup.",
    )
    p.add_argument("--target", default="", help="папка установки (по умолчанию %%LOCALAPPDATA%%\\Programs\\office-live-mcp)")
    p.add_argument("--pause", action="store_true", help="в конце ждать Enter (запуск двойным щелчком)")
    ns, setup_args = p.parse_known_args(args)
    try:
        return _install(ns, setup_args)
    finally:
        if ns.pause:
            try:
                input("\nНажмите Enter, чтобы закрыть окно...")
            except EOFError:
                pass


def _install(ns, setup_args: list[str]) -> int:
    print(f"=== Office Live MCP {__version__}: установка ===")
    if not frozen():
        print("Эта команда устанавливает собранную программу (office-live-mcp.exe). Из папки с исходниками используйте "
              "install.cmd или `python -m office_live setup`.", file=sys.stderr)
        return 2
    total = 5
    source = Path(sys.executable).parent
    root = Path(ns.target) if ns.target else install_root()
    app = root / "app"

    _step(1, total, "Windows")
    ok, text = _check_windows()
    print(("  OK   " if ok else "  FAIL ") + text)
    if not ok:
        return 1

    _step(2, total, "Microsoft Excel и Word")
    found = {label: _office_installed(progid) for label, progid in (("Excel", "Excel.Application"), ("Word", "Word.Application"))}
    for label, present in found.items():
        print(f"  {'OK   ' if present else 'НЕТ  '}{label} " + ("установлен" if present else "не установлен — его инструменты работать не будут"))
    if not any(found.values()):
        print("  FAIL нужен настольный Excel или Word (Microsoft 365 / Office 2016+). Установите Office и запустите установку ещё раз.")
        return 1

    _step(3, total, f"Копирование программы в {app}")
    if _same(source, app):
        print("  OK   запущено из установленной копии — копировать нечего")
    else:
        root.mkdir(parents=True, exist_ok=True)
        problem = _replace_app(source, app)
        if problem:
            print(f"  FAIL {problem}")
            return 1
        _cleanup_leftovers(root)
        print("  OK   скопировано")
    exe = app / EXE_NAME

    _step(4, total, "Проверка установленной копии")
    rc = subprocess.call([str(exe), "doctor"])
    if rc != 0:
        print("  (часть проверок не прошла — см. [FAIL] выше; подключение продолжается)")

    _step(5, total, "Подключение ИИ-агентов")
    rc = subprocess.call([str(exe), "setup", *setup_args])
    print(f"\nУстановлено в: {app}")
    target = f' --target "{root}"' if ns.target else ""
    print(f'Обновить: запустить install.cmd из нового архива.  Удалить: "{exe}" uninstall{target}')
    return rc


def uninstall_cmd(args: list[str]) -> int:
    p = argparse.ArgumentParser(prog="office-live-mcp uninstall", description="Отключить Office Live MCP от всех агентов и удалить установленную программу.")
    p.add_argument("--keep-files", action="store_true", help="только отключить агентов, папку программы оставить")
    p.add_argument("--target", default="", help="папка установки (по умолчанию %%LOCALAPPDATA%%\\Programs\\office-live-mcp)")
    ns = p.parse_args(args)
    clients = ",".join(k for k in CLIENT_LABELS if k != "vscode")
    rc = setup_cmd(["--remove", "--yes", "--clients", clients])
    root = Path(ns.target) if ns.target else install_root()
    if ns.keep_files or not root.exists():
        return rc
    if frozen() and _same(Path(sys.executable).parent, root / "app"):
        # свой exe удалить нельзя, пока он работает: удаляем папку отдельным процессом через пару секунд
        subprocess.Popen(
            f'cmd /c ping -n 3 127.0.0.1 >nul & rd /s /q "{root}"',
            creationflags=getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
            close_fds=True,
        )
        print(f"Папка {root} будет удалена через несколько секунд.")
    else:
        shutil.rmtree(root, ignore_errors=True)
        print(f"Удалено: {root}" if not root.exists() else f"Не удалось полностью удалить {root}; удалите папку вручную.")
    return rc
