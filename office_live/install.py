"""Установка и удаление автономной сборки office-live-mcp.exe (Python и зависимости внутри).

install: проверяет Windows и Office, копирует папку программы в %LOCALAPPDATA%\\Programs\\office-live-mcp\\app (права
администратора не нужны), проверяет установленную копию (doctor) и подключает её к агентам (setup). Путь в конфигах
агентов постоянный, поэтому обновление — тот же install из нового архива; настройки агентов сохраняются.
"""

import argparse
import base64
import json
import os
import platform
import re
import shutil
import stat
import subprocess
import sys
import time
from pathlib import Path

from . import __version__
from .cli import CLIENT_LABELS, _atomic_write, _office_installed, setup_cmd

EXE_NAME = "office-live-mcp.exe"
MARKER = "office-live-mcp.install"
PRODUCT = "office-live-mcp"


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


def _plain_dir(path: Path) -> bool:
    return path.is_dir() and not (path.lstat().st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT)


def _owned(root: Path) -> bool:
    try:
        return _plain_dir(root) and json.loads((root / MARKER).read_text(encoding="utf-8"))["product"] == PRODUCT
    except (OSError, ValueError, KeyError, TypeError):
        return False


def _write_marker(root: Path) -> None:
    # До копирования: даже прерванная первая установка имеет владельца. Чужой app проверен вызывающим кодом.
    if not (root / MARKER).exists():
        with (root / MARKER).open("x", encoding="utf-8") as stream:
            json.dump({"product": PRODUCT, "version": __version__, "created": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())}, stream)


def _replace_app(source: Path, app: Path) -> tuple[str, Path | None]:
    """Старую копию сохраняем до doctor; при неудаче переключения возвращаем её на место."""
    staging = app.with_name(f"app.new-{time.time_ns()}")
    staging.mkdir()  # никогда не удаляем и не перезаписываем прежнюю папку с совпавшим именем
    try:
        shutil.copytree(source, staging, dirs_exist_ok=True)
    except OSError:
        shutil.rmtree(staging)
        raise
    old = None
    if app.exists():
        old = app.with_name(f"app.old-{time.strftime('%Y%m%d%H%M%S')}")
        try:
            app.rename(old)  # папку с запущенным exe Windows переименовать не даст — значит, агент сейчас использует сервер
        except OSError:
            shutil.rmtree(staging, ignore_errors=True)
            return (
                "установленная копия сейчас используется: закройте ИИ-агентов, к которым подключён Office Live "
                "(Claude, Cursor, Codex…), и запустите установку ещё раз", None
            )
    try:
        staging.rename(app)
    except OSError as exc:
        if old is not None:
            old.rename(app)
        shutil.rmtree(staging, ignore_errors=True)
        return f"не удалось установить новую копию; прежняя копия сохранена: {exc}", None
    return "", old


def _starts(exe: Path) -> int:
    """0, если установленный exe запускается и отдаёт каталог инструментов (все модули и зависимости на месте)."""
    try:
        out = subprocess.run([str(exe), "tools", "--json"], capture_output=True, timeout=120)
    except (OSError, subprocess.TimeoutExpired) as exc:
        print(f"  FAIL запуск {exe.name}: {exc}")
        return 1
    try:
        if out.returncode == 0 and len(json.loads(out.stdout.decode("utf-8"))) > 0:
            return 0
    except ValueError:
        pass
    print(f"  FAIL {exe.name} tools --json: код {out.returncode}; {out.stderr.decode('utf-8', 'replace')[-300:]}")
    return 1


def _leftovers(root: Path) -> list[Path]:
    return [d for d in root.iterdir() if re.fullmatch(r"app\.(?:old-[0-9]{14}|new-[0-9]+)", d.name)
            and _plain_dir(d) and (d / EXE_NAME).is_file()]


def _cleanup_leftovers(root: Path) -> None:
    """Остатки прерванных обновлений (app.old-*, app.new-*), которые тогда были заняты."""
    if not _owned(root):
        return
    for d in _leftovers(root):
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
    old = None
    if ((root / MARKER).exists() and not _owned(root)) or (app.exists() and not _owned(root)):
        print(f"  FAIL нет действительной метки {MARKER}; чужая папка app не будет заменена")
        return 1
    if (root.exists() and not _plain_dir(root)) or (app.exists() and not _plain_dir(app)):
        print("  FAIL папка установки/app не должна быть ссылкой или junction")
        return 1
    root.mkdir(parents=True, exist_ok=True)
    _write_marker(root)
    if _same(source, app):
        print("  OK   запущено из установленной копии — копировать нечего")
    else:
        problem, old = _replace_app(source, app)
        if problem:
            print(f"  FAIL {problem}")
            return 1
        print("  OK   скопировано")
    exe = app / EXE_NAME

    _step(4, total, "Проверка установленной копии")
    # doctor — для информации: он «падает» и на машине без Word (или без Excel), где программа полностью рабочая, поэтому
    # судить об исправности копии по нему нельзя. Решает запуск установленного exe с загрузкой всех инструментов.
    try:
        subprocess.call([str(exe), "doctor"])
    except OSError as exc:
        print(f"  FAIL doctor: {exc}")
    rc = _starts(exe)
    if rc != 0:
        if old is not None:
            failed = root / f"app.new-{time.time_ns()}"
            app.rename(failed)
            try:
                old.rename(app)
            except OSError:
                failed.rename(app)
                raise
            shutil.rmtree(failed, ignore_errors=True)
            print("  FAIL проверка обновления не прошла; восстановлена прежняя копия. Подключение не изменено.")
            return 1
        print("  FAIL установленная программа не запускается. Файлы сохранены; агенты не подключены.")
        retry = subprocess.list2cmdline([sys.executable, "install", "--target", str(root), *setup_args])
        print(f"  Устраните ошибку запуска и повторите установку. Повторить: {retry}")
        return 1
    else:
        marker = json.loads((root / MARKER).read_text(encoding="utf-8"))
        marker["version"] = __version__
        _atomic_write(root / MARKER, json.dumps(marker))
        _cleanup_leftovers(root)

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
    root = Path(ns.target) if ns.target else install_root()
    root_exists = not ns.keep_files and os.path.lexists(root)  # учитываем и оборванные ссылки/junction
    if root_exists:
        if not _owned(root):
            print(f"Отказ от удаления: нет действительной метки {root / MARKER} или папка установки не является обычной. Ничего не изменено.")
            return 1
        if os.path.lexists(root / "app") and not _plain_dir(root / "app"):  # отсутствие app (прерванная установка) не мешает
            print("Отказ от удаления: app не является обычной папкой. Ничего не изменено.")
            return 1
    clients = ",".join(k for k in CLIENT_LABELS if k != "vscode")
    rc = setup_cmd(["--remove", "--yes", "--clients", clients])
    if rc != 0:
        print("Удаление остановлено: не удалось отключить перечисленных выше клиентов. Файлы программы сохранены.")
        return rc
    if ns.keep_files:
        return rc
    if not root_exists:
        print(f"Файлов не было: папка установки {root} отсутствует.")
        return rc
    if frozen() and _same(Path(sys.executable).parent, root / "app"):
        # свой exe удалить нельзя, пока он работает: удаляем папку отдельным процессом через пару секунд
        script = _self_delete_script(root.resolve())
        encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        subprocess.Popen(
            ["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded],
            creationflags=getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
            close_fds=True,
        )
        print(f"Файлы программы в {root} будут удалены через несколько секунд; посторонние файлы сохранятся.")
    else:
        try:
            if (root / "app").exists():
                shutil.rmtree(root / "app")
            for d in _leftovers(root):
                shutil.rmtree(d)
            (root / MARKER).unlink()
            if not any(root.iterdir()):
                root.rmdir()
        except OSError as exc:
            print(f"Не удалось удалить файлы программы: {exc}")
            return 1
        print(f"Файлы программы удалены: {root}")
    return rc


def _self_delete_script(root: Path) -> str:
    """Один shell, буквальные пути, повторная проверка владельца; root удаляется только пустым."""
    literal = str(root).replace("'", "''")
    return rf"""$ErrorActionPreference = 'Stop'
Start-Sleep -Seconds 2
$installRoot = '{literal}'
$rootItem = Get-Item -LiteralPath $installRoot
if ($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) {{ exit 1 }}
$marker = Join-Path $installRoot '{MARKER}'
if ((Get-Content -LiteralPath $marker -Raw | ConvertFrom-Json).product -ne '{PRODUCT}') {{ exit 1 }}
$targets = @(Get-ChildItem -LiteralPath $installRoot -Directory -Force | Where-Object {{
    -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -and
    ($_.Name -eq 'app' -or ($_.Name -cmatch '^app\.(old-[0-9]{{14}}|new-[0-9]+)$' -and
    (Test-Path -LiteralPath (Join-Path $_.FullName '{EXE_NAME}') -PathType Leaf)))
}})
foreach ($item in $targets) {{
    if ($item.Parent.FullName -ne $rootItem.FullName) {{ exit 1 }}
    Remove-Item -LiteralPath $item.FullName -Recurse -Force
}}
Remove-Item -LiteralPath $marker -Force
if (@(Get-ChildItem -LiteralPath $installRoot -Force).Count -eq 0) {{ [IO.Directory]::Delete($installRoot) }}
"""
