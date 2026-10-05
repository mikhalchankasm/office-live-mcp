"""Установка и удаление автономной сборки office-live-mcp.exe (Python и зависимости внутри).

install: проверяет Windows и Office, копирует папку программы в %LOCALAPPDATA%\\Programs\\office-live-mcp\\app (права
администратора не нужны), проверяет установленную копию по MCP и подключает её к агентам (setup). Путь в конфигах
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
from contextlib import nullcontext
from pathlib import Path

from . import __version__
from .cli import CLIENT_LABELS, _atomic_write, _office_installed, setup_cmd
from .state import file_lock, read_state, remember_install, state_path
from .install_files import check_path, check_tree, digest, remaining_paths, removal_plan, remove_files, write_manifest

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
        check_path(root / MARKER, root)
        return _plain_dir(root) and json.loads((root / MARKER).read_text(encoding="utf-8"))["product"] == PRODUCT
    except (OSError, ValueError, KeyError, TypeError):
        return False


def _write_marker(root: Path) -> None:
    # До копирования: даже прерванная первая установка имеет владельца. Чужой app проверен вызывающим кодом.
    if not (root / MARKER).exists():
        with (root / MARKER).open("x", encoding="utf-8") as stream:
            json.dump({"product": PRODUCT, "version": __version__, "created": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())}, stream)


def _exe_available(exe: Path) -> bool:
    """Windows иногда позволяет переименовать каталог с загруженным EXE: проверяем и сам файл."""
    if not exe.exists():
        return True
    import ctypes
    from ctypes import wintypes

    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    create = kernel.CreateFileW
    create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
                       wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    create.restype = wintypes.HANDLE
    handle = create(str(exe), 0xC0000000, 0, None, 3, 0x80, None)
    if handle == ctypes.c_void_p(-1).value:
        return False
    close = kernel.CloseHandle
    close.argtypes = [wintypes.HANDLE]
    close.restype = wintypes.BOOL
    close(handle)
    return True


def _discard_staging(staging: Path) -> None:
    check_tree(staging.parent, staging)
    shutil.rmtree(staging)


def _replace_app(source: Path, app: Path) -> tuple[str, Path | None]:
    """Старая копия хранится до проверки MCP; при неудаче переключения возвращаем её на место."""
    staging = app.with_name(f"app.new-{time.time_ns()}")
    staging.mkdir()  # никогда не удаляем и не перезаписываем прежнюю папку с совпавшим именем
    try:
        check_tree(source, source)
        shutil.copytree(source, staging, dirs_exist_ok=True)
        write_manifest(staging)
    except (OSError, ValueError):
        _discard_staging(staging)
        raise
    if _starts(staging / EXE_NAME):
        _discard_staging(staging)
        kept = "прежняя копия сохранена" if app.exists() else "исходный архив сохранён"
        return f"новая копия не прошла проверку; {kept}, переключение не выполнялось", None
    old = None
    if app.exists():
        old = app.with_name(f"app.old-{time.time_ns()}")
        try:
            if not _exe_available(app / EXE_NAME):
                raise PermissionError("EXE занят или нет доступа к файлу")
            app.rename(old)
        except OSError:
            _discard_staging(staging)
            return (
                "установленная копия занята или недоступна: закройте ИИ-агентов, к которым подключён Office Live "
                "(Claude, Cursor, Codex…), проверьте права доступа и запустите установку ещё раз", None
            )
    try:
        staging.rename(app)
    except OSError as exc:
        if old is not None:
            old.rename(app)
        _discard_staging(staging)
        return f"не удалось установить новую копию; прежняя копия сохранена: {exc}", None
    return "", old


def _starts(exe: Path) -> int:
    from .probe import smoke

    try:
        count = smoke([str(exe)], on_started=lambda: print(f"  OK   EXE запускается: {exe}"))
        print(f"  OK   MCP отвечает: initialize → initialized → tools/list ({count}); завершение по EOF проверено")
        return 0
    except (OSError, ValueError, RuntimeError, subprocess.SubprocessError) as exc:
        print(f"  FAIL проверка EXE/MCP {exe}: {exc}")
        return 1


def _leftovers(root: Path) -> list[Path]:
    return [d for d in root.iterdir() if re.fullmatch(r"app\.(?:old-(?:[0-9]{14}|[0-9]{19})|new-[0-9]+)", d.name)
            and _plain_dir(d) and ((d / EXE_NAME).is_file() or (d / "office-live-mcp.files.json").is_file())]


def _cleanup_leftovers(root: Path) -> None:
    """Остатки прерванных обновлений (app.old-*, app.new-*), которые тогда были заняты."""
    if not _owned(root):
        return
    for d in _leftovers(root):
        try:
            remaining = remove_files(root, d)
            if remaining:
                print(f"Сохранены посторонние/изменённые файлы: {d}; проверьте и удалите вручную.")
        except (OSError, ValueError) as exc:
            print(f"Остатки обновления сохранены в {d}: {exc}")


def install_cmd(args: list[str]) -> int:
    p = argparse.ArgumentParser(
        prog="office-live-mcp install",
        description="Установить Office Live MCP на этот компьютер и подключить к ИИ-агентам. Остальные параметры передаются в setup.",
    )
    p.add_argument("--target", default="", help="папка установки (по умолчанию %%LOCALAPPDATA%%\\Programs\\office-live-mcp)")
    p.add_argument("--new-location", action="store_true", help="разрешить другую папку при уже существующей установке (старую копию не удаляет)")
    p.add_argument("--dry-run", action="store_true", help="показать план без копирования, проверочных запусков и записи настроек")
    p.add_argument("--pause", action="store_true", help="в конце ждать Enter (запуск двойным щелчком)")
    ns, setup_args = p.parse_known_args(args)
    try:
        try:
            with nullcontext() if ns.dry_run else file_lock(state_path().with_name("maintenance")):
                return _install(ns, setup_args)
        except (OSError, ValueError) as exc:
            print(f"Ошибка установки: {exc}", file=sys.stderr)
            return 1
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
    state = read_state()
    recorded = Path(state["root"]) if state["root"] else None
    existing = recorded if recorded and _owned(recorded) else None
    # Старая установка без state: запуск её exe сохраняет нестандартный путь.
    if existing is None and _owned(source.parent) and source.name == "app":
        existing = source.parent
    root = (Path(ns.target) if ns.target else existing or install_root()).absolute()
    check_path(root, root)
    previous_location = recorded or existing
    if ns.target and previous_location and not _same(root, previous_location) and not ns.new_location:
        print(f"Уже есть установка в {previous_location}. Для другой папки нужен --new-location; прежняя копия останется.")
        return 1
    updating = _owned(root)
    app = root / "app"
    if not _same(source, app) and root.is_relative_to(source):
        raise ValueError("Каталог установки не должен находиться внутри копируемой папки архива")

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
    check_tree(root, app)
    if ns.dry_run:
        check_tree(source, source)
        for path in source.rglob("*"):
            if path.is_file():
                print(f"ПЛАН: скопировать {path} → {app / path.relative_to(source)}")
        print(f"ПЛАН: проверить EXE/MCP до и после переключения; записать {root / MARKER}, манифест и {state_path()}")
        if updating and not any(a == "--clients" or a.startswith("--clients=") for a in setup_args):
            print("ПЛАН: сохранить все регистрации и их настройки; новые клиенты не добавляются.")
            return 0
        return setup_cmd([*setup_args, "--dry-run"], server=(str(app / EXE_NAME), []))
    root.mkdir(parents=True, exist_ok=True)
    _write_marker(root)
    if _same(source, app):
        print("  OK   запущено из установленной копии — копировать нечего")
    else:
        problem, old = _replace_app(source, app)
        if problem:
            print(f"  FAIL {problem}")
            print("Повторить: " + subprocess.list2cmdline([sys.executable, "install", "--target", str(root), *setup_args]))
            return 1
        print("  OK   скопировано")
    exe = app / EXE_NAME

    _step(4, total, "Проверка установленной копии")
    # Проверка запуска не обращается к открытым документам Office.
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
            remove_files(root, failed)
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
        remember_install(root)
        if updating:
            from .registrations import adopt_legacy

            adopt_legacy(root)
        _cleanup_leftovers(root)

    _step(5, total, "Подключение ИИ-агентов")
    if updating and not any(a == "--clients" or a.startswith("--clients=") for a in setup_args):
        print("Существующие регистрации и их настройки сохранены. Новых клиентов можно добавить через setup --clients …")
        rc = 0
    else:
        rc = subprocess.call([str(exe), "setup", *setup_args])
    print(f"\nПрограмма установлена в: {app}. Для загрузки сервера нужен перезапуск клиента.")
    target = f' --target "{root}"' if ns.target else ""
    print(f'Обновить: запустить install.cmd из нового архива.  Удалить: "{exe}" uninstall{target}')
    return rc


def uninstall_cmd(args: list[str]) -> int:
    p = argparse.ArgumentParser(prog="office-live-mcp uninstall", description="Отключить записи этой установки и удалить собственные файлы.")
    p.add_argument("--keep-files", action="store_true", help="только отключить клиентов, файлы программы оставить")
    p.add_argument("--target", default="", help="корень установки; по умолчанию — state.json или каталог своего exe")
    p.add_argument("--dry-run", action="store_true", help="показать план без изменений")
    ns = p.parse_args(args)
    try:
        with nullcontext() if ns.dry_run else file_lock(state_path().with_name("maintenance")):
            return _uninstall(ns)
    except (OSError, ValueError) as exc:
        print(f"Удаление остановлено: {exc}. Файлы программы сохранены.")
        return 1


def _uninstall(ns) -> int:
    from . import registrations as reg

    state = read_state()
    own_root = Path(sys.executable).parent.parent if frozen() else None
    root = Path(ns.target) if ns.target else own_root if own_root and _owned(own_root) else Path(state["root"]) if state["root"] else install_root()
    root = root.absolute()
    exists = os.path.lexists(root)
    if exists and not ns.keep_files:
        check_path(root, root)
        check_path(root / MARKER, root)
        if not _owned(root):
            raise ValueError(f"Нет действительной метки {root / MARKER}; ничего не изменено")
        check_tree(root, root / "app")
        if (root / "app").exists() and not (root / "app").is_dir():
            raise ValueError("app не является обычной папкой")
        for folder in _leftovers(root):
            check_tree(root, folder)

    command = str(root / "app" / EXE_NAME)
    specs = [r for r in state["registrations"] if reg.same_path(r["root"], root)]
    # Старые установки не имели state: читаем только известные пользовательские адреса.
    known = {reg.record_key(r) for r in specs}
    for client in CLIENT_LABELS:
        if client == "vscode":
            continue
        spec = reg.target(client)
        if reg.record_key(spec) not in known and Path(spec["path"]).exists():
            specs.append(spec)
    blocked, failures = [], []
    for spec in specs:
        try:
            _, _, current = reg.read_entry(spec)
            dependent = current is not None and reg.belongs(current, command, [])
            if current is not None and not dependent:
                print(f"Оставлена чужая запись {spec['name']}: {spec['path']}; команда {current.get('command')!r}")
                continue
            ok, message = reg.remove(spec, command, [], ns.dry_run)
            print(f"{spec['client']} [{spec['scope']}] {spec['path']}: {message}")
            if not ok:
                failures.append(spec["path"])
                if dependent:
                    blocked.append(spec["path"])
        except (ValueError, OSError) as exc:
            failures.append(spec["path"])
            blocked.append(spec["path"])  # нечитаемый файл: зависимость нельзя безопасно исключить
            print(f"Осталась регистрация {spec['path']}: {exc}; уберите office-live вручную и повторите uninstall.")
    if blocked or ns.keep_files:
        print(f"Файлы сохранены: {root}. " + ("Не удалось отключить зависимые регистрации: " + ", ".join(blocked) if blocked else "Указан --keep-files."))
        return 1 if failures else 0
    if not exists:
        print(f"Файлов не было: {root}")
        return 1 if failures else 0
    folders = [root / "app", *_leftovers(root)]
    for folder in folders:
        own, keep = removal_plan(root, folder)
        for path in own:
            print(f"{'ПЛАН: ' if ns.dry_run else ''}Удаление файла: {path}")
        for path in keep:
            print(f"Остаётся посторонний/изменённый файл: {path}; при необходимости удалите вручную.")
    if ns.dry_run:
        print(f"ПЛАН: удалить метку {root / MARKER}, пустые каталоги; состояние и резервные копии конфигураций сохраняются.")
        return 1 if failures else 0
    if frozen() and _same(Path(sys.executable).parent, root / "app"):
        report = state_path().with_name("uninstall-report.json")
        script = _self_delete_script(root, report)
        helper = state_path().with_name(f"uninstall-{time.time_ns()}.ps1")
        _atomic_write(helper, "\ufeff" + script)
        _atomic_write(report, json.dumps({"status": "pending", "root": str(root)}))
        subprocess.Popen(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(helper)],
                         creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), close_fds=True, cwd=str(root.parent),
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print(f"Удаление файлов назначено после выхода EXE. Результат и точный список остатков: {report}. При ошибке закройте клиентов и повторите uninstall из архива с --target \"{root}\".")
    else:
        errors = []
        planned = {p for folder in folders for p in removal_plan(root, folder)[0]}
        for folder in folders:
            try:
                remove_files(root, folder)
            except (OSError, ValueError) as exc:
                errors.append(str(exc))
        if not errors:
            (root / MARKER).unlink(missing_ok=True)
        remaining = remaining_paths(root)
        for path in sorted(planned - set(remaining)):
            print(f"Удалён файл: {path}")
        if not remaining:
            root.rmdir()
        print(("Удаление неполное. " if errors else "Удаление завершено. ") +
              ("Осталось: " + ", ".join(str(p) for p in remaining) if remaining else "Остатков файлов нет."))
        if errors:
            print("Ошибки: " + "; ".join(errors) + ". Закройте клиентов, проверьте права и повторите uninstall; оставшиеся пути можно проверить вручную.")
            return 1
    return 1 if failures else 0


def _self_delete_script(root: Path, report: Path | None = None) -> str:
    """Удаляем только заранее проверенные файлы; фоновый процесс повторяет проверки границ и ссылок."""
    report = report or state_path().with_name("uninstall-report.json")
    files = []
    for folder in [root / "app", *_leftovers(root)]:
        own, _ = removal_plan(root, folder)
        files.extend({"path": str(p.absolute()), "hash": digest(p)} for p in own)
    marker = root / MARKER
    files.append({"path": str(marker), "hash": digest(marker)})
    payload = base64.b64encode(json.dumps(files).encode("utf-8")).decode("ascii")
    literal = str(root).replace("'", "''")
    report_literal = str(report).replace("'", "''")
    return rf"""$ErrorActionPreference = 'Stop'
Start-Sleep -Seconds 2
$installRoot = '{literal}'
$reportPath = '{report_literal}'
$removed = @()
$errors = @()
function Check-Path($path) {{
    $full = [IO.Path]::GetFullPath($path)
    if ($full -ne $installRoot -and -not $full.StartsWith($installRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {{ throw "Outside root: $full" }}
    $part = $full
    while ($part) {{
        if (Test-Path -LiteralPath $part) {{
            if ((Get-Item -Force -LiteralPath $part).Attributes -band [IO.FileAttributes]::ReparsePoint) {{ throw "Reparse point: $part" }}
        }}
        if ($part -eq $installRoot) {{ break }}  # предков выше корня не проверяем
        $part = [IO.Path]::GetDirectoryName($part)
    }}
}}
try {{
    Check-Path $installRoot
    $marker = Join-Path $installRoot '{MARKER}'
    if ((Get-Content -LiteralPath $marker -Raw | ConvertFrom-Json).product -ne '{PRODUCT}') {{ throw 'Invalid marker' }}
    $files = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{payload}')) | ConvertFrom-Json
    foreach ($file in $files) {{
        try {{
            if ($errors.Count -and ($file.path -eq $marker -or [IO.Path]::GetFileName($file.path) -eq 'office-live-mcp.files.json')) {{ continue }}
            Check-Path $file.path
            if (Test-Path -LiteralPath $file.path) {{
                $stream = [IO.File]::OpenRead($file.path)
                $hasher = [Security.Cryptography.SHA256]::Create()
                try {{ $hash = [BitConverter]::ToString($hasher.ComputeHash($stream)).Replace('-', '') }}
                finally {{ $stream.Dispose(); $hasher.Dispose() }}
                if ($hash -ne $file.hash) {{ throw "Changed file: $($file.path)" }}
                Remove-Item -LiteralPath $file.path -Force
                $removed += $file.path
            }}
        }} catch {{ $errors += $_.Exception.Message }}
    }}
    # Без рекурсивного удаления; ссылки и чужие каталоги не обходятся.
    $dirs = @($files | ForEach-Object {{ [IO.Path]::GetDirectoryName($_.path) }} | Select-Object -Unique)
    foreach ($start in $dirs) {{
        $dir = $start
        while ($dir -ne $installRoot -and $dir.StartsWith($installRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {{
            Check-Path $dir
            if ((Test-Path -LiteralPath $dir) -and @(Get-ChildItem -LiteralPath $dir -Force).Count -eq 0) {{ [IO.Directory]::Delete($dir) }}
            $dir = [IO.Path]::GetDirectoryName($dir)
        }}
    }}
    if (@(Get-ChildItem -LiteralPath $installRoot -Force).Count -eq 0) {{ [IO.Directory]::Delete($installRoot) }}
}} catch {{ $errors += $_.Exception.Message }}
$remaining = @()
if (Test-Path -LiteralPath $installRoot) {{
    $pending = @($installRoot)
    while ($pending.Count -gt 0) {{
        $current = $pending[0]
        $pending = @($pending | Select-Object -Skip 1)
        foreach ($item in @(Get-ChildItem -LiteralPath $current -Force)) {{
            if ($item.PSIsContainer -and -not ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) {{ $pending += $item.FullName }}
            else {{ $remaining += $item.FullName }}
        }}
    }}
}}
[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($reportPath)) | Out-Null
@{{ status = 'complete'; removed = $removed; remaining = $remaining; errors = $errors }} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $reportPath -Encoding UTF8
if ($errors.Count) {{ exit 1 }}
"""
