"""Принадлежность файлов установки и границы удаления. Пользовательские файлы остаются."""

import hashlib
import json
import os
import stat
from pathlib import Path

MANIFEST = "office-live-mcp.files.json"


def digest(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest() if hasattr(hashlib, "file_digest") else _digest(stream)


def _digest(stream):
    hasher = hashlib.sha256()
    for chunk in iter(lambda: stream.read(1024 * 1024), b""):
        hasher.update(chunk)
    return hasher.hexdigest()


def check_path(path: Path, root: Path) -> None:
    path, root = path.absolute(), root.absolute()
    if not path.is_relative_to(root):
        raise ValueError(f"Путь вне каталога установки: {path}")
    # Предков выше корня не проверяем: перенаправленный AppData или профиль — выбор пользователя, а не обход границы.
    for part in (path, *(q for q in path.parents if q.is_relative_to(root))):
        if os.path.lexists(part):
            info = part.lstat()
            if stat.S_ISLNK(getattr(info, "st_mode", 0)) or getattr(info, "st_file_attributes", 0) & stat.FILE_ATTRIBUTE_REPARSE_POINT:
                raise ValueError(f"Отказ: ссылка/junction (reparse point) в пути {part}")


def check_tree(root: Path, folder: Path) -> None:
    check_path(folder, root)
    if not folder.exists():
        return
    for current, dirs, files in os.walk(folder, followlinks=False):
        for name in dirs + files:
            check_path(Path(current) / name, root)


def remaining_paths(folder: Path) -> list[Path]:
    """Отчёт не обходит junction даже в посторонних каталогах."""
    result, pending = [], [folder] if folder.is_dir() else []
    while pending:
        for path in pending.pop().iterdir():
            result.append(path)
            info = path.lstat()
            if stat.S_ISDIR(info.st_mode) and not getattr(info, "st_file_attributes", 0) & stat.FILE_ATTRIBUTE_REPARSE_POINT:
                pending.append(path)
    return result


def write_manifest(app: Path) -> None:
    from .cli import _atomic_write

    check_tree(app, app)
    files = {p.relative_to(app).as_posix(): digest(p) for p in app.rglob("*") if p.is_file() and p.name != MANIFEST}
    _atomic_write(app / MANIFEST, json.dumps(files, ensure_ascii=False, indent=2))


def removal_plan(root: Path, folder: Path) -> tuple[list[Path], list[Path]]:
    check_tree(root, folder)
    if not folder.exists():
        return [], []
    manifest = folder / MANIFEST
    if manifest.exists():
        data = json.loads(manifest.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError(f"Повреждён список собственных файлов: {manifest}")
        own = []
        for name, expected in data.items():
            relative = Path(name)
            if relative.is_absolute() or ".." in relative.parts or ":" in name:
                raise ValueError(f"Недопустимый путь в {manifest}: {name}")
            path = folder / relative
            check_path(path, root)
            if path.is_file() and digest(path) == expected:
                own.append(path)
        own.append(manifest)
    else:
        # Старый архив не имел манифеста: безопасно распознаём только основную программу.
        exe = folder / "office-live-mcp.exe"
        own = [exe] if exe.is_file() else []
    keep = [p for p in folder.rglob("*") if p.is_file() and p not in set(own)]
    return own, keep


def remove_files(root: Path, folder: Path) -> list[Path]:
    own, _ = removal_plan(root, folder)
    for path in own:
        check_path(path, root)
        path.unlink()
    for path in sorted((p for p in folder.rglob("*") if p.is_dir()), key=lambda p: len(p.parts), reverse=True):
        check_path(path, root)
        if not any(path.iterdir()):
            path.rmdir()
    if folder.exists() and not any(folder.iterdir()):
        folder.rmdir()
    return remaining_paths(folder) if folder.exists() else []
