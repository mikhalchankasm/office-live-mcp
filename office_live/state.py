"""Состояние установки и блокировка совместных изменений конфигурации."""

import json
import os
import time
from contextlib import contextmanager
from pathlib import Path

from . import __version__


def state_path() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / "office-live-mcp" / "state.json"


@contextmanager
def file_lock(path: Path, timeout: float = 5):
    """Постоянный файл-замок: удаление после unlock создало бы гонку между ожидающими процессами."""
    import msvcrt

    lock = path.with_name(path.name + ".lock")
    lock.parent.mkdir(parents=True, exist_ok=True)
    with lock.open("a+b") as stream:
        if stream.tell() == 0:
            stream.write(b"\0")
            stream.flush()
        deadline = time.monotonic() + timeout
        while True:
            stream.seek(0)
            try:
                msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
                break
            except OSError:
                if time.monotonic() >= deadline:
                    raise OSError(f"Файл занят другим установщиком: {path}. Повторите команду позже.") from None
                time.sleep(0.05)
        try:
            yield
        finally:
            stream.seek(0)
            msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)


def read_state() -> dict:
    path = state_path()
    if not path.exists():
        return {"schema": 1, "root": "", "version": "", "registrations": []}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if (not isinstance(data, dict) or data.get("schema") != 1 or not isinstance(data.get("root"), str)
                or not isinstance(data.get("registrations"), list)):
            raise ValueError("неверная структура")
        for record in data["registrations"]:
            if (not isinstance(record, dict)
                    or any(not isinstance(record.get(k), str) for k in ("client", "scope", "path", "name", "command", "root", "fingerprint", "time"))
                    or not isinstance(record.get("keys"), list) or not all(isinstance(k, str) for k in record["keys"])
                    or not record["keys"] or record["scope"] not in {"user", "project", "custom"}
                    or not Path(record["path"]).is_absolute() or not Path(record["root"]).is_absolute()):
                raise ValueError("повреждена регистрация")
        return data
    except (ValueError, OSError) as exc:
        raise ValueError(f"Не удалось прочитать {path}: {exc}. Восстановите файл состояния из копии; ничего не перезаписано.") from None


def write_state(data: dict) -> None:
    from .cli import _atomic_write

    if state_path().exists():
        _atomic_write(state_path().with_suffix(".json.bak"), state_path().read_text(encoding="utf-8"))
    _atomic_write(state_path(), json.dumps(data, ensure_ascii=False, indent=2))


def remember_install(root: Path) -> None:
    with file_lock(state_path()):
        data = read_state()
        data.update(root=str(root.resolve()), version=__version__)
        write_state(data)
