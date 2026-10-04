"""Проверка путей: белый список каталогов, запрет «опасных» мест и неизвестных расширений."""

import os

from . import config
from .errors import ToolError

EXCEL_EXTS = {"xlsx", "xlsm", "xlsb", "xls", "csv", "pdf", "txt", "html", "htm", "ods", "xltx", "xltm"}
WORD_EXTS = {"docx", "docm", "doc", "rtf", "pdf", "txt", "html", "htm", "odt", "dotx", "dotm", "xml"}
IMAGE_EXTS = {"png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "emf", "wmf", "svg"}


def _deny_dirs() -> list[str]:
    env = os.environ
    appdata = env.get("APPDATA", "")
    pdata = env.get("ProgramData", "")
    candidates = [
        os.path.join(appdata, "Microsoft", "Excel", "XLSTART"),
        os.path.join(appdata, "Microsoft", "Word", "STARTUP"),
        os.path.join(appdata, "Microsoft", "Windows", "Start Menu", "Programs", "Startup"),
        os.path.join(pdata, "Microsoft", "Windows", "Start Menu", "Programs", "StartUp"),
        env.get("WINDIR", ""),
        env.get("ProgramFiles", ""),
        env.get("ProgramFiles(x86)", ""),
    ]
    return [os.path.normcase(os.path.abspath(c)) for c in candidates if c]


def _inside(path: str, root: str) -> bool:
    try:
        return os.path.commonpath([path, root]) == root
    except ValueError:  # разные диски
        return False


def check_path(path: str, purpose: str = "read", exts: set | None = None, settings=None) -> str:
    """Возвращает нормализованный абсолютный путь или бросает ToolError.

    purpose: 'read' (открыть/вставить файл) или 'write' (сохранить/экспортировать).
    exts: допустимые расширения без точки (для write).
    """
    cfg = settings or config.SETTINGS
    if not path or not str(path).strip():
        raise ToolError("Path is empty.")
    full = os.path.abspath(os.path.expandvars(str(path).strip().strip('"')))
    norm = os.path.normcase(full)

    if cfg.allowed_dirs:
        allowed = [os.path.normcase(str(d)) for d in cfg.allowed_dirs]
        if not any(_inside(norm, a) for a in allowed):
            raise ToolError(
                f"Path is outside the allowed directories (OFFICE_LIVE_ALLOWED_DIRS): {full}. "
                f"Allowed: {[str(d) for d in cfg.allowed_dirs]}"
            )

    if purpose == "write":
        for bad in _deny_dirs():
            if _inside(norm, bad):
                raise ToolError(f"Writing into a system/startup directory is not allowed: {full}")
        ext = os.path.splitext(full)[1].lstrip(".").lower()
        if exts is not None and ext not in exts:
            raise ToolError(f"Unsupported file extension '.{ext}'. Allowed: {sorted(exts)}")
    return full
