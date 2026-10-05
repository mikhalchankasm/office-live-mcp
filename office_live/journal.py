"""Журнал по документам и добровольный лист «Лог»; служебные записи не вызывают инструменты."""

import hashlib
import ntpath
import re
import sys
import threading
import time
from collections import deque

import pywintypes

from . import config
from .errors import ToolError
from .registry import _brief, office_tool
from .wd_common import document_path, pick_document
from .xl_common import pick_workbook, sheet_names, suspend_events, workbook_path

LOG_NAME = "OfficeLive_Log"
LOG_SHEET = "Лог"
_lock = threading.Lock()
_warned = False


def _line(value):
    return str(value).replace("\r", "\\r").replace("\n", "\\n")


def journal_path(identity):
    base = ntpath.basename(identity)
    base = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", base).strip(" .")[:100] or "document"
    digest = hashlib.sha1(identity.lower().encode("utf-8")).hexdigest()[:12]
    return config.SETTINGS.journal_dir / f"{base} [{digest}].md"


def describe(arguments):
    sheet = arguments.get("dest_sheet") or arguments.get("sheet", "")
    cells = next((arguments[k] for k in ("dest_cell", "cells", "cell", "top_left", "dest", "source", "lines") if arguments.get(k)), "")
    where = f"{sheet}!{cells}" if sheet and cells and "!" not in str(cells) else cells
    summary = ", ".join(
        f"{key}={_brief(value)!r}" if isinstance(value, str) else f"{key}={_brief(value)}"
        for key, value in arguments.items()
        if key not in {"workbook", "document", "dest_workbook", "sheet", "dest_sheet", "cells", "cell", "top_left", "dest_cell"}
    )
    return _line(where), _line(summary)


def append(targets, tool, arguments, error=None, *, where=None, summary=None):
    """Ровно одна строка на каждую фактически выбранную цель, без COM-доступа."""
    global _warned
    if not config.SETTINGS.journal:
        return
    location, details = describe(arguments)
    location = location if where is None else _line(where)
    details = details if summary is None else _line(summary)
    result = "ok" if error is None else "ошибка: " + _line(str(error)[:200])
    row = f"- {time.strftime('%Y-%m-%d %H:%M:%S')} · {_line(tool)} · {location} · {details} · {result}\n"
    with _lock:
        for target in dict.fromkeys(targets or ()):
            try:
                kind, _, identity = target.partition(":")
                if kind not in {"workbook", "document"} or not identity:
                    continue
                path = journal_path(identity)
                path.parent.mkdir(parents=True, exist_ok=True)
                with path.open("a", encoding="utf-8", newline="\n") as stream:
                    if stream.tell() == 0:
                        stream.write(f"# {_line(identity)}\n")
                    stream.write(row)
            except Exception as exc:  # noqa: BLE001 — сбой одной цели не теряет записи остальных
                if not _warned:
                    _warned = True
                    print(f"[office-live] journal is not writable ({exc}); records are being lost", file=sys.stderr)


def sheet_enabled(wb):
    try:
        wb.Names(LOG_NAME)
        return True
    except (pywintypes.com_error, AttributeError, KeyError):
        return False


def append_sheet(app, wb, tool, arguments, *, where=None, summary=None):
    """Возвращает предупреждение вместо исключения; вызов внутри действующего COM-контекста."""
    try:
        if not sheet_enabled(wb):
            return None
        suspend_events(app, force=True)
        previous_book, previous = app.ActiveWorkbook, app.ActiveSheet
        try:
            if LOG_SHEET not in sheet_names(wb):
                ws = wb.Worksheets.Add(None, wb.Sheets(wb.Sheets.Count))
                ws.Name = LOG_SHEET
                ws.Range("A1:E1").Value = (("Время", "Действие", "Где", "Что сделано", "Результат"),)
                ws.Range("A1:E1").Font.Bold = True
            else:
                ws = wb.Worksheets(LOG_SHEET)
            row = max(2, int(ws.Cells(ws.Rows.Count, 1).End(-4162).Row) + 1)  # xlUp
            location, details = describe(arguments)
            target = ws.Range(f"A{row}:E{row}")
            target.NumberFormat = "@"
            target.Value = ((time.strftime("%Y-%m-%d %H:%M:%S"), _line(tool), location if where is None else _line(where),
                             details if summary is None else _line(summary), "ok"),)
        finally:
            if previous is not None:
                previous.Activate()
            elif previous_book is not None:
                previous_book.Activate()
    except Exception as exc:  # noqa: BLE001
        return str(exc)[:200]
    return None


@office_tool("history", "write", title="Document journal", read_actions=("read", "status"), destructive=True)
def office_journal(workbook: str = "", document: str = "", action: str = "read", lines: int = 50, delete_sheet: bool = False) -> dict:
    """Read a document's on-disk change journal or enable/disable the optional Excel 'Лог' sheet. Pass exactly one workbook or document name/path. Read and status are available in read-only mode.

    Args:
        workbook, document: exactly one target name/path.
        action: 'read', 'status', 'enable_sheet' or 'disable_sheet' (sheet actions are Excel only).
        lines: number of recent journal lines to read (positive).
        delete_sheet: disable_sheet also deletes the 'Лог' sheet; otherwise it is kept.
    """
    act = action.lower()
    if act not in {"read", "status", "enable_sheet", "disable_sheet"}:
        raise ToolError("action must be read, status, enable_sheet or disable_sheet.")
    if bool(workbook.strip()) == bool(document.strip()):
        raise ToolError("Pass exactly one workbook or document name/path.")
    if document and act not in {"read", "status"}:
        raise ToolError("Word supports the on-disk journal only; log sheets are Excel-only.")
    if lines < 1:
        raise ToolError("lines must be positive.")
    if workbook:
        app, obj = pick_workbook(workbook)
        identity = workbook_path(obj) or obj.Name
    else:
        app, obj = pick_document(document)
        identity = document_path(obj) or obj.Name
    if act == "enable_sheet":
        obj.Names.Add(LOG_NAME, "=TRUE", False)
    elif act == "disable_sheet":
        if sheet_enabled(obj):
            obj.Names(LOG_NAME).Delete()
        if delete_sheet and LOG_SHEET in sheet_names(obj):
            old = app.DisplayAlerts
            app.DisplayAlerts = False
            try:
                obj.Worksheets(LOG_SHEET).Delete()
            finally:
                app.DisplayAlerts = old
    path = journal_path(identity)
    out = {"file": identity, "journal_path": str(path), "journal_enabled": config.SETTINGS.journal,
           "sheet_enabled": sheet_enabled(obj) if workbook else False}
    if act == "read":
        try:
            with path.open(encoding="utf-8") as stream:
                out["lines"] = [line.rstrip("\n") for line in deque(stream, maxlen=lines)]
        except FileNotFoundError:
            out["lines"] = []
        except OSError as exc:
            raise ToolError(f"Could not read the journal: {exc}") from None
    return out
