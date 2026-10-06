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
from .privacy import redact, safe_text
from .registry import _brief, office_tool
from .wd_common import document_path, pick_document
from .xl_common import pick_workbook, sheet_names, suspend_events, workbook_path

LOG_NAME = "OfficeLive_Log"
LOG_SHEET = "Лог"
SHEET_MARKER = "OfficeLive_LogSheet"
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
    arguments = redact(arguments)
    sheet = arguments.get("dest_sheet") or arguments.get("sheet", "")
    cells = next((arguments[k] for k in ("dest_cell", "cells", "cell", "top_left", "dest", "source", "lines") if arguments.get(k)), "")
    where = f"{sheet}!{cells}" if sheet and cells and "!" not in str(cells) else cells
    fragment_edit = "expected_text" in arguments and "new_text" in arguments
    summary = ", ".join(
        f"{key}={(value[:200] if fragment_edit and key in {'expected_text', 'new_text'} else _brief(value))!r}" if isinstance(value, str) else f"{key}={_brief(value)}"
        for key, value in arguments.items()
        if key not in {"workbook", "document", "dest_workbook", "sheet", "dest_sheet", "cells", "cell", "top_left", "dest_cell"}
    )
    return _line(where), _line(summary)


def target_links(items, kind, identity):
    from .links import parse

    out = []
    for link in items or ():
        try:
            app, params = parse(link["uri"])
            if (app == "excel") == (kind == "workbook") and params.get("book", params.get("doc")) == identity:
                out.append(link)
        except (ToolError, KeyError, TypeError):
            continue
    return out


def append(targets, tool, arguments, error=None, *, where=None, summary=None, links=None):
    """Ровно одна строка на каждую фактически выбранную цель, без COM-доступа."""
    global _warned
    if not config.SETTINGS.journal:
        return
    location, details = describe(arguments)
    location = location if where is None else _line(safe_text(where, arguments))
    details = details if summary is None else _line(safe_text(summary, arguments))
    result = "ok" if error is None else "ошибка: " + _line(safe_text(error, arguments)[:200])
    with _lock:
        for target in dict.fromkeys(targets or ()):
            try:
                kind, _, identity = target.partition(":")
                if kind not in {"workbook", "document"} or not identity:
                    continue
                from .links import markdown

                places = target_links(links, kind, identity) if error is None else []
                linked_location = ", ".join(markdown(link) for link in places) or location
                row = f"- {time.strftime('%Y-%m-%d %H:%M:%S')} · {_line(tool)} · {linked_location} · {details} · {result}\n"
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


def owns_sheet(ws):
    try:
        return str(ws.Names(SHEET_MARKER).RefersTo).upper() == "=TRUE"
    except (pywintypes.com_error, AttributeError, KeyError):
        return False


def _existing_sheet(wb):
    return next((wb.Worksheets(name) for name in sheet_names(wb) if name.casefold() == LOG_SHEET.casefold()), None)


def _check_owner(ws):
    if ws is not None and not owns_sheet(ws):
        raise ToolError("The sheet 'Лог' belongs to you, not Office Live. Rename it before enabling the log sheet.")


def append_sheet(app, wb, tool, arguments, *, where=None, summary=None, links=None):
    """Возвращает предупреждение вместо исключения; вызов внутри действующего COM-контекста."""
    try:
        if not sheet_enabled(wb):
            return None
        suspend_events(app, force=True)
        previous_book, previous = app.ActiveWorkbook, app.ActiveSheet
        try:
            ws = _existing_sheet(wb)
            _check_owner(ws)
            if ws is None:
                ws = wb.Worksheets.Add(None, wb.Sheets(wb.Sheets.Count))
                ws.Name = LOG_SHEET
                ws.Names.Add(SHEET_MARKER, "=TRUE", False)
                ws.Range("A1:E1").Value = (("Время", "Действие", "Где", "Что сделано", "Результат"),)
                ws.Range("A1:E1").Font.Bold = True
            last_row = int(ws.Rows.Count)
            bottoms = [ws.Cells(last_row, col) for col in range(1, 6)]
            if any(cell.Value is not None for cell in bottoms):
                raise ToolError("The 'Лог' sheet is full; no journal row was written.")
            row = max(2, max(int(cell.End(-4162).Row) for cell in bottoms) + 1)  # xlUp
            location, details = describe(arguments)
            target = ws.Range(f"A{row}:E{row}")
            target.NumberFormat = "@"
            target.Value = ((time.strftime("%Y-%m-%d %H:%M:%S"), _line(tool), location if where is None else _line(safe_text(where, arguments)),
                             details if summary is None else _line(safe_text(summary, arguments)), "ok"),)
            # One row has one address cell: use the first actual destination in this workbook.
            from .links import parse

            for link in target_links(links, "workbook", workbook_path(wb) or wb.Name):
                _, params = parse(link["uri"])
                name, address = params.get("sheet"), params.get("range")
                if name and address and name != ws.Name and name in sheet_names(wb):
                    subaddress = "'" + name.replace("'", "''") + "'!" + address
                    ws.Hyperlinks.Add(ws.Cells(row, 3), "", subaddress, "Office Live", link["label"])
                    break
        finally:
            if previous is not None:
                previous.Activate()
            elif previous_book is not None:
                previous_book.Activate()
    except Exception as exc:  # noqa: BLE001
        return safe_text(exc, arguments)[:200]
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
        _check_owner(_existing_sheet(obj))
        obj.Names.Add(LOG_NAME, "=TRUE", False)
    elif act == "disable_sheet":
        if sheet_enabled(obj):
            obj.Names(LOG_NAME).Delete()
        ws = _existing_sheet(obj) if delete_sheet else None
        if ws is not None:
            _check_owner(ws)
            old = app.DisplayAlerts
            app.DisplayAlerts = False
            try:
                ws.Delete()
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
