"""Excel: книги, листы, чтение/запись, формулы, структура, сортировка, фильтры, имена, выделение."""

import os
import re
import unicodedata

import pywintypes

from . import com
from .errors import PartialChangeError, ToolError
from .registry import office_tool
from .safety import EXCEL_EXTS, check_path
from .util import (
    cell_kind, MAX_COLS, MAX_ROWS, a1_cell, a1_range, col_letter, col_number, parse_a1,
    DATE_FORMATS, excel_date_serial, parse_color, quote_sheet, smart_number, to_com_grid, to_grid,
)
from .xl_common import (
    addr_of, all_workbooks, localize_formula, workbook_allowed, suspend_events, bounded_range, bounds, clip_to_used, count_nonempty, error_cells, formula_cells, get_range,
    pick_sheet, pick_workbook, preview, read_grid, ref_label, sheet_is_empty, sheet_names, sub_range,
    validate_sheet_name, number_format_for_write,
)

Cell = str | int | float | bool | None
GRID_HINT = "2-D array: a list of rows, each row a list of cell values (a flat list is treated as one row)."

# коды форматов SaveAs
_SAVE_FORMATS = {"xlsx": 51, "xlsm": 52, "xlsb": 50, "xls": 56, "csv": 62, "txt": 42, "html": 44, "htm": 44, "ods": 60, "xltx": 54, "xltm": 53}


def _autosave(wb) -> bool | None:
    try:
        return bool(wb.AutoSaveOn)
    except (pywintypes.com_error, AttributeError):
        return None


def _pid(app):
    try:
        import win32process

        return win32process.GetWindowThreadProcessId(int(app.Hwnd))[1]
    except Exception:
        return None


def _is_visible(wb) -> bool:
    try:
        return wb.Windows.Count > 0 and bool(wb.Windows(1).Visible)
    except pywintypes.com_error:
        return False


def _used_address(app, ws):
    return None if sheet_is_empty(app, ws) else addr_of(ws.UsedRange)


def _sheet_list(wb) -> list[str]:
    return sheet_names(wb, any_type=True)


def _restore_alerts(app, previous):
    try:
        app.DisplayAlerts = previous
    except pywintypes.com_error:
        pass


# ================================================================== книги


@office_tool("excel_core", "read", title="List open workbooks")
def excel_list_workbooks() -> dict:
    """List every open Excel workbook across ALL running Excel instances: name, path, saved flag, sheets, visibility, and which one is active. Call this first. Returns an empty list (not an error) when Excel has no workbooks."""
    apps = com.apps("excel")
    books = []
    active_workbook = active_sheet = None
    for idx, app in enumerate(apps):
        act = app.ActiveWorkbook
        act_allowed = act is not None and workbook_allowed(act)  # название запрещённой книги не раскрываем и как «активной»
        act_name = act.Name if act_allowed else None
        if idx == 0 and act_allowed:
            active_workbook = act_name
            sh = app.ActiveSheet
            active_sheet = sh.Name if sh is not None else None
        for i in range(1, app.Workbooks.Count + 1):
            wb = app.Workbooks(i)
            if not workbook_allowed(wb):
                continue  # вне OFFICE_LIVE_ALLOWED_DIRS: даже имя не показываем
            books.append({
                "name": wb.Name,
                "path": wb.Path or None,
                "saved": bool(wb.Saved),
                "read_only": bool(wb.ReadOnly),
                "visible": _is_visible(wb),
                "sheets": _sheet_list(wb),
                "active": idx == 0 and act_name == wb.Name,
                "instance": idx,
                "autosave": _autosave(wb),
            })
    return {
        "excel_running": True,
        "instances": len(apps),
        "workbooks": books,
        "active_workbook": active_workbook,
        "active_sheet": active_sheet,
    }


@office_tool("excel_core", "read", title="Workbook overview")
def excel_workbook_info(workbook: str = "") -> dict:
    """Detailed overview of one workbook: every sheet with its used range, visibility, tables, charts, pivot tables, shapes, protection and filters, plus defined names and calculation mode.

    Args:
        workbook: exact name, unique part of the name, full path, or '' for the active workbook.
    """
    app, wb = pick_workbook(workbook)
    sheets = []
    for i in range(1, wb.Sheets.Count + 1):
        sh = wb.Sheets(i)
        info = {"index": i, "name": sh.Name, "visible": {-1: "visible", 0: "hidden", 2: "very_hidden"}.get(int(sh.Visible), "visible")}
        try:
            if int(sh.Type) != -4167:  # лист диаграммы
                info["type"] = "chart_sheet"
                sheets.append(info)
                continue
        except (pywintypes.com_error, AttributeError):
            pass
        info["type"] = "worksheet"
        used = sh.UsedRange
        if sheet_is_empty(app, sh):
            info.update({"used_range": None, "rows": 0, "columns": 0})
        else:
            r1, c1, r2, c2 = bounds(used)
            info.update({"used_range": addr_of(used), "rows": r2 - r1 + 1, "columns": c2 - c1 + 1})
        info["tables"] = [sh.ListObjects(j).Name for j in range(1, sh.ListObjects.Count + 1)]
        info["charts"] = int(sh.ChartObjects().Count)
        info["pivot_tables"] = int(sh.PivotTables().Count)
        info["shapes"] = int(sh.Shapes.Count)
        info["protected"] = bool(sh.ProtectContents)
        info["autofilter"] = bool(sh.AutoFilterMode)
        sheets.append(info)
    names = []
    for j in range(1, min(wb.Names.Count, 100) + 1):
        nm = wb.Names(j)
        if bool(nm.Visible):  # скрытые служебные имена (_FilterDatabase и т.п.) не показываем
            names.append({"name": nm.Name, "refers_to": nm.RefersTo})
    calc = {-4105: "automatic", -4135: "manual", 2: "semiautomatic"}.get(int(app.Calculation), str(app.Calculation))
    act = wb.ActiveSheet
    return {
        "workbook": wb.Name,
        "path": wb.FullName if wb.Path else None,
        "saved": bool(wb.Saved),
        "read_only": bool(wb.ReadOnly),
        "autosave": _autosave(wb),
        "calculation": calc,
        "active_sheet": act.Name if act is not None else None,
        "sheets": sheets,
        "defined_names": names,
        "excel_version": str(app.Version),
        "decimal_separator": app.International[2],
        "list_separator": app.International[4],
    }


@office_tool("excel_core", "open", title="Open workbook")
def excel_open_workbook(path: str, read_only: bool = False) -> dict:
    """Open an existing workbook file in Excel (starts Excel if needed). Macros are disabled during opening. If the file is already open, returns it instead of opening twice.

    Args:
        path: full path to .xlsx/.xlsm/.xlsb/.xls/.csv file.
        read_only: open read-only (always forced when the server runs in readonly mode).
    """
    from . import config

    full = check_path(path, "read")
    if not os.path.isfile(full):
        raise ToolError(f"File not found: {full}")
    for _, wb in (all_workbooks() if _excel_running() else []):
        if wb.Path and os.path.normcase(os.path.abspath(wb.FullName)) == os.path.normcase(full):
            return {"ok": True, "already_open": True, "workbook": wb.Name, "path": wb.FullName, "sheets": _sheet_list(wb)}
    app = com.primary_app("excel", launch=True)
    force_ro = bool(read_only or config.SETTINGS.readonly)
    try:
        with com.macros_disabled(app):
            # Open(Filename, UpdateLinks, ReadOnly, Format, Password): пароль-пустышка — зашифрованный файл не зависнет на диалоге
            wb = app.Workbooks.Open(full, 0, force_ro, None, "__office_live_no_password__")
    except pywintypes.com_error as exc:
        raise ToolError(f"Excel could not open {full}: {com.com_error_text(exc)} (password-protected or corrupt file?)") from None
    return {"ok": True, "already_open": False, "workbook": wb.Name, "path": wb.FullName, "read_only": bool(wb.ReadOnly), "sheets": _sheet_list(wb)}


def _excel_running() -> bool:
    try:
        com.apps("excel")
        return True
    except ToolError:
        return False


@office_tool("excel_core", "write", title="New workbook")
def excel_new_workbook(sheets: list[str] | None = None) -> dict:
    """Create a new empty workbook in Excel (starts Excel if needed). The workbook is unsaved until excel_save_as.

    Args:
        sheets: optional sheet names to create (the first default sheet is renamed, the rest are added).
    """
    app = com.primary_app("excel", launch=True)
    suspend_events(app)
    wb = app.Workbooks.Add()
    if sheets:
        for nm in sheets:
            validate_sheet_name(nm)
        wb.Worksheets(1).Name = sheets[0]
        for nm in sheets[1:]:
            last = wb.Worksheets(wb.Worksheets.Count)
            ws = wb.Worksheets.Add(None, last)
            ws.Name = nm
        wb.Worksheets(1).Activate()
    return {"ok": True, "workbook": wb.Name, "sheets": _sheet_list(wb)}


@office_tool("excel_core", "destructive", title="Close workbook")
def excel_close_workbook(workbook: str, save: bool = False, discard: bool = False) -> dict:
    """Close an open workbook. Refuses when it has unsaved changes unless you choose: save=true (saves first, needs an existing path) or discard=true (throws the changes away). Never quits Excel itself.

    Args:
        workbook: exact workbook name or full path (required).
        save: save before closing.
        discard: close without saving unsaved changes.
    """
    if not workbook:
        raise ToolError("'workbook' is required for closing.")
    app, wb = pick_workbook(workbook, allow_autosave=True)
    name = wb.Name
    saved_first = False
    if not wb.Saved:
        if save:
            if not wb.Path:
                raise ToolError(f"{name} has never been saved; use excel_save_as first or pass discard=true.")
            check_path(wb.FullName, "write")  # зона доступа и запретные каталоги — и для сохранения уже открытой книги
            wb.Save()
            saved_first = True
        elif not discard:
            raise ToolError(f"{name} has unsaved changes. Pass save=true to save them or discard=true to drop them.")
    wb.Close(False)
    remaining = [w.Name for a, w in all_workbooks()] if _excel_running() else []
    return {"ok": True, "closed": name, "saved_before_close": saved_first, "open_workbooks": remaining}


@office_tool("excel_core", "save", title="Save workbook")
def excel_save(workbook: str = "") -> dict:
    """Save a workbook to its current file (Ctrl+S). A workbook that was never saved needs excel_save_as instead.

    Args:
        workbook: exact name or '' for the active workbook.
    """
    app, wb = pick_workbook(workbook, allow_autosave=True)
    if not wb.Path:
        raise ToolError(f"{wb.Name} has never been saved - use excel_save_as with a path.")
    if wb.ReadOnly:
        raise ToolError(f"{wb.Name} is open read-only; use excel_save_as to write a copy.")
    check_path(wb.FullName, "write")  # зона доступа и запретные каталоги
    wb.Save()
    return {"ok": True, "workbook": wb.Name, "path": wb.FullName, "saved": bool(wb.Saved)}


@office_tool("excel_core", "save", title="Save workbook as")
def excel_save_as(workbook: str, path: str, overwrite: bool = False) -> dict:
    """Save a workbook under a new name/format. Format follows the extension: .xlsx .xlsm .xlsb .xls .csv (UTF-8) .txt .html .ods. Refuses to replace an existing file unless overwrite=true. Afterwards the workbook is known by its NEW file name.

    Args:
        workbook: exact name of the open workbook (required).
        path: destination full path with extension.
        overwrite: allow replacing an existing file.
    """
    app, wb = pick_workbook(workbook, allow_autosave=True)
    full = check_path(path, "write", EXCEL_EXTS - {"pdf"})
    if os.path.exists(full) and not overwrite:
        raise ToolError(f"File already exists: {full}. Pass overwrite=true to replace it.")
    ext = full.rsplit(".", 1)[-1].lower()
    fmt = _SAVE_FORMATS.get(ext)
    if fmt is None:
        raise ToolError(f"Unsupported extension .{ext}. Use one of {sorted(_SAVE_FORMATS)}; for PDF use excel_export_pdf.")
    os.makedirs(os.path.dirname(full), exist_ok=True)
    before = app.DisplayAlerts
    app.DisplayAlerts = False
    try:
        wb.SaveAs(full, fmt)
    finally:
        _restore_alerts(app, before)
    return {"ok": True, "workbook": wb.Name, "path": wb.FullName, "overwrite": overwrite, "note": "The workbook is now named after the new file."}


@office_tool("excel_core", "save", title="Export to PDF")
def excel_export_pdf(workbook: str, path: str, sheet: str = "", overwrite: bool = False) -> dict:
    """Export a workbook (or one sheet) to a PDF file using its print settings.

    Args:
        workbook: exact workbook name (required).
        path: destination .pdf path.
        sheet: sheet name to export only that sheet; '' exports the whole workbook.
        overwrite: allow replacing an existing file.
    """
    app, wb = pick_workbook(workbook, allow_autosave=True)
    full = check_path(path, "write", {"pdf"})
    if os.path.exists(full) and not overwrite:
        raise ToolError(f"File already exists: {full}. Pass overwrite=true to replace it.")
    os.makedirs(os.path.dirname(full), exist_ok=True)
    target = pick_sheet(wb, sheet, any_type=True) if sheet else wb
    # ExportAsFixedFormat(Type=0 pdf, Filename, Quality=0, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish)
    target.ExportAsFixedFormat(0, full, 0, True, False, None, None, False)
    return {"ok": True, "path": full, "size_bytes": os.path.getsize(full) if os.path.exists(full) else None}


# ================================================================== чтение


@office_tool("excel_core", "read", title="Read range")
def excel_read_range(
    workbook: str = "",
    sheet: str = "",
    cells: str = "",
    max_cells: int = 5000,
    include_formulas: bool = False,
    display_text: bool = False,
) -> dict:
    """Read the values of a range. Dates come back as ISO strings, errors as '#DIV/0!' etc. Whole-column/row references and an empty `cells` are clipped to the used range. Large reads are truncated at max_cells and the response tells you the remaining range.

    Args:
        workbook: exact name or '' for the active workbook.
        sheet: sheet name or '' for the active sheet.
        cells: 'A1:J10', 'A:C', 'Sheet2!B2:C5', a defined name, or '' for the whole used range.
        max_cells: cap on returned cells (default 5000, max 100000).
        include_formulas: also return a sparse list of the formulas in the range.
        display_text: return the text exactly as shown on screen (formatted numbers/dates); slower, max 3000 cells.
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells)
    if int(rng.Areas.Count) > 1:
        raise ToolError("Multi-area ranges cannot be read at once; read each area separately.")
    clipped = clip_to_used(app, ws, rng)
    out = {"workbook": wb.Name, "sheet": ws.Name}
    if clipped is None:
        out.update({"cells": cells, "rows": 0, "columns": 0, "values": [], "truncated": False})
        return out
    r1, c1, r2, c2 = bounds(clipped)
    rows, cols = r2 - r1 + 1, c2 - c1 + 1
    limit = max(1, min(int(max_cells), 100000))
    if display_text:
        limit = min(limit, 3000)
    if cols > limit:
        raise ToolError(f"Range has {cols} columns, more than max_cells={limit}. Narrow the columns.")
    target, truncated, remaining = clipped, False, None
    if rows * cols > limit:
        keep = max(1, limit // cols)
        r2k = r1 + keep - 1
        target = sub_range(ws, r1, c1, r2k, c2)
        truncated, remaining = True, a1_range(r2k + 1, c1, r2, c2)
        rows = keep
    grid = read_grid(target, "text" if display_text else "values")
    out.update({
        "cells": addr_of(target), "start_cell": a1_cell(r1, c1), "rows": rows, "columns": cols,
        "values": grid, "truncated": truncated,
    })
    if remaining:
        out["remaining_range"] = remaining
    if not cells:
        out["note"] = "Whole used range."
    if include_formulas:
        out["formulas"] = formula_cells(target)
    return out


@office_tool("excel_core", "read", title="Find in workbook")
def excel_find(
    query: str,
    workbook: str = "",
    sheet: str = "",
    all_sheets: bool = False,
    look_in: str = "values",
    match_case: bool = False,
    whole_cell: bool = False,
    regex: bool = False,
    max_results: int = 100,
    count_all: bool = False,
) -> dict:
    """Search cell values or formulas. Returns matching cells with their address and content.

    Args:
        query: text to find (or a regular expression when regex=true).
        workbook: exact name or '' for the active workbook.
        sheet: restrict to this sheet ('' = active sheet unless all_sheets=true).
        all_sheets: search every worksheet of the workbook.
        look_in: 'values' (default) or 'formulas'.
        match_case, whole_cell, regex: matching options.
        max_results: stop searching once more than this many matches exist (the answer then says truncated=true).
        count_all: keep scanning after max_results to count every match (up to 100000) - slower on big sheets.
    """
    if not query:
        raise ToolError("'query' is empty.")
    flags = 0 if match_case else re.IGNORECASE
    try:
        pat = re.compile(query if regex else re.escape(query), flags)
    except re.error as exc:
        raise ToolError(f"Bad regular expression: {exc}") from None
    app, wb = pick_workbook(workbook)
    if all_sheets:
        sheets = [wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1)]
    else:
        sheets = [pick_sheet(wb, sheet)]
    results, total, cap = [], 0, max(1, int(max_results))
    what = "formulas" if look_in == "formulas" else "values"
    stop = False
    for ws in sheets:
        if stop:
            break
        if sheet_is_empty(app, ws):
            continue
        used = ws.UsedRange
        r1, c1, r2, c2 = bounds(used)
        step = max(1, 20000 // max(1, c2 - c1 + 1))
        for start in range(r1, r2 + 1, step):
            if stop:
                break
            end = min(r2, start + step - 1)
            grid = read_grid(sub_range(ws, start, c1, end, c2), what)
            for i, row in enumerate(grid):
                for j, v in enumerate(row):
                    if v is None or v == "":
                        continue
                    s = v if isinstance(v, str) else str(v)
                    if not (pat.fullmatch(s) if whole_cell else pat.search(s)):
                        continue
                    total += 1
                    if len(results) < cap:
                        results.append({"sheet": ws.Name, "cell": a1_cell(start + i, c1 + j), "value": v})
                    elif not count_all:
                        stop = True  # нашлось больше, чем просили, — дальше не ищем
                    if total >= 100000:
                        stop = True
                    if stop:
                        break
                if stop:
                    break
    out = {"workbook": wb.Name, "query": query, "total_matches": total, "returned": len(results), "truncated": stop, "matches": results}
    if stop:
        out["note"] = "Search stopped early: total_matches is a lower bound. Raise max_results, narrow the sheet/range, or pass count_all=true."
    return out


@office_tool("excel_core", "read", title="Get user selection")
def excel_get_selection(max_cells: int = 200) -> dict:
    """What the user currently has selected in Excel (active workbook/sheet/range) and the values in it. Use this for requests like 'format what I selected' or 'sum this column'."""
    apps = com.apps("excel")
    app = apps[0]
    wb = app.ActiveWorkbook
    if wb is None:
        return {"workbook": None, "selection": None}
    if not workbook_allowed(wb):
        return {"workbook": None, "selection": None, "note": "The active workbook is outside the folders this server may use; the user has to activate another workbook."}
    ws = app.ActiveSheet
    sel = app.Selection
    out = {"workbook": wb.Name, "sheet": ws.Name if ws is not None else None}
    try:
        addr = sel.Address.replace("$", "")
    except (pywintypes.com_error, AttributeError):
        out["selection"] = None
        out["note"] = "The selection is not a cell range (a chart or shape is selected)."
        return out
    out["selection"] = addr
    try:
        r1, c1, r2, c2 = bounds(sel)
        rows, cols = r2 - r1 + 1, c2 - c1 + 1
        out.update({"rows": rows, "columns": cols})
        if int(sel.Areas.Count) == 1 and rows * cols <= max_cells:
            out["values"] = to_grid(sel.Value)
        else:
            out["note"] = "Selection is large or multi-area; read it with excel_read_range."
    except pywintypes.com_error:
        pass
    return out


@office_tool("excel_core", "ui", title="Select range")
def excel_select_range(workbook: str, sheet: str, cells: str) -> dict:
    """Activate a sheet and select (and scroll to) a range so the user sees it - useful to point at a result.

    Args:
        workbook: exact workbook name.
        sheet: sheet name.
        cells: range to select.
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells, empty_means_used=False)
    return select_resolved_range(app, wb, ws, rng)


def select_resolved_range(app, wb, ws, rng):
    """Shared by the tool and the URI handler after all target validation."""
    try:
        wb.Activate()
        ws.Activate()
        app.Goto(rng, True)
    except pywintypes.com_error as exc:
        raise ToolError("Could not select the range (the workbook window may be hidden): " + com.com_error_text(exc)) from None
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "selected": addr_of(rng)}


# ================================================================== запись


def _clean_text_value(value, operations, decimal_separator, explicit_decimal):
    if "clean" in operations:
        value = "".join(c for c in value if (unicodedata.category(c) != "Cc" or c == "\n") and c not in "\u200b\u200c\u200d\ufeff")
    if "trim" in operations:
        value = re.sub(r" +", " ", value.replace("\u00a0", " ").replace("\u202f", " ")).strip(" ")
    for op in ("upper", "lower", "proper"):
        if op in operations:
            value = value.title() if op == "proper" else getattr(value, op)()
    reason = ""
    if "text_to_number" in operations:
        core = value.strip().removesuffix("%").strip().strip("()").lstrip("+-").strip()
        if re.match(r"^0\d", core):
            reason = "leading_zero"
        elif not explicit_decimal and re.fullmatch(r"\d+[.,]\d{3}", core):
            reason = "ambiguous"
        elif not re.fullmatch(r"\+\d{7,}", value.strip()):
            number = smart_number(value, decimal_separator)
            if number is not None:
                value = number
    return value, reason


def _clean_input(value, prefix):
    if isinstance(value, str) and (prefix == "'" or value.startswith(("=", "+", "-", "@", "'"))
                                   or any(c.isdigit() for c in value) or value.casefold() in {"true", "false", "истина", "ложь"}):
        return "'" + value
    return value


def _clean_matches(cell, expected):
    actual = cell.Value2
    if isinstance(expected, str):
        return not bool(cell.HasFormula) and (actual == expected or (expected == "" and actual is None))
    return not isinstance(actual, (str, bool)) and actual == expected


@office_tool("excel_core", "write", title="Clean text", preview_arg="preview")
def excel_clean_text(
    workbook: str,
    sheet: str,
    cells: str,
    operations: list[str],
    decimal_separator: str = "",
    preview: bool = True,
    date_format: str = "",
    date_number_format: str = "",
) -> dict:
    """Clean text constants or parse dates by an explicit format, with a preview by default; preserve formulas, literal text and leading-zero codes. Writes only changed cells in vertical runs, verifies their types, and supports office_undo even after a partial COM failure.

    Args:
        workbook, sheet: exact target workbook and sheet.
        cells: required rectangle (whole rows/columns clipped to used range), at most 100000 cells.
        operations: nonempty list of trim, clean, upper/lower/proper (at most one case operation), text_to_number or text_to_date. Fixed order: clean, trim, case, conversion.
        decimal_separator: comma or dot; empty uses Excel's separator and skips ambiguous 1,234 / 1.234.
        date_format: required for text_to_date: DD.MM.YYYY, DD/MM/YYYY, MM/DD/YYYY, YYYY-MM-DD, DD-MM-YYYY or YYYY.MM.DD, optionally followed by HH:MM or HH:MM:SS.
        date_number_format: Excel display code; default is the lower-case date_format. Dates use the workbook's 1900/1904 system; invalid or pre-epoch dates remain unchanged (skipped_not_date).
        preview: true returns counts and up to 50 examples without writes, undo history or change journal; false writes only after the undo snapshot is ready (refuses if it cannot be made; with OFFICE_LIVE_UNDO=0 it writes without one).

    Protected sheets and pivot intersections are refused. Non-text, formulas and merged-cell followers are skipped.
    trim collapses spaces/NBSP while preserving newlines; clean removes controls except LF and zero-width characters.
    Text-to-number leaves dates, phone-like strings, leading zeros and numbers beyond 15 digits as text; @ changes to General only for converted numbers.
    """
    if not isinstance(operations, list) or not operations or any(not isinstance(op, str) or op not in {"trim", "clean", "upper", "lower", "proper", "text_to_number", "text_to_date"} for op in operations):
        raise ToolError("operations must be a nonempty list of trim, clean, upper, lower, proper, text_to_number, text_to_date.")
    if len(set(operations)) != len(operations) or len(set(operations) & {"upper", "lower", "proper"}) > 1:
        raise ToolError("operations must be unique, with at most one of upper, lower, proper.")
    if decimal_separator not in {"", ",", "."}:
        raise ToolError("decimal_separator must be ',' or '.' or empty.")
    dates = "text_to_date" in operations
    if dates and (date_format not in DATE_FORMATS or "text_to_number" in operations):
        raise ToolError("text_to_date requires a supported date_format and cannot be combined with text_to_number.")
    if not dates and (date_format or date_number_format):
        raise ToolError("date_format/date_number_format require text_to_date.")
    app, wb = pick_workbook(workbook)
    date1904 = bool(wb.Date1904) if dates else False
    date_style = number_format_for_write(app, date_number_format or date_format.lower()) if dates else None
    if not preview and wb.ReadOnly:
        raise ToolError("Workbook is read-only; no cells were changed.")
    ws, rng = bounded_range(app, wb, sheet, cells, 100000)
    if bool(ws.ProtectContents):
        raise ToolError("The sheet is protected (ProtectContents); no cells were changed.")
    result = {"preview": preview, "workbook": wb.Name, "sheet": ws.Name, "changed": 0, "skipped_formulas": 0,
              "skipped_non_text": 0, "skipped_merged": 0, "skipped_ambiguous": 0, "skipped_leading_zero": 0,
              "ambiguous": [], "examples": [], "format_changed": [], "rewritten_as_text": [], "skipped_not_date": 0, "not_date": []}
    if rng is None:
        return result
    pivots = ws.PivotTables()
    overlaps = [pivots(i).Name for i in range(1, int(pivots.Count) + 1) if app.Intersect(rng, pivots(i).TableRange2) is not None]
    if overlaps:
        raise ToolError(f"Range intersects pivot tables: {overlaps}; no cells were changed.")
    separator = decimal_separator or app.International[2]
    r1, c1, _, _ = bounds(rng)
    raw, raw_formulas = rng.Value2, rng.Formula
    values = raw if isinstance(raw, tuple) and raw and isinstance(raw[0], tuple) else ((raw,),)
    formulas = raw_formulas if isinstance(raw_formulas, tuple) and raw_formulas and isinstance(raw_formulas[0], tuple) else ((raw_formulas,),)
    any_merged = rng.MergeCells is not False  # None — частично объединён
    changes = []
    # Массивы читаются один раз; поячеечные вызовы COM — только для ячеек, которые действительно изменятся.
    for j in range(int(rng.Columns.Count)):
        for i, row in enumerate(values):
            cell = None
            if any_merged:  # поячеечная проверка объединений — только если они есть в диапазоне
                cell = ws.Cells(r1 + i, c1 + j)
                if bool(cell.MergeCells) and (int(cell.MergeArea.Row), int(cell.MergeArea.Column)) != (r1 + i, c1 + j):
                    result["skipped_merged"] += 1
                    continue
            maybe_formula = isinstance(formulas[i][j], str) and formulas[i][j].startswith("=")
            if not isinstance(row[j], str):
                result["skipped_formulas" if maybe_formula else "skipped_non_text"] += 1
                continue
            value, reason = _clean_text_value(row[j], operations, separator, bool(decimal_separator))
            address = a1_cell(r1 + i, c1 + j)
            if dates:
                cell = cell or ws.Cells(r1 + i, c1 + j)
                if bool(cell.HasFormula) or bool(cell.HasSpill):
                    result["skipped_formulas"] += 1
                    continue
                value = excel_date_serial(value, date_format, date1904)
                if value is None:
                    result["skipped_not_date"] += 1
                    if len(result["not_date"]) < 50:
                        result["not_date"].append(address)
                    continue
            if value == row[j]:
                if maybe_formula:
                    result["skipped_formulas"] += 1
                elif reason:
                    result["skipped_" + reason] += 1
                    if reason == "ambiguous":
                        result["ambiguous"].append(address)
                continue
            cell = cell or ws.Cells(r1 + i, c1 + j)
            # «=…» в массиве Formula бывает и у текста с апострофом; спрашиваем ячейку только у реальных изменений
            if (maybe_formula and bool(cell.HasFormula)) or bool(cell.HasSpill):  # HasSpill — часть растёкшегося массива
                result["skipped_formulas"] += 1
                continue
            if reason:
                result["skipped_" + reason] += 1
                if reason == "ambiguous":
                    result["ambiguous"].append(address)
            if isinstance(value, str) and len(value) > 32767:
                raise ToolError(f"Cleaned text in {address} exceeds 32767 characters; no cells were changed.")
            changes.append({"r": r1 + i, "c": c1 + j, "address": address, "cell": cell, "value": value,
                            "input": _clean_input(value, cell.PrefixCharacter), "format": cell.NumberFormat, "merged": bool(cell.MergeCells)})
            if len(result["examples"]) < 50:
                result["examples"].append({"cell": address, "before": row[j], "after": value})
    result["changed"] = len(changes)
    if preview or not changes:
        return result
    from .undo import require_undo

    require_undo("workbook", app, wb)
    groups = []
    for change in changes:
        previous = groups[-1][-1] if groups else None
        if previous and previous["c"] == change["c"] and previous["r"] + 1 == change["r"] and not (previous["merged"] or change["merged"]):
            groups[-1].append(change)
        else:
            groups.append([change])
    applied = 0
    try:
        for group in groups:
            for change in group:
                if dates or (not isinstance(change["value"], str) and change["format"] == "@"):
                    change["cell"].NumberFormat = date_style if dates else "General"
                    result["format_changed"].append(change["address"])
            target = sub_range(ws, group[0]["r"], group[0]["c"], group[-1]["r"], group[-1]["c"])
            target.Value2 = tuple((change["input"],) for change in group)
            for change in group:
                if not _clean_matches(change["cell"], change["value"]):
                    if isinstance(change["value"], str):
                        change["cell"].Value2 = "'" + change["value"]
                        result["rewritten_as_text"].append(change["address"])
                    if not _clean_matches(change["cell"], change["value"]):
                        raise ToolError(f"Read-back mismatch at {change['address']}.")
                applied += 1
                change["verified"] = True
    except (pywintypes.com_error, ToolError) as exc:
        confirmed = applied
        unknown = 0
        for change in changes:
            if change.get("verified"):
                continue
            try:
                confirmed += int(_clean_matches(change["cell"], change["value"]))
            except (pywintypes.com_error, ToolError):
                unknown += 1
        raise PartialChangeError(f"Clean text interrupted: applied={confirmed}, remaining={len(changes) - confirmed}, unverified={unknown}; "
                                 f"format_changed={result['format_changed']}. {com.translate(exc)}") from None
    return {**result, "applied": applied, "remaining": 0}


@office_tool("excel_core", "write", title="Split column", preview_arg="preview")
def excel_split_column(
    workbook: str, sheet: str, cells: str, delimiter: str = "", widths: list[int] | None = None,
    max_parts: int = 0, consecutive_as_one: bool = False, trim: bool = True, destination: str = "",
    types: list[str] | None = None, preview: bool = True,
) -> dict:
    """Split text constants in Python, preserving literal text; preview by default. Checks and snapshots the ENTIRE output rectangle before writing, with office_undo and partial-failure recovery.

    Args:
        workbook, sheet: exact workbook and source sheet; output stays on this sheet.
        cells: required single column; whole columns clipped to used range. Non-text constants and formulas/spills are refused; blank rows produce blanks.
        delimiter: comma, semicolon, tab, space, pipe or a literal nonempty string; exactly one of delimiter/widths is required. Quotes have no CSV escaping semantics.
        widths: positive character widths; an extra final part holds the remainder.
        max_parts: 0 uses the actual maximum; positive values keep the unsplit remainder in the final part.
        consecutive_as_one: treat adjacent delimiters as one.
        trim: strip whitespace from parts.
        destination: one A1 cell on the same sheet; empty overwrites the source in place. An explicit destination must be entirely empty, including the source.
        types: text/number for each part, omitted trailing types default to text. Ambiguous numbers and identifiers stay text.
        preview: true returns output range, counters and up to 20 examples without journal/history; false writes after a successful snapshot (unless undo disabled).

    Refuses protection, merges, tables, pivots, occupied output, sheet overflow and output over 100000 cells.
    """
    if bool(delimiter) == (widths is not None):
        raise ToolError("Pass exactly one nonempty delimiter or widths.")
    if widths is not None and (not isinstance(widths, list) or not widths or any(type(n) is not int or n < 1 for n in widths)):
        raise ToolError("widths must be a nonempty list of positive integers.")
    if type(max_parts) is not int or max_parts < 0:
        raise ToolError("max_parts must be a nonnegative integer.")
    if types is not None and (not isinstance(types, list) or any(t not in {"text", "number"} for t in types)):
        raise ToolError("types must contain only text or number.")
    sep = {"comma": ",", "semicolon": ";", "tab": "\t", "space": " ", "pipe": "|"}.get(delimiter, delimiter)
    app, wb = pick_workbook(workbook)
    # Check the un-clipped shape as well: A:B must not become a single used column.
    if not preview and wb.ReadOnly:
        raise ToolError("Workbook is read-only; no cells were changed.")
    _, original = get_range(wb, sheet, cells, empty_means_used=False)
    if int(original.Columns.Count) != 1:
        raise ToolError("cells must be exactly one column.")
    ws, source = bounded_range(app, wb, sheet, cells, 100000)
    if bool(ws.ProtectContents):
        raise ToolError("The sheet is protected; no cells were changed.")
    result = {"preview": preview, "workbook": wb.Name, "sheet": ws.Name, "parts": 0, "output_range": None,
              "rows": 0, "output_cells": 0, "converted_numbers": 0, "ambiguous_numbers": [], "examples": [],
              "rewritten_as_text": [], "format_changed": []}
    if source is None:
        return result
    if source.MergeCells is not False:
        raise ToolError("Source contains merged cells.")
    r1, c1, r2, _ = bounds(source)
    raw = source.Value2
    rows = raw if isinstance(raw, tuple) and raw and isinstance(raw[0], tuple) else ((raw,),)
    formulas, nontexts, parts = [], [], []
    for i, (value,) in enumerate(rows):
        cell = ws.Cells(r1 + i, c1)
        if bool(cell.HasFormula) or bool(cell.HasSpill):
            formulas.append(a1_cell(r1 + i, c1))
        elif value is not None and not isinstance(value, str):
            nontexts.append(a1_cell(r1 + i, c1))
        text = value if isinstance(value, str) else ""
        if widths is not None:
            cuts, start = [], 0
            for width in widths[:max_parts - 1 if max_parts else None]:
                cuts.append(text[start:start + width])
                start += width
            cuts.append(text[start:])
        else:
            pattern = "(?:" + re.escape(sep) + ")" + ("+" if consecutive_as_one else "")
            cuts = [text] if max_parts == 1 else re.split(pattern, text, maxsplit=max_parts - 1 if max_parts else 0)
        parts.append([p.strip() if trim else p for p in cuts])
    if formulas or nontexts:
        raise ToolError(f"Source formulas/spills: {formulas[:50]}; non-text constants: {nontexts[:50]}. No cells were changed.")
    count = max(map(len, parts))
    if len(types or []) > count:
        raise ToolError("types has more entries than output parts.")
    if destination:
        try:
            dr, dc, er, ec = parse_a1(destination)
        except ValueError:
            raise ToolError("destination must be one A1 cell on the source sheet.") from None
        if (dr, dc) != (er, ec):
            raise ToolError("destination must be one A1 cell.")
    else:
        dr, dc = r1, c1
    end, right = dr + len(parts) - 1, dc + count - 1
    if dr < 1 or dc < 1 or end > MAX_ROWS or right > MAX_COLS:
        raise ToolError("Output exceeds sheet boundaries.")
    if len(parts) * count > 100000:
        raise ToolError("Output exceeds 100000 cells.")
    output = sub_range(ws, dr, dc, end, right)
    if output.MergeCells is not False:
        raise ToolError("Output contains merged cells.")
    for collection, attr in ((ws.ListObjects, "Range"), (ws.PivotTables(), "TableRange2")):
        overlaps = [collection(i).Name for i in range(1, int(collection.Count) + 1)
                    if app.Intersect(output, getattr(collection(i), attr)) is not None]
        if overlaps:
            raise ToolError(f"Output intersects tables/pivots: {overlaps}.")
    occupied, changes = [], []
    for j in range(count):
        for i, row in enumerate(parts):
            cell = ws.Cells(dr + i, dc + j)
            address = a1_cell(dr + i, dc + j)
            in_source = not destination and dc + j == c1 and r1 <= dr + i <= r2
            if not in_source and (cell.Value2 is not None or bool(cell.HasFormula) or bool(cell.HasSpill)):
                occupied.append(address)
            value = row[j] if j < len(row) else ""
            if len(value) > 32767:
                raise ToolError("An output part exceeds 32767 characters.")
            if j < len(types or []) and types[j] == "number":
                value, reason = _clean_text_value(value, ["text_to_number"], app.International[2], False)
                if reason == "ambiguous":
                    result["ambiguous_numbers"].append(address)
                result["converted_numbers"] += int(not isinstance(value, str))
            changes.append({"cell": cell, "address": address, "value": value, "input": _clean_input(value, cell.PrefixCharacter),
                            "format": cell.NumberFormat})
    if occupied:
        raise ToolError(f"Output is not empty: {occupied[:50]}; no cells were changed.")
    result.update(parts=count, rows=len(parts), output_cells=len(changes), output_range=addr_of(output),
                  examples=[{"source": a1_cell(r1 + i, c1), "before": rows[i][0],
                             "after": [changes[j * len(parts) + i]["value"] for j in range(count)]} for i in range(min(20, len(parts)))])
    if preview:
        return result
    from .undo import require_undo

    require_undo("workbook", app, wb, plan={"_output_sheet": ws.Name, "_output_range": addr_of(output)})
    applied = 0
    try:
        for j in range(count):
            group = changes[j * len(parts):(j + 1) * len(parts)]
            for change in group:
                if not isinstance(change["value"], str) and change["format"] == "@":
                    change["cell"].NumberFormat = "General"
                    result["format_changed"].append(change["address"])
            sub_range(ws, dr, dc + j, end, dc + j).Value2 = tuple((change["input"],) for change in group)
            for change in group:
                if not _clean_matches(change["cell"], change["value"]):
                    if isinstance(change["value"], str):
                        change["cell"].Value2 = "'" + change["value"]
                        result["rewritten_as_text"].append(change["address"])
                    if not _clean_matches(change["cell"], change["value"]):
                        raise ToolError(f"Read-back mismatch at {change['address']}.")
                applied += 1
    except (pywintypes.com_error, ToolError) as exc:
        raise PartialChangeError(f"Split interrupted: applied={applied} verified cells, remaining={len(changes) - applied} may include partial writes; "
                                 f"format_changed={result['format_changed']}. {com.translate(exc)}") from None
    return {**result, "applied": applied, "remaining": 0}


def _assign_formula(rng, value, r1c1: bool = False):
    """Formula2 (динамические массивы, без неявного пересечения @), с откатом на Formula; r1c1 — FormulaR1C1."""
    if r1c1:
        try:
            rng.FormulaR1C1 = value
        except pywintypes.com_error as exc:
            raise ToolError(
                "Excel could not parse the R1C1 formula. Use English function names, commas, and R[-1]C / RC[2] references. " + com.com_error_text(exc)
            ) from None
        return
    try:
        rng.Formula2 = value
        return
    except (pywintypes.com_error, AttributeError):
        pass
    try:
        rng.Formula = value
    except pywintypes.com_error as exc:
        raise ToolError(
            "Excel could not parse the formula. Use English function names, comma separators, balanced parentheses, "
            "quoted sheet names ('My Sheet'!A1) and absolute refs where needed. " + com.com_error_text(exc)
        ) from None


def _prepare_target(ws, rng, rows: int, cols: int):
    """Однoячеечная цель разворачивается до размера данных; иначе форма должна совпасть."""
    r1, c1, r2, c2 = bounds(rng)
    cur_rows, cur_cols = r2 - r1 + 1, c2 - c1 + 1
    if (cur_rows, cur_cols) == (rows, cols):
        return rng, False
    if (cur_rows, cur_cols) == (1, 1):
        if r1 + rows - 1 > MAX_ROWS or c1 + cols - 1 > MAX_COLS:
            raise ToolError("Data does not fit on the sheet from that start cell.")
        return sub_range(ws, r1, c1, r1 + rows - 1, c1 + cols - 1), False
    if (rows, cols) == (1, 1):
        return rng, True  # скаляр заполняет весь диапазон
    raise ToolError(
        f"Range {addr_of(rng)} is {cur_rows}x{cur_cols} but the data is {rows}x{cols}. "
        "Pass just the top-left cell (e.g. 'A1') to fill the needed area automatically."
    )


@office_tool("excel_core", "write", title="Write values")
def excel_write_range(
    workbook: str,
    sheet: str,
    cells: str,
    values: list[list[Cell]] | list[Cell] | Cell,
    value_mode: str = "auto",
    overwrite: bool = True,
) -> dict:
    """Write a block of values. Pass a single top-left cell (e.g. 'B2') and the block is placed from there; rows shorter than the widest are padded with blanks; a single value fills the whole target range. Returns a read-back.

    In value_mode='auto' strings are interpreted like typed input ('123' becomes a number, '=A1+1' a formula, '2026-10-04' a date, '12%' a percentage). Use value_mode='text' to store every string literally as text (IDs, codes, phone numbers).

    Args:
        workbook: exact workbook name (required for safe targeting).
        sheet: sheet name ('' = active sheet).
        cells: top-left cell or full range.
        values: 2-D array of rows (a flat list is one row; a single value fills the whole target range). null leaves a cell blank.
        value_mode: 'auto' or 'text'.
        overwrite: false refuses to write when the target already holds non-empty cells.
    """
    if value_mode not in ("auto", "text"):
        raise ToolError("value_mode must be 'auto' or 'text'.")
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells, empty_means_used=False)
    grid = to_com_grid(values)
    rows, cols = len(grid), len(grid[0])
    for row in grid:
        for v in row:
            if isinstance(v, str) and len(v) > 32767:
                raise ToolError("A cell value is longer than 32767 characters.")
    target, fill = _prepare_target(ws, rng, rows, cols)
    before = count_nonempty(app, target)
    if before > 0 and not overwrite:
        raise ToolError(f"Target {addr_of(target)} already contains {before} non-empty cells; pass overwrite=true to replace them.")
    if value_mode == "text":
        grid = tuple(tuple(("'" + v) if isinstance(v, str) and v != "" else v for v in row) for row in grid)
    try:
        target.Value = grid[0][0] if fill else grid
    except pywintypes.com_error as exc:
        raise ToolError("Excel refused the write (protected sheet, merged cells in the way, or invalid data): " + com.com_error_text(exc)) from None
    r1, c1, _, _ = bounds(target)
    back = preview(target)
    types: dict[str, int] = {}
    for row in back:
        for v in row:
            k = cell_kind(v)
            if k != "empty":
                types[k] = types.get(k, 0) + 1
    return {
        "ok": True, "workbook": wb.Name, "sheet": ws.Name, "cells": addr_of(target),
        "written_shape": [rows, cols], "overwritten_nonempty_cells": max(before, 0),
        "readback": back, "readback_types": types, "errors": error_cells(back, r1, c1),
    }


@office_tool("excel_core", "write", title="Set formulas")
def excel_set_formula(workbook: str, sheet: str, cells: str, formula: str | list[list[str]], r1c1: bool = False) -> dict:
    """Put formulas into a cell or fill a whole range. A single formula string assigned to a multi-cell range is filled like Excel's fill-down: relative references adjust per cell (write '=A2*B2' with cells='C2:C100'). A 2-D array assigns each cell its own formula. Dynamic-array formulas (SORT, FILTER, UNIQUE, SEQUENCE) spill normally. Returns computed values and any error cells.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        cells: target cell or range.
        formula: '=SUM(A1:A10)' for the whole range, or a 2-D array of formula strings matching the range shape. In A1 mode write it for the TOP-LEFT cell of `cells`.
        r1c1: true = the formula is in R1C1 notation ('=RC[-1]*2', '=SUM(R[-3]C:R[-1]C)'), position independent: the same text is right for every cell. Use it to copy the formula_r1c1 reported by excel_describe_layout.
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells, empty_means_used=False)
    if isinstance(formula, str):
        if not formula.startswith("="):
            raise ToolError("Formula must start with '='.")
        _assign_formula(rng, formula, r1c1)
        target = rng
    else:
        grid = to_com_grid(formula)
        for row in grid:
            for f in row:
                if not isinstance(f, str) or not f.startswith("="):
                    raise ToolError("Every entry of the formula array must be a string starting with '='.")
        target, _ = _prepare_target(ws, rng, len(grid), len(grid[0]))
        _assign_formula(target, grid, r1c1)
    r1, c1, _, _ = bounds(target)
    back = preview(target)
    first = a1_cell(r1, c1)
    return {
        "ok": True, "workbook": wb.Name, "sheet": ws.Name, "cells": addr_of(target),
        "first_formula": ws.Range(first).Formula, "values": back, "errors": error_cells(back, r1, c1),
    }


@office_tool("excel_core", "destructive", title="Clear range")
def excel_clear_range(workbook: str, sheet: str, cells: str, what: str = "contents") -> dict:
    """Clear parts of a range.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        cells: range to clear.
        what: 'contents' (values and formulas, keeps formatting), 'formats', 'all', 'comments', or 'hyperlinks'.
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells, empty_means_used=False)
    filled = max(count_nonempty(app, rng), 0)
    w = what.lower()
    if w == "contents":
        rng.ClearContents()
    elif w == "formats":
        rng.ClearFormats()
    elif w == "all":
        rng.Clear()
    elif w == "comments":
        rng.ClearComments()
    elif w == "hyperlinks":
        rng.ClearHyperlinks()
    else:
        raise ToolError("what must be one of: contents, formats, all, comments, hyperlinks")
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cleared": addr_of(rng), "what": w, "nonempty_cells_before": filled}


@office_tool("excel_core", "write", title="Find and replace")
def excel_replace(
    find: str,
    replace: str,
    workbook: str,
    sheet: str = "",
    cells: str = "",
    match_case: bool = False,
    whole_cell: bool = False,
) -> dict:
    """Replace text inside cell contents/formulas (keeps formatting). Returns how many cells and occurrences were affected.

    Args:
        find: text to look for.
        replace: replacement text (can be empty).
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        cells: restrict to a range ('' = the whole used range of the sheet).
        match_case: case-sensitive search.
        whole_cell: match only cells whose entire content equals `find`.
    """
    if not find:
        raise ToolError("'find' is empty.")
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells)
    rng = clip_to_used(app, ws, rng)
    if rng is None or sheet_is_empty(app, ws):
        return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cells_changed": 0, "occurrences": 0}
    before = to_grid(rng.Formula)
    flags = 0 if match_case else re.IGNORECASE
    pat = re.compile(re.escape(find), flags)
    cells_hit = occ = 0
    for row in before:
        for v in row:
            if v is None:
                continue
            text = _excel_text(v)
            n = (1 if pat.fullmatch(text) else 0) if whole_cell else len(pat.findall(text))
            if n:
                cells_hit += 1
                occ += n
    changed = 0
    if occ:
        # Replace(What, Replacement, LookAt: 1=whole 2=part, SearchOrder: 1=rows, MatchCase, MatchByte, SearchFormat, ReplaceFormat).
        # Символы * ? ~ в What — шаблоны Excel, а нужен буквальный поиск; форматы поиска/замены из диалога пользователя сбрасываем.
        rng.Replace(_wildcard_escape(find), replace, 1 if whole_cell else 2, 1, bool(match_case), False, False, False)
        after = to_grid(rng.Formula)
        changed = sum(1 for r0, r1 in zip(before, after, strict=False) for x, y in zip(r0, r1, strict=False) if x != y)
    out = {
        "ok": True, "workbook": wb.Name, "sheet": ws.Name, "scope": addr_of(rng),
        "cells_matching": cells_hit, "occurrences": occ, "cells_changed": changed,
    }
    if changed != cells_hit:
        out["warning"] = (
            f"{cells_hit} cell(s) matched but {changed} actually changed (locked cells, a replacement equal to the original "
            "text, or Excel interpreted the text differently). Read the range to check."
        )
    return out


def _excel_text(v) -> str:
    """Как Excel видит значение ячейки при поиске: целое число без «.0», логические — TRUE/FALSE."""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return v if isinstance(v, str) else str(v)


def _wildcard_escape(text: str) -> str:
    return text.replace("~", "~~").replace("*", "~*").replace("?", "~?")


# ================================================================== листы


def _position_args(wb, position: str):
    """position: end | start | before:<лист> | after:<лист> -> позиционные аргументы (Before, After)."""
    p = (position or "end").strip()
    low = p.lower()
    if low == "end":
        return None, wb.Sheets(wb.Sheets.Count)
    if low == "start":
        return wb.Sheets(1), None
    if low.startswith("before:"):
        return pick_sheet(wb, p.split(":", 1)[1].strip(), any_type=True), None
    if low.startswith("after:"):
        return None, pick_sheet(wb, p.split(":", 1)[1].strip(), any_type=True)
    raise ToolError("position must be 'end', 'start', 'before:<sheet>' or 'after:<sheet>'")


@office_tool("excel_core", "write", title="Add worksheet")
def excel_add_worksheet(workbook: str, name: str = "", position: str = "end", activate: bool = False) -> dict:
    """Add a worksheet. Fails if the name is taken. By default the user's active sheet is left unchanged.

    Args:
        workbook: exact workbook name.
        name: new sheet name (max 31 chars, none of []:*?/\\); '' keeps Excel's default.
        position: 'end' (default), 'start', 'before:<sheet>' or 'after:<sheet>'.
        activate: make the new sheet the active one.
    """
    app, wb = pick_workbook(workbook)
    if name:
        validate_sheet_name(name)
        if name.lower() in [n.lower() for n in _sheet_list(wb)]:
            raise ToolError(f"Sheet '{name}' already exists in {wb.Name}.")
    prev = wb.ActiveSheet
    before, after = _position_args(wb, position)
    ws = wb.Worksheets.Add(before, after)  # Add(Before, After) — только позиционно
    if name:
        ws.Name = name
    if not activate and prev is not None:
        try:
            prev.Activate()
        except pywintypes.com_error:
            pass
    return {"ok": True, "workbook": wb.Name, "added_sheet": ws.Name, "index": int(ws.Index), "sheets": _sheet_list(wb)}


@office_tool("excel_core", "write", title="Manage sheet", destructive=True)
def excel_manage_sheet(
    workbook: str,
    sheet: str,
    action: str,
    new_name: str = "",
    position: str = "",
    color: str = "",
    dest_workbook: str = "",
) -> dict:
    """Rename, copy, move, hide/unhide, color or activate a sheet.

    Args:
        workbook: exact workbook name.
        sheet: the sheet to act on.
        action: 'rename' (needs new_name), 'copy' (optional new_name and position), 'move' (needs position), 'hide', 'very_hide', 'unhide', 'tab_color' (needs color, 'none' removes it), 'activate'.
        new_name: new/copy name.
        position: 'end', 'start', 'before:<sheet>' or 'after:<sheet>' for copy/move.
        color: '#RRGGBB' or color name for tab_color.
        dest_workbook: for 'copy' - copy the sheet into ANOTHER open workbook (formats, merged cells, widths, conditional formats, charts come along).
    """
    app, wb = pick_workbook(workbook)
    ws = pick_sheet(wb, sheet, any_type=True)
    act = action.lower()
    if act == "rename":
        validate_sheet_name(new_name)
        ws.Name = new_name
    elif act == "copy":
        if new_name:
            validate_sheet_name(new_name)
        if dest_workbook:
            _, dwb = pick_workbook(dest_workbook)
            pos = (position or "end").lower()
            before, after = _position_args(dwb, pos)
            count = int(dwb.Sheets.Count)
            if pos == "start":
                new_index = 1
            elif pos.startswith("before:"):
                new_index = int(before.Index)
            elif pos.startswith("after:"):
                new_index = int(after.Index) + 1
            else:
                new_index = count + 1
            ws.Copy(before, after)
            copy = dwb.Sheets(new_index)
            if new_name:
                copy.Name = new_name
            try:
                ws.Activate()
            except pywintypes.com_error:
                pass
            return {"ok": True, "workbook": wb.Name, "action": act, "copied_to": dwb.Name, "sheet": copy.Name, "sheets": _sheet_list(dwb)}
        before, after = _position_args(wb, position or "after:" + ws.Name)
        ws.Copy(before, after)  # Copy(Before, After)
        copy = wb.ActiveSheet  # копия становится активной
        if new_name:
            copy.Name = new_name
        ws.Activate()
    elif act == "move":
        if not position:
            raise ToolError("'position' is required for move.")
        before, after = _position_args(wb, position)
        ws.Move(before, after)
    elif act in ("hide", "very_hide", "unhide"):
        if act != "unhide":
            visible = [i for i in range(1, wb.Sheets.Count + 1) if int(wb.Sheets(i).Visible) == -1]
            if len(visible) <= 1 and int(ws.Visible) == -1:
                raise ToolError("Cannot hide the last visible sheet.")
        ws.Visible = {"hide": 0, "very_hide": 2, "unhide": -1}[act]
    elif act == "tab_color":
        if not color:
            raise ToolError("'color' is required for tab_color.")
        if color.lower() == "none":
            ws.Tab.ColorIndex = -4142
        else:
            ws.Tab.Color = parse_color(color)
    elif act == "activate":
        wb.Activate()
        ws.Activate()
    else:
        raise ToolError("action must be one of: rename, copy, move, hide, very_hide, unhide, tab_color, activate")
    return {"ok": True, "workbook": wb.Name, "action": act, "sheets": _sheet_list(wb)}


@office_tool("excel_core", "destructive", title="Delete worksheet")
def excel_delete_sheet(workbook: str, sheet: str, confirm: bool = False) -> dict:
    """Permanently delete a sheet with all its data. Requires confirm=true. Refuses to delete the last sheet.

    Args:
        workbook: exact workbook name.
        sheet: sheet to delete.
        confirm: must be true.
    """
    if not confirm:
        raise ToolError("Deleting a sheet is irreversible. Ask the user, then call again with confirm=true.")
    app, wb = pick_workbook(workbook)
    ws = pick_sheet(wb, sheet, any_type=True)
    if wb.Sheets.Count <= 1:
        raise ToolError("Cannot delete the only sheet of a workbook.")
    name = ws.Name
    before = app.DisplayAlerts
    app.DisplayAlerts = False
    try:
        ws.Delete()
    finally:
        _restore_alerts(app, before)
    return {"ok": True, "deleted": name, "sheets": _sheet_list(wb)}


# ================================================================== структура


def _rows_cols_range(ws, axis: str, start: int, count: int):
    if axis not in ("rows", "columns"):
        raise ToolError("axis must be 'rows' or 'columns'.")
    if start < 1 or count < 1:
        raise ToolError("start and count must be >= 1.")
    if axis == "rows":
        return ws.Range(f"{start}:{start + count - 1}")
    return ws.Range(f"{col_letter(start)}:{col_letter(start + count - 1)}")


@office_tool("excel_core", "write", title="Insert rows/columns")
def excel_insert_rows_columns(workbook: str, sheet: str, axis: str, start: int, count: int = 1) -> dict:
    """Insert empty rows or columns; existing content shifts down/right.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        axis: 'rows' or 'columns'.
        start: 1-based row number / column number where the new ones are inserted (before it).
        count: how many to insert.
    """
    app, wb = pick_workbook(workbook)
    ws = pick_sheet(wb, sheet)
    _rows_cols_range(ws, axis, int(start), int(count)).Insert()
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "inserted": axis, "start": int(start), "count": int(count)}


@office_tool("excel_core", "destructive", title="Delete rows/columns")
def excel_delete_rows_columns(workbook: str, sheet: str, axis: str, start: int, count: int = 1) -> dict:
    """Delete rows or columns together with their content; the rest shifts up/left. References to deleted cells become #REF!.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        axis: 'rows' or 'columns'.
        start: 1-based first row/column to delete.
        count: how many to delete.
    """
    app, wb = pick_workbook(workbook)
    ws = pick_sheet(wb, sheet)
    rng = _rows_cols_range(ws, axis, int(start), int(count))
    lost = max(count_nonempty(app, rng), 0)
    rng.Delete()
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "deleted": axis, "start": int(start), "count": int(count), "nonempty_cells_lost": lost}


@office_tool("excel_core", "write", title="Column width / row height")
def excel_set_dimensions(
    workbook: str,
    sheet: str,
    cells: str = "",
    column_width: float | None = None,
    row_height: float | None = None,
    autofit: str = "",
) -> dict:
    """Set column width / row height or autofit them (to hide rows/columns use excel_hide_rows_columns).

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        cells: range whose columns/rows are affected ('' = the whole sheet for autofit).
        column_width: width in characters (e.g. 12).
        row_height: height in points (e.g. 18).
        autofit: 'columns', 'rows' or 'both' - fit to content.
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells)
    if not cells:
        rng = ws.Cells
    applied = []
    if column_width is not None:
        rng.ColumnWidth = float(column_width)
        applied.append(f"column_width={column_width}")
    if row_height is not None:
        rng.RowHeight = float(row_height)
        applied.append(f"row_height={row_height}")
    if autofit:
        mode = autofit.lower()
        if mode in ("columns", "both"):
            rng.Columns.AutoFit()
        if mode in ("rows", "both"):
            rng.Rows.AutoFit()
        if mode not in ("columns", "rows", "both"):
            raise ToolError("autofit must be 'columns', 'rows' or 'both'.")
        applied.append(f"autofit={mode}")
    if not applied:
        raise ToolError("Nothing to do: pass column_width, row_height or autofit.")
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cells": addr_of(rng) if cells else "entire sheet", "applied": applied}


def _runs(numbers: list[int]) -> list[tuple[int, int]]:
    out = []
    for n in sorted(set(numbers)):
        if out and n == out[-1][1] + 1:
            out[-1] = (out[-1][0], n)
        else:
            out.append((n, n))
    return out


def _line_spec(spec: str, axis: str) -> list[tuple[int, int]]:
    """'5-10,15' (строки) или 'C:E,H' / '3-5' (столбцы) -> [(first, last)]."""
    out = []
    for part in spec.replace(" ", "").split(","):
        if not part:
            continue
        sep = ":" if ":" in part else "-"
        a, _, b = part.partition(sep)
        b = b or a

        def num(x):
            return int(x) if x.isdigit() else col_number(x)

        lo, hi = sorted((num(a), num(b)))
        out.append((lo, hi))
    if not out:
        raise ToolError("Empty row/column specification.")
    return out


def _where_match(v, op: str, target) -> bool:
    blank = v is None or v == ""
    if op == "blank":
        return blank
    if op == "not_blank":
        return not blank
    s = "" if v is None else str(v)
    t = "" if target is None else str(target)
    if op == "equals":
        return s.lower() == t.lower()
    if op == "not_equals":
        return s.lower() != t.lower()
    if op == "contains":
        return t.lower() in s.lower()
    if op == "not_contains":
        return t.lower() not in s.lower()
    if op == "in":
        return s.lower() in {str(x).lower() for x in (target if isinstance(target, list) else [target])}
    try:
        a, b = float(v), float(target)
    except (TypeError, ValueError):
        return False
    return {"greater": a > b, "greater_equal": a >= b, "less": a < b, "less_equal": a <= b}.get(op, False)


@office_tool("excel_core", "write", title="Hide / show rows and columns")
def excel_hide_rows_columns(
    workbook: str,
    sheet: str = "",
    axis: str = "rows",
    lines: str = "",
    hidden: bool = True,
    where: dict | None = None,
    cells: str = "",
    has_header: bool = True,
) -> dict:
    """Hide or unhide whole rows/columns - either by number or by a condition on the data ('hide every row whose Status is Done'). Hidden rows keep their data; unhide with hidden=false.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        axis: 'rows' or 'columns'.
        lines: rows like '5-10,15' or columns like 'C:E,H' (or numbers '3-5'). Use `lines` OR `where`.
        hidden: true = hide, false = show.
        where: condition (rows only), e.g. {"column": "Status", "operator": "equals", "value": "Done"}. column = header text or letter; operator = equals | not_equals | contains | not_contains | blank | not_blank | greater | greater_equal | less | less_equal | in (value = list).
        cells: data block for `where` (default: the whole used range).
        has_header: the block's first row is a header (never hidden by `where`; used to look up the column name).
    """
    app, wb = pick_workbook(workbook)
    ax = axis.lower()
    if ax not in ("rows", "columns"):
        raise ToolError("axis must be 'rows' or 'columns'.")
    if bool(lines) == bool(where):
        raise ToolError("Pass exactly one of `lines` or `where`.")
    if where:
        if ax != "rows":
            raise ToolError("`where` works only with axis='rows'.")
        ws, rng = get_range(wb, sheet, cells)  # лист берём из ссылки 'Лист!A1:F9', а не из параметра sheet
        rng = clip_to_used(app, ws, rng)
        if rng is None:
            return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "count": 0}
        r1, c1, r2, c2 = bounds(rng)
        grid = to_grid(rng.Value)
        header = grid[0] if has_header else []
        col = _resolve_column(ws, rng, where.get("column"), header, has_header)
        op = str(where.get("operator", "equals")).lower()
        if op not in ("equals", "not_equals", "contains", "not_contains", "blank", "not_blank", "greater", "greater_equal", "less", "less_equal", "in"):
            raise ToolError("Unknown operator in `where`.")
        start = 1 if has_header else 0
        rows = [r1 + i for i in range(start, len(grid)) if _where_match(grid[i][col - 1], op, where.get("value"))]
        spans = _runs(rows)
    else:
        ws = pick_sheet(wb, sheet)
        spans = _line_spec(lines, ax)
    for a, b in spans:
        target = ws.Range(f"{a}:{b}") if ax == "rows" else ws.Range(f"{col_letter(a)}:{col_letter(b)}")
        (target.EntireRow if ax == "rows" else target.EntireColumn).Hidden = bool(hidden)
    desc = [f"{a}:{b}" if ax == "rows" else f"{col_letter(a)}:{col_letter(b)}" for a, b in spans]
    return {
        "ok": True, "workbook": wb.Name, "sheet": ws.Name, "axis": ax, "hidden": bool(hidden),
        "affected": desc[:100], "count": sum(b - a + 1 for a, b in spans),
    }


@office_tool("excel_core", "write", title="Group rows/columns (outline)")
def excel_group_rows_columns(workbook: str, sheet: str = "", axis: str = "rows", lines: str = "", action: str = "group") -> dict:
    """Create or remove outline groups (the +/- buttons in the margin) over rows or columns, or collapse/expand every group on the sheet.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        axis: 'rows' or 'columns'.
        lines: rows like '5-10' or columns like 'C:E' (for group/ungroup).
        action: 'group', 'ungroup', 'collapse' (show only level 1) or 'expand' (show all levels).
    """
    app, wb = pick_workbook(workbook)
    ws = pick_sheet(wb, sheet)
    ax = axis.lower()
    if ax not in ("rows", "columns"):
        raise ToolError("axis must be 'rows' or 'columns'.")
    act = action.lower()
    if act in ("group", "ungroup"):
        for a, b in _line_spec(lines, ax):
            target = ws.Range(f"{a}:{b}") if ax == "rows" else ws.Range(f"{col_letter(a)}:{col_letter(b)}")
            if act == "group":
                target.Group()
            else:
                target.Ungroup()
    elif act in ("collapse", "expand"):
        level = 1 if act == "collapse" else 8
        if ax == "rows":
            ws.Outline.ShowLevels(level, 0)
        else:
            ws.Outline.ShowLevels(0, level)
    else:
        raise ToolError("action must be 'group', 'ungroup', 'collapse' or 'expand'.")
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "axis": ax, "action": act}


@office_tool("excel_core", "write", title="Copy or move range")
def excel_copy_range(
    workbook: str,
    sheet: str,
    source: str,
    dest_cell: str,
    dest_sheet: str = "",
    dest_workbook: str = "",
    what: str = "all",
    move: bool = False,
    transpose: bool = False,
) -> dict:
    """Copy (or move) a range to another place, optionally into another sheet or workbook.

    Args:
        workbook: workbook that holds the source.
        sheet: source sheet ('' = active).
        source: source range.
        dest_cell: top-left destination cell; for what='all'/'formats' it may be a larger range (e.g. 'A20:AT200') that the source block is tiled over - use it to stamp a template row's look onto many rows.
        dest_sheet: destination sheet ('' = same sheet).
        dest_workbook: destination workbook ('' = same workbook).
        what: 'all' (values, formulas, formats), 'values', 'formulas' (as typed), or 'formats' (uses the clipboard briefly).
        move: cut instead of copy (only with what='all').
        transpose: swap rows and columns (what='values' or 'formulas' only).
    """
    app, wb = pick_workbook(workbook)
    ws, src = get_range(wb, sheet, source, empty_means_used=False)
    dapp, dwb = pick_workbook(dest_workbook) if dest_workbook else (app, wb)
    dws = pick_sheet(dwb, dest_sheet) if dest_sheet else (ws if dwb is wb else pick_sheet(dwb, ""))
    r1, c1, r2, c2 = bounds(src)
    rows, cols = r2 - r1 + 1, c2 - c1 + 1
    d_r1, d_c1, _, _ = parse_a1(dest_cell)
    w = what.lower()
    if transpose:
        rows, cols = cols, rows
    dest = sub_range(dws, d_r1, d_c1, d_r1 + rows - 1, d_c1 + cols - 1)
    dest_head = dws.Range(dest_cell.replace("$", "")) if ":" in dest_cell else dws.Range(a1_cell(d_r1, d_c1))  # диапазон-цель: Excel тиражирует образец
    if w == "all":
        if transpose:
            raise ToolError("transpose works only with what='values' or 'formulas'.")
        if move:
            src.Cut(dest_head)
        else:
            src.Copy(dest_head)
            app.CutCopyMode = False
    elif w in ("values", "formulas"):
        if move:
            raise ToolError("move works only with what='all'.")
        data = src.Value if w == "values" else src.Formula
        if transpose:
            rows_data = [list(r) for r in data] if isinstance(data, tuple) and data and isinstance(data[0], tuple) else ([list(data)] if isinstance(data, tuple) else [[data]])
            data = tuple(tuple(r) for r in zip(*rows_data, strict=False))
        if w == "values":
            dest.Value = data
        else:
            dest.Formula = data
    elif w == "formats":
        src.Copy()
        try:
            dest_head.PasteSpecial(-4122, -4142, False, False)  # xlPasteFormats
        finally:
            app.CutCopyMode = False
    else:
        raise ToolError("what must be one of: all, values, formulas, formats")
    actual_dest = dest_head if w in {"all", "formats"} and int(dest_head.Rows.Count) * int(dest_head.Columns.Count) > 1 else dest
    return {"ok": True, "from": ref_label(ws, src), "to": ref_label(dws, actual_dest), "what": w, "moved": bool(move),
            "source_workbook": wb.FullName, "dest_workbook": dwb.FullName}


def _resolve_column(ws, rng, key, header_names: list, has_header: bool) -> int:
    """Ключ сортировки/фильтра -> номер столбца (1-based) внутри диапазона."""
    r1, c1, r2, c2 = bounds(rng)
    width = c2 - c1 + 1
    if isinstance(key, int) or (isinstance(key, str) and key.strip().isdigit()):
        idx = int(key)
        if not 1 <= idx <= width:
            raise ToolError(f"Column index {idx} is outside the range (1..{width}).")
        return idx
    k = str(key).strip()
    if has_header:
        for i, h in enumerate(header_names, start=1):
            if h is not None and str(h).strip().lower() == k.lower():
                return i
    if re.fullmatch(r"[A-Za-z]{1,3}", k):
        idx = col_number(k) - c1 + 1
        if 1 <= idx <= width:
            return idx
    raise ToolError(f"Column '{key}' not found. Headers: {header_names if has_header else '(no header row)'}; range columns: {col_letter(c1)}..{col_letter(c2)}")


@office_tool("excel_core", "write", title="Sort range")
def excel_sort_range(
    workbook: str,
    sheet: str,
    cells: str,
    keys: list[dict],
    has_header: bool = True,
) -> dict:
    """Sort a block of data by one or more columns (rows move as a whole).

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        cells: the data block, including the header row when has_header=true.
        keys: ordered list like [{"column": "Revenue", "order": "desc"}, {"column": "B"}]. column = header text, a column letter, or a 1-based index inside the block; order = 'asc' (default) or 'desc'.
        has_header: first row of the block is a header and is not sorted.
    """
    if not keys:
        raise ToolError("'keys' must contain at least one sort key.")
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells, empty_means_used=False)
    r1, c1, r2, c2 = bounds(rng)
    header_names = to_grid(sub_range(ws, r1, c1, r1, c2).Value)[0] if has_header else []
    data_r1 = r1 + (1 if has_header else 0)
    if data_r1 > r2:
        raise ToolError("The block has no data rows to sort.")
    plan = []
    for k in keys:  # все ключи проверяем ДО очистки прежних полей сортировки
        col = _resolve_column(ws, rng, k.get("column"), header_names, has_header)
        order = str(k.get("order", "asc")).lower()
        if order not in ("asc", "desc"):
            raise ToolError("order must be 'asc' or 'desc'.")
        plan.append((col, order))
    sort = ws.Sort
    sort.SortFields.Clear()
    used = []
    for col, order in plan:
        key_range = sub_range(ws, data_r1, c1 + col - 1, r2, c1 + col - 1)
        sort.SortFields.Add(key_range, 0, 1 if order == "asc" else 2)  # Add(Key, SortOn=values, Order)
        used.append({"column": col_letter(c1 + col - 1), "order": order})
    sort.SetRange(rng)
    sort.Header = 1 if has_header else 2
    sort.MatchCase = False
    sort.Orientation = 1
    sort.Apply()
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "sorted": addr_of(rng), "keys": used}


_FILTER_OPS = {"and": 1, "or": 2, "top_items": 3, "bottom_items": 4, "top_percent": 5, "bottom_percent": 6, "values": 7}


@office_tool("excel_core", "write", title="Filter range")
def excel_filter_range(
    workbook: str,
    sheet: str,
    cells: str = "",
    filters: list[dict] | None = None,
    clear: str = "",
    has_header: bool = True,
) -> dict:
    """Apply AutoFilter conditions (rows are hidden, not deleted) or clear them.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        cells: data block including the header row (needed to apply filters).
        filters: list of {"column": header/letter/index, "criteria": ">100" | "=Moscow" | "<>" | "=*abc*" | ["a","b"] (list of exact values), "operator": optional 'and'/'or'/'top_items'/'bottom_items'/'top_percent'/'bottom_percent', "criteria2": optional second condition}.
        clear: 'show_all' re-shows every row but keeps the filter arrows; 'remove' deletes the filter completely.
        has_header: the block's first row holds headers.
    """
    app, wb = pick_workbook(workbook)
    if clear:
        ws = get_range(wb, sheet, cells, empty_means_used=False)[0] if cells else pick_sheet(wb, sheet)
        c = clear.lower()
        if c == "show_all":
            if ws.AutoFilterMode and ws.FilterMode:
                ws.ShowAllData()
        elif c == "remove":
            if ws.AutoFilterMode:
                ws.AutoFilterMode = False
        else:
            raise ToolError("clear must be 'show_all' or 'remove'.")
        return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cleared": c}
    if not filters:
        raise ToolError("Pass 'filters' to apply or 'clear' to remove.")
    ws, rng = get_range(wb, sheet, cells, empty_means_used=False)
    r1, c1, r2, c2 = bounds(rng)
    header_names = to_grid(sub_range(ws, r1, c1, r1, c2).Value)[0] if has_header else []
    prepared = []
    for f in filters:  # все условия разбираем и проверяем ДО снятия прежнего фильтра
        col = _resolve_column(ws, rng, f.get("column"), header_names, has_header)
        crit = f.get("criteria")
        op = f.get("operator")
        op_code = _FILTER_OPS.get(str(op).lower()) if op else None
        if op and op_code is None:
            raise ToolError(f"Unknown operator {op!r}. Use {sorted(_FILTER_OPS)}.")
        if isinstance(crit, list):
            op_code = 7
            crit = tuple(str(x) for x in crit)
        elif crit is not None:
            crit = str(crit)
        crit2 = f.get("criteria2")
        prepared.append((col, crit, op_code or 1, None if crit2 is None else str(crit2), f.get("criteria")))  # Operator типизирован: None нельзя, по умолчанию xlAnd=1
    if ws.AutoFilterMode:
        ws.AutoFilterMode = False
    applied = []
    for col, crit, op_code, crit2, shown in prepared:
        try:
            rng.AutoFilter(col, crit, op_code, crit2)  # AutoFilter(Field, Criteria1, Operator, Criteria2)
        except pywintypes.com_error as exc:
            raise ToolError(
                f"Excel rejected the filter on column {col_letter(c1 + col - 1)}: {com.com_error_text(exc)}. The previous filter had been removed"
                + (f"; already applied before the failure: {applied}" if applied else "") + "."
            ) from None
        applied.append({"column": col_letter(c1 + col - 1), "criteria": shown})
    visible = 0
    try:
        body = sub_range(ws, r1 + (1 if has_header else 0), c1, r2, c2)
        for area in body.SpecialCells(12).Areas:  # xlCellTypeVisible
            visible += int(area.Rows.Count)
    except pywintypes.com_error:
        visible = 0
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "filtered": addr_of(rng), "applied": applied, "visible_data_rows": visible, "total_data_rows": r2 - r1 + (0 if has_header else 1)}


@office_tool("excel_core", "destructive", title="Remove duplicate rows")
def excel_remove_duplicates(workbook: str, sheet: str, cells: str, columns: list[str | int] | None = None, has_header: bool = True) -> dict:
    """Delete duplicate rows from a block of data, keeping the first occurrence. Comparison is case-insensitive like Excel's own command.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        cells: data block including the header row when has_header=true.
        columns: which columns define a duplicate (header text, letter or 1-based index); default: all columns.
        has_header: first row is a header.
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells, empty_means_used=False)
    r1, c1, r2, c2 = bounds(rng)
    width = c2 - c1 + 1
    grid = to_grid(rng.Value)
    header_names = grid[0] if has_header else []
    cols = [_resolve_column(ws, rng, c, header_names, has_header) for c in columns] if columns else list(range(1, width + 1))
    seen, dupes = set(), 0
    for row in grid[1 if has_header else 0:]:
        key = tuple(str(row[c - 1]).lower() if row[c - 1] is not None else None for c in cols)
        if key in seen:
            dupes += 1
        else:
            seen.add(key)
    rng.RemoveDuplicates(tuple(cols), 1 if has_header else 2)  # RemoveDuplicates(Columns, Header)
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "block": addr_of(rng), "duplicate_rows_removed": dupes, "key_columns": [col_letter(c1 + c - 1) for c in cols]}


@office_tool("excel_core", "write", title="Merge / unmerge cells", destructive=True)
def excel_merge_cells(workbook: str, sheet: str, cells: str, unmerge: bool = False, across: bool = False) -> dict:
    """Merge a range into one cell (only the top-left value survives) or unmerge it.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        cells: range to merge/unmerge.
        unmerge: split merged cells back.
        across: merge each row separately instead of the whole block.
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells, empty_means_used=False)
    before = app.DisplayAlerts
    app.DisplayAlerts = False
    try:
        if unmerge:
            rng.UnMerge()
        else:
            rng.Merge(bool(across))
    finally:
        _restore_alerts(app, before)
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cells": addr_of(rng), "action": "unmerge" if unmerge else "merge"}


@office_tool("excel_core", "write", title="Defined names", read_actions=("list",), destructive=True)
def excel_manage_names(workbook: str, action: str = "list", name: str = "", sheet: str = "", cells: str = "", formula: str = "") -> dict:
    """List, add or delete defined names (named ranges / named formulas).

    Args:
        workbook: exact workbook name.
        action: 'list', 'add' or 'delete'.
        name: the name (letters, digits, underscore; not a cell address).
        sheet: sheet of the range for 'add'.
        cells: range for 'add' (e.g. 'B2:B50'); alternatively pass `formula`.
        formula: for 'add' a refers-to formula like '=Sheet1!$A$1:$A$10' or '=0.2' (a named constant).
    """
    app, wb = pick_workbook(workbook)
    act = action.lower()
    if act == "list":
        items = []
        for i in range(1, wb.Names.Count + 1):
            nm = wb.Names(i)
            if bool(nm.Visible):
                items.append({"name": nm.Name, "refers_to": nm.RefersTo})
        return {"workbook": wb.Name, "names": items}
    if not name:
        raise ToolError("'name' is required.")
    if act == "add":
        if formula:
            refers = formula if formula.startswith("=") else "=" + formula
        elif cells:
            ws, rng = get_range(wb, sheet, cells, empty_means_used=False)
            r1, c1, r2, c2 = bounds(rng)
            refers = f"={quote_sheet(ws.Name)}!${col_letter(c1)}${r1}" + (f":${col_letter(c2)}${r2}" if (r1, c1) != (r2, c2) else "")
        else:
            raise ToolError("Pass `cells` (with `sheet`) or `formula` for 'add'.")
        # Names.Add через позднее связывание разбирает формулу по ЛОКАЛЬНЫМ правилам (в русском Excel '=SUM(A1,B1)'
        # и '=0.2' — ошибка), а RefersTo читается по-английски: переводим, как для условного форматирования.
        wb.Names.Add(name, localize_formula(app, wb, refers))
        return {"ok": True, "workbook": wb.Name, "added": name, "refers_to": str(wb.Names(name).RefersTo)}
    if act == "delete":
        wb.Names(name).Delete()
        return {"ok": True, "workbook": wb.Name, "deleted": name}
    raise ToolError("action must be 'list', 'add' or 'delete'.")


@office_tool("excel_core", "write", title="Recalculate / calculation mode")
def excel_calculate(workbook: str = "", scope: str = "workbook", mode: str = "") -> dict:
    """Force recalculation and/or change the calculation mode (affects the whole Excel instance).

    Args:
        workbook: exact workbook name ('' = active) - used to pick the sheet/instance.
        scope: 'workbook' (recalculate changed cells), 'sheet' (active/named sheet only), 'full' (all formulas), 'rebuild' (full rebuild with dependency tree).
        mode: '' leaves it unchanged, or 'automatic' / 'manual' / 'semiautomatic'.
    """
    app, wb = pick_workbook(workbook)
    if mode:
        codes = {"automatic": -4105, "manual": -4135, "semiautomatic": 2}
        if mode.lower() not in codes:
            raise ToolError("mode must be 'automatic', 'manual' or 'semiautomatic'.")
        app.Calculation = codes[mode.lower()]
    s = scope.lower()
    if s == "workbook":
        app.Calculate()
    elif s == "sheet":
        wb.ActiveSheet.Calculate()
    elif s == "full":
        app.CalculateFull()
    elif s == "rebuild":
        app.CalculateFullRebuild()
    else:
        raise ToolError("scope must be 'workbook', 'sheet', 'full' or 'rebuild'.")
    calc = {-4105: "automatic", -4135: "manual", 2: "semiautomatic"}.get(int(app.Calculation))
    return {"ok": True, "workbook": wb.Name, "calculation_mode": calc, "scope": s}
