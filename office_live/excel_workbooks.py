"""Excel workbook lifecycle and saving."""

import os
import pywintypes
from . import com
from .errors import ToolError
from .registry import office_tool
from .safety import EXCEL_EXTS, check_path
from .xl_common import (
    addr_of,
    all_workbooks,
    bounds,
    pick_sheet,
    pick_workbook,
    sheet_is_empty,
    sheet_names,
    suspend_events,
    validate_sheet_name,
    workbook_allowed,
)


_SAVE_FORMATS = {"xlsx": 51, "xlsm": 52, "xlsb": 50, "xls": 56, "csv": 62, "txt": 42, "html": 44, "htm": 44, "ods": 60, "xltx": 54, "xltm": 53}

_LEGACY_FORMATS = {
    56,  # xlExcel8
    43,  # xlExcel9795
    39,  # xlExcel7 / xlExcel5
    33,  # xlExcel4
    29,  # xlExcel3
    16,  # xlExcel2
    17,  # xlTemplate8
    18,  # xlAddIn8
    -4143,  # xlWorkbookNormal
}
_LEGACY_EXTENSIONS = {".xls", ".xlt", ".xla"}
_COMPATIBILITY_NOTE = (
    "The compatibility check was skipped when saving to Excel 97-2003 format; "
    "features unsupported by this format may have been lost. "
    "To keep them, ask the user before saving a copy in a modern format with excel_save_as: "
    ".xlsx for a workbook without macros, .xlsm to keep VBA macros. "
    "Other questions Excel asks while saving this format (such as 'recalculate all formulas when the file is opened') "
    "got Excel's default answer. "
    "The saved file stores compatibility checking as disabled, so after reopening it Excel will not run the check "
    "on save until the user turns it back on."
)


def _save_workbook(app, wb) -> bool:
    """Save legacy binary books without modal prompts; leave other formats alone.

    Workbook.CheckCompatibility=False stops the Compatibility Checker, but live probes (0.4.6) showed a second modal
    question during the same Save - "recalculate all formulas when the file is opened?" - that only
    Application.DisplayAlerts=False answers (with Excel's default, "yes"). So both are off for this Save only, and the
    reply says that the save's questions got default answers."""
    try:
        legacy = wb.FileFormat in _LEGACY_FORMATS
    except (pywintypes.com_error, AttributeError):
        legacy = os.path.splitext(wb.FullName)[1].lower() in _LEGACY_EXTENSIONS
    if not legacy:
        wb.Save()
        return False
    if wb.ReadOnly:
        raise ToolError(f"{wb.Name} is open read-only; use excel_save_as to write a copy.")

    alerts, compatibility = app.DisplayAlerts, wb.CheckCompatibility
    saved_after = False
    try:
        app.DisplayAlerts = False
        wb.CheckCompatibility = False
        wb.Save()
        saved_after = bool(wb.Saved)
    finally:
        try:
            wb.CheckCompatibility = compatibility
            if saved_after and not wb.Saved:
                # The file on disk keeps CheckCompatibility=False; restoring the in-memory flag can mark the book
                # dirty although the content is saved. Clear only that case: never after a failed Save() or a book
                # that was already dirty right after Save().
                wb.Saved = True
        except Exception:
            pass  # Restoration must not replace a Save() exception.
        _restore_alerts(app, alerts)  # independent of the CheckCompatibility restore above
    return True


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
    """Close an open workbook. Refuses when it has unsaved changes unless you choose: save=true (saves first, needs an existing path) or discard=true (throws the changes away). Never quits Excel itself. Saving .xls skips compatibility prompts and reports possible loss of unsupported features.

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
    compatibility_skipped = False
    if not wb.Saved:
        if save:
            if not wb.Path:
                raise ToolError(f"{name} has never been saved; use excel_save_as first or pass discard=true.")
            check_path(wb.FullName, "write")  # зона доступа и запретные каталоги — и для сохранения уже открытой книги
            compatibility_skipped = _save_workbook(app, wb)
            if compatibility_skipped and not wb.Saved:
                raise ToolError(f"{name} still has unsaved changes after Save(); the workbook was left open. Use excel_save_as to write a copy.")
            saved_first = True
        elif not discard:
            raise ToolError(f"{name} has unsaved changes. Pass save=true to save them or discard=true to drop them.")
    wb.Close(False)
    remaining = [w.Name for a, w in all_workbooks()] if _excel_running() else []
    result = {"ok": True, "closed": name, "saved_before_close": saved_first, "open_workbooks": remaining}
    if compatibility_skipped:
        result.update(compatibility_check_skipped=True, note=_COMPATIBILITY_NOTE)
    return result


@office_tool("excel_core", "save", title="Save workbook")
def excel_save(workbook: str = "") -> dict:
    """Save a workbook to its current file (Ctrl+S). A workbook that was never saved needs excel_save_as instead. Saving .xls skips compatibility prompts and reports possible loss of unsupported features.

    Args:
        workbook: exact name or '' for the active workbook.
    """
    app, wb = pick_workbook(workbook, allow_autosave=True)
    if not wb.Path:
        raise ToolError(f"{wb.Name} has never been saved - use excel_save_as with a path.")
    if wb.ReadOnly:
        raise ToolError(f"{wb.Name} is open read-only; use excel_save_as to write a copy.")
    check_path(wb.FullName, "write")  # зона доступа и запретные каталоги
    compatibility_skipped = _save_workbook(app, wb)
    result = {"ok": True, "workbook": wb.Name, "path": wb.FullName, "saved": bool(wb.Saved)}
    if compatibility_skipped:
        result.update(compatibility_check_skipped=True, note=_COMPATIBILITY_NOTE)
    return result


@office_tool("excel_core", "save", title="Save workbook as")
def excel_save_as(workbook: str, path: str, overwrite: bool = False) -> dict:
    """Save a workbook under a new name/format. Format follows the extension: .xlsx .xlsm .xlsb .xls .csv (UTF-8) .txt .html .ods. Refuses to replace an existing file unless overwrite=true. Returns links with the NEW file name; old link names resolve only inside this server session, not the separate browser link handler.

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
        old_names = (str(wb.Name), str(wb.FullName))
        wb.SaveAs(full, fmt)
        from .navigation import saved_as

        link_result = saved_as("excel", wb, old_names)
    finally:
        _restore_alerts(app, before)
    return {"ok": True, "workbook": wb.Name, "path": wb.FullName, "overwrite": overwrite,
            "note": "The workbook is now named after the new file. Old links resolve only within this server session; use the new links in the chat.", **link_result}


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
