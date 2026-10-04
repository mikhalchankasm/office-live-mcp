"""Работа по образцу: разбор файла, копия из шаблона, «чертёж» листа (шапка, столбцы, виды строк, правила), протягивание.

Сценарий «сделай аналог образца с новыми данными»:
  1. office_inspect_file(path)            — быстро понять файл, не открывая (листы, макросы, правила форматирования);
  2. excel_create_from_template(...)      — копия образца (формат, условное форматирование, макросы сохраняются);
  3. excel_describe_layout(...)           — чертёж листа: шапка, столбцы и их смысл, виды строк, формулы, правила;
  4. очистить область данных, записать новые данные (excel_write_range), протянуть формулы (excel_autofill / excel_set_formula).
"""

import os
import re
import shutil
import time
from collections import Counter

import pythoncom
import pywintypes

from . import com
from .errors import ToolError
from .excel_analysis import _kind
from .excel_core import excel_open_workbook
from .excel_format import _remember_view, _restore_view, cf_rules_detail
from .inspect_file import inspect_docx, inspect_xlsx
from .registry import office_tool
from .safety import EXCEL_EXTS, WORD_EXTS, check_path
from .util import a1_cell, col_letter, color_to_hex, to_grid
from .wd_common import pick_document
from .word_core import word_open_document
from .xl_common import (
    addr_of, bounds, get_range, number_format_for_read, pick_sheet, pick_workbook, preview, sheet_is_empty, sub_range,
)

MISSING = pythoncom.Missing
_MACRO_EXTS = {"xlsm", "xlsb", "xltm", "xlam"}
_TEMPLATE_EXTS = {"xltx", "xltm", "xlt"}


def _ext(path: str) -> str:
    return os.path.splitext(path)[1].lstrip(".").lower()


# ================================================================== разбор файла без открытия


@office_tool("bridge", "read", title="Inspect a file without opening it")
def office_inspect_file(path: str = "", name: str = "", peek_rows: int = 6, include_vba_source: bool = False, vba_module: str = "") -> dict:
    """Analyse an .xlsx/.xlsm/.docx/.docm file WITHOUT opening it in Office: sheets and their size, merged ranges, frozen panes, hidden columns, conditional-format rules, validations, tables/charts/pivots, defined names, the first rows of content, and the macros (module and procedure names, optionally the VBA source code - read only, never executed). Use it first to understand a sample file.

    Args:
        path: full path of the file. Alternatively pass `name` of a currently open, saved workbook/document.
        name: name of an open workbook or document (its file on disk is analysed - unsaved edits are not included).
        peek_rows: how many leading rows/paragraphs to preview (max 20).
        include_vba_source: include macro source code in the result.
        vba_module: restrict source output to one module.
    """
    full = ""
    if path:
        full = check_path(path, "read")
    elif name:
        try:
            _, wb = pick_workbook(name)
            full = wb.FullName if wb.Path else ""
        except ToolError:
            _, doc = pick_document(name)
            full = doc.FullName if doc.Path else ""
        if not full:
            raise ToolError("That document has never been saved, so there is no file to analyse.")
        full = check_path(full, "read")
    else:
        raise ToolError("Pass `path` or `name`.")
    if not os.path.isfile(full):
        raise ToolError(f"File not found: {full}")
    ext = _ext(full)
    try:
        if ext in ("xlsx", "xlsm", "xltx", "xltm"):
            info = inspect_xlsx(full, peek_rows, include_vba_source, vba_module)
        elif ext in ("docx", "docm", "dotx", "dotm"):
            info = inspect_docx(full, min(int(peek_rows), 20) or 8, include_vba_source, vba_module)
        else:
            raise ToolError(f"Unsupported file type '.{ext}' (supported: xlsx, xlsm, xltx, xltm, docx, docm, dotx, dotm).")
    except (OSError, KeyError) as exc:
        raise ToolError(f"Could not read the file as an Office package: {exc}") from None
    except Exception as exc:  # noqa: BLE001 — zipfile.BadZipFile и разбор XML
        if isinstance(exc, ToolError):
            raise
        raise ToolError(f"Could not analyse the file ({type(exc).__name__}: {exc}).") from None
    info["size_kb"] = round(os.path.getsize(full) / 1024)
    return info


# ================================================================== копия из образца


@office_tool("excel_core", "save", title="New workbook from a sample/template")
def excel_create_from_template(template_path: str, new_path: str, overwrite: bool = False) -> dict:
    """Make a new workbook as an exact copy of a sample file (all formatting, merged cells, column widths, conditional formats, formulas, named ranges, and macros are kept) and open it. The sample itself is never modified. Then clear the old data and write the new data. For .xlsm samples keep the .xlsm extension; macros stay disabled in the copy until the user enables them in Excel.

    Args:
        template_path: full path of the sample (.xlsx/.xlsm/.xltx/.xltm).
        new_path: full path of the new file; same extension as the sample (templates .xltx/.xltm become .xlsx/.xlsm).
        overwrite: allow replacing an existing file at new_path.
    """
    src = check_path(template_path, "read")
    if not os.path.isfile(src):
        raise ToolError(f"Template not found: {src}")
    dst = check_path(new_path, "write", EXCEL_EXTS - {"pdf", "csv", "txt", "html", "htm"})
    s_ext, d_ext = _ext(src), _ext(dst)
    if s_ext in _TEMPLATE_EXTS:
        want = "xlsm" if s_ext == "xltm" else "xlsx"
        if d_ext != want:
            raise ToolError(f"A .{s_ext} template produces a .{want} file; use that extension for new_path.")
    elif s_ext != d_ext:
        raise ToolError(f"new_path must keep the sample's extension .{s_ext} (macros and formats depend on it).")
    if os.path.exists(dst) and not overwrite:
        raise ToolError(f"File already exists: {dst}. Pass overwrite=true to replace it.")
    if os.path.normcase(src) == os.path.normcase(dst):
        raise ToolError("new_path must differ from the template path.")
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    if s_ext in _TEMPLATE_EXTS:
        app = com.primary_app("excel", launch=True)
        with com.macros_disabled(app):
            wb = app.Workbooks.Add(src)
            before = app.DisplayAlerts
            app.DisplayAlerts = False
            try:
                wb.SaveAs(dst, 52 if d_ext == "xlsm" else 51)
            finally:
                app.DisplayAlerts = before
        info = {"ok": True, "workbook": wb.Name, "path": wb.FullName, "sheets": [wb.Sheets(i).Name for i in range(1, int(wb.Sheets.Count) + 1)]}
    else:
        if os.path.exists(dst) and overwrite and _is_open(dst):
            raise ToolError(f"{dst} is open in Office; close it before overwriting.")
        shutil.copy2(src, dst)
        info = excel_open_workbook(dst, False)
    info["created_from"] = src
    info["macros_in_file"] = d_ext in _MACRO_EXTS
    if d_ext in _MACRO_EXTS:
        info["note"] = "Macros are disabled in this session's copy (they are kept in the file). Enable them in Excel if the user needs them."
    return info


def _is_open(path: str) -> bool:
    try:
        for app in com.apps("excel"):
            for i in range(1, int(app.Workbooks.Count) + 1):
                wb = app.Workbooks(i)
                if wb.Path and os.path.normcase(os.path.abspath(wb.FullName)) == os.path.normcase(path):
                    return True
    except ToolError:
        pass
    return False


@office_tool("word_core", "save", title="New document from a sample/template")
def word_create_from_template(template_path: str, new_path: str, overwrite: bool = False) -> dict:
    """Make a new Word document as an exact copy of a sample (.docx/.docm/.dotx/.dotm: styles, headers, footers, tables, placeholders, macros kept) and open it. The sample is never modified. Then use word_fill_placeholders / word_insert_text / word_write_table to put the new content in.

    Args:
        template_path: full path of the sample.
        new_path: full path of the new file; same extension as the sample (.dotx/.dotm templates become .docx/.docm).
        overwrite: allow replacing an existing file.
    """
    src = check_path(template_path, "read")
    if not os.path.isfile(src):
        raise ToolError(f"Template not found: {src}")
    dst = check_path(new_path, "write", WORD_EXTS - {"pdf", "txt", "html", "htm", "rtf", "xml", "odt"})
    s_ext, d_ext = _ext(src), _ext(dst)
    if s_ext in ("dotx", "dotm"):
        want = "docm" if s_ext == "dotm" else "docx"
        if d_ext != want:
            raise ToolError(f"A .{s_ext} template produces a .{want} file; use that extension for new_path.")
    elif s_ext != d_ext:
        raise ToolError(f"new_path must keep the sample's extension .{s_ext}.")
    if os.path.exists(dst) and not overwrite:
        raise ToolError(f"File already exists: {dst}. Pass overwrite=true to replace it.")
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    if s_ext in ("dotx", "dotm"):
        app = com.primary_app("word", launch=True)
        with com.macros_disabled(app):
            doc = app.Documents.Add(src)
        before = app.DisplayAlerts
        app.DisplayAlerts = 0
        try:
            doc.SaveAs2(dst, 13 if d_ext == "docm" else 16)
        finally:
            app.DisplayAlerts = before
        info = {"ok": True, "document": doc.Name, "path": doc.FullName}
    else:
        shutil.copy2(src, dst)
        info = word_open_document(dst, False)
    info["created_from"] = src
    return info


# ================================================================== протягивание


_FILL = {"default": 0, "copy": 1, "series": 2, "formats": 3, "values": 4, "days": 5, "weekdays": 6, "months": 7, "years": 8, "linear": 9, "growth": 10}


@office_tool("excel_core", "write", title="Autofill (drag the fill handle)")
def excel_autofill(workbook: str, sheet: str, source: str, dest: str, fill_type: str = "default") -> dict:
    """Do what dragging the fill handle does: extend a pattern from `source` over `dest` (which must contain `source`). Formulas adjust their relative references, numbers/dates continue as a series, formats are copied. Use it to carry a template row's formulas AND formatting down to new rows.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        source: the cells that hold the pattern, e.g. 'A5:AT5' (one or several rows/columns).
        dest: the full target range including the source, e.g. 'A5:AT300'.
        fill_type: default (smart) | copy (repeat without series) | series | formats (formatting only) | values | days | weekdays | months | years | linear | growth.
    """
    app, wb = pick_workbook(workbook)
    ws, src = get_range(wb, sheet, source, empty_means_used=False)
    _, dst = get_range(wb, sheet, dest, empty_means_used=False)
    t = fill_type.lower()
    if t not in _FILL:
        raise ToolError(f"fill_type must be one of {sorted(_FILL)}")
    try:
        src.AutoFill(dst, _FILL[t])
    except pywintypes.com_error as exc:
        raise ToolError("Excel could not autofill (the destination must include and extend the source range in one direction): " + com.com_error_text(exc)) from None
    r1, c1, r2, c2 = bounds(dst)
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "filled": addr_of(dst), "fill_type": t, "preview": preview(sub_range(ws, r1, c1, min(r2, r1 + 4), c2), 120)}


# ================================================================== чертёж листа


def _freeze_state(app, wb, ws) -> dict:
    prev = _remember_view(app)
    try:
        wb.Activate()
        ws.Activate()
        win = app.ActiveWindow
        return {
            "frozen": bool(win.FreezePanes), "split_row": int(win.SplitRow), "split_column": int(win.SplitColumn),
            "zoom": int(win.Zoom), "gridlines": bool(win.DisplayGridlines),
        }
    except pywintypes.com_error:
        return {}
    finally:
        _restore_view(app, prev)


def _clean(s) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()


@office_tool("excel_analysis", "read", title="Describe sheet layout (blueprint)")
def excel_describe_layout(
    workbook: str = "",
    sheet: str = "",
    header_rows: int = 0,
    key_columns: list[str] | None = None,
    analyze_row_styles: bool = True,
    max_style_rows: int = 800,
    max_columns: int = 120,
) -> dict:
    """Produce a BLUEPRINT of a sheet so it can be re-created with new data: where the title/header block ends (frozen rows), merged header cells, every column (header text, width, hidden, number format, data type, typical values, the formula it contains), the KINDS of data rows (e.g. section rows vs item rows - grouped by how they actually look on screen incl. conditional formatting - with examples and code patterns), conditional-format rules in English with their colors, hidden rows, tables/charts/pivots/validation/macros. Read it, then build the analogue.

    Args:
        workbook: exact name or '' for the active workbook.
        sheet: sheet name or '' for active.
        header_rows: number of header rows (0 = take the frozen rows, else 1).
        key_columns: column letters whose look defines the row kinds (default: the two most filled text columns).
        analyze_row_styles: classify rows by appearance (slower: a few COM calls per row). Disable for huge sheets.
        max_style_rows: how many data rows to classify (evenly sampled if there are more).
        max_columns: cap on described columns.
    """
    started = time.monotonic()
    app, wb = pick_workbook(workbook)
    ws = pick_sheet(wb, sheet)
    if sheet_is_empty(app, ws):
        return {"workbook": wb.Name, "sheet": ws.Name, "empty": True}
    used = ws.UsedRange
    r1, c1, r2, c2 = bounds(used)
    ncols = min(c2 - c1 + 1, int(max_columns))
    rows_limit = max(1, min(r2 - r1 + 1, 400000 // max(ncols, 1)))
    block = sub_range(ws, r1, c1, r1 + rows_limit - 1, c1 + ncols - 1)
    values = to_grid(block.Value)
    r1c1 = to_grid(block.FormulaR1C1)

    def row_has_data(row):
        return any(v not in (None, "") for v in row)

    last_idx = max((i for i, row in enumerate(values) if row_has_data(row)), default=0)
    view = _freeze_state(app, wb, ws)
    # SplitRow считается от верха окна (строка 1), а используемая область может начинаться ниже — переводим в число строк шапки
    frozen_rows = int(view.get("split_row") or 0) if view.get("frozen") else 0
    relative = frozen_rows - (r1 - 1)
    hdr = int(header_rows) or (relative if relative >= 1 else 1)
    hdr = max(1, min(hdr, last_idx + 1))
    data_first_idx = hdr  # индекс первой строки данных в values
    out: dict = {
        "workbook": wb.Name, "sheet": ws.Name, "used_range": addr_of(used),
        "data_rows": {"first": r1 + data_first_idx, "last": r1 + last_idx, "count": max(0, last_idx - data_first_idx + 1)},
        "header_rows": hdr, "view": view,
        "truncated": {
            "columns": (c2 - c1 + 1) > ncols, "rows": (r2 - r1 + 1) > rows_limit,
            "note": "Only the first max_columns columns / the first rows were analysed." if (c2 - c1 + 1) > ncols or (r2 - r1 + 1) > rows_limit else None,
        },
    }
    # --- шапка: значения и объединения
    header_texts = {}
    header_merges = []
    wide_rows = set()
    for i in range(hdr):
        for j in range(ncols):
            cell = ws.Cells(r1 + i, c1 + j)
            try:
                merged = bool(cell.MergeCells)
            except pywintypes.com_error:
                merged = False
            if merged:
                area = cell.MergeArea
                ar1, ac1, ar2, ac2 = bounds(area)
                if (ar1, ac1) == (r1 + i, c1 + j):
                    header_merges.append(addr_of(area))
                    if ac2 - ac1 + 1 >= max(3, ncols // 2):
                        wide_rows.add(i)  # объединение на пол-листа — строка-заголовок («название отчёта»)
                    v = values[i][j]
                    if v not in (None, ""):
                        for cc in range(ac1, ac2 + 1):
                            for rr in range(ar1, ar2 + 1):
                                header_texts.setdefault((rr - r1, cc - c1), _clean(v))
            else:
                v = values[i][j]
                if v not in (None, ""):
                    header_texts[(i, j)] = _clean(v)
    title_rows = set(wide_rows)
    for i in range(hdr):
        # почти пустая строка — заголовок отчёта, дата и т.п. Но в узкой таблице (меньше 5 столбцов) это обычная шапка
        if ncols >= 5 and sum(1 for j in range(ncols) if values[i][j] not in (None, "")) <= 2:
            title_rows.add(i)
    out["header_merged_ranges"] = header_merges[:60]
    out["title_rows"] = [
        {"row": r1 + i, "cells": [{"cell": a1_cell(r1 + i, c1 + j), "text": _clean(values[i][j])[:100]} for j in range(ncols) if values[i][j] not in (None, "")][:6]}
        for i in sorted(title_rows)
    ]
    for key in [k for k in header_texts if k[0] in title_rows]:
        del header_texts[key]
    # --- столбцы
    cols = []
    data_rows_vals = values[data_first_idx:last_idx + 1]
    data_rows_f = r1c1[data_first_idx:last_idx + 1]
    for j in range(ncols):
        absc = c1 + j
        labels = []
        for i in range(hdr):
            t = header_texts.get((i, j))
            if t and t not in labels:
                labels.append(t)
        colvals = [row[j] for row in data_rows_vals]
        kinds = Counter(_kind(v) for v in colvals)
        present = [v for v in colvals if _kind(v) != "empty"]
        cell0 = ws.Cells(r1 + data_first_idx, absc)
        entry = {
            "column": col_letter(absc), "header": " / ".join(labels)[:140] or None,
            "width": float(cell0.ColumnWidth), "hidden": bool(cell0.EntireColumn.Hidden),
            "number_format": number_format_for_read(app, cell0.NumberFormat), "non_empty": len(present),
            "types": {k: v for k, v in kinds.items() if k != "empty"},
        }
        if present:
            uniq = []
            for v in present:
                if v not in uniq:
                    uniq.append(v)
                if len(uniq) > 30:
                    break
            if len(uniq) <= 12 and len(present) > len(uniq):
                counts = Counter(str(v) for v in present)
                entry["values"] = [{"value": k[:40], "count": n} for k, n in counts.most_common(8)]
            else:
                entry["samples"] = [str(v)[:40] for v in present[:3]]
            nums = [v for v in present if _kind(v) == "number"]
            if nums:
                entry["range"] = [min(nums), max(nums)]
            dates = sorted(str(v) for v in present if _kind(v) == "date")
            if dates:
                entry["dates"] = [dates[0], dates[-1]]
        fcounts = Counter(row[j] for row in data_rows_f if isinstance(row[j], str) and row[j].startswith("="))
        if fcounts:
            f, n = fcounts.most_common(1)[0]
            first_idx = next(k for k, row in enumerate(data_rows_f) if row[j] == f)
            example_cell = ws.Cells(r1 + data_first_idx + first_idx, absc)
            entry["formula_r1c1"] = {"formula": f[:200], "cells": n, "variants": len(fcounts)}
            entry["formula_example"] = {"cell": addr_of(example_cell), "formula": example_cell.Formula[:300]}
        cols.append(entry)
    out["columns"] = cols
    # --- виды строк по внешнему виду
    if analyze_row_styles and last_idx >= data_first_idx:
        if key_columns:
            from .util import col_number

            keys = [col_number(k) - c1 for k in key_columns if c1 <= col_number(k) <= c1 + ncols - 1]
        else:
            ranked = sorted(range(ncols), key=lambda j: (-cols[j]["non_empty"] if cols[j]["types"].get("text", 0) >= 0.5 * max(cols[j]["non_empty"], 1) else 1, j))
            keys = sorted(ranked[:2])
        row_idxs = [i for i in range(data_first_idx, last_idx + 1) if row_has_data(values[i])]
        step = max(1, len(row_idxs) // max(1, int(max_style_rows)))
        sampled = row_idxs[::step]
        groups: dict = {}
        # скрытые строки одним вызовом: видимые области ПЕРВОГО ВИДИМОГО столбца -> пропуски
        first_abs, last_abs = r1 + row_idxs[0], r1 + row_idxs[-1]
        hidden_rows = []
        probe_col = next((c1 + j for j in range(ncols) if not bool(ws.Cells(first_abs, c1 + j).EntireColumn.Hidden)), None)
        if probe_col is not None:
            try:
                visible = set()
                for area in ws.Range(f"{col_letter(probe_col)}{first_abs}:{col_letter(probe_col)}{last_abs}").SpecialCells(12).Areas:  # xlCellTypeVisible
                    a1_, _, a2_, _ = bounds(area)
                    visible.update(range(a1_, a2_ + 1))
                hidden_rows = [r for r in range(first_abs, last_abs + 1) if r not in visible]
            except pywintypes.com_error:
                hidden_rows = list(range(first_abs, last_abs + 1))  # видимых ячеек нет: скрыты все строки
        partial = False
        for i in sampled:
            if time.monotonic() - started > 60:
                partial = True
                break
            rr = r1 + i
            sig = []
            for j in keys:
                df = ws.Cells(rr, c1 + j).DisplayFormat
                fill = None if int(df.Interior.ColorIndex) == -4142 else color_to_hex(df.Interior.Color)
                sig.append((fill, bool(df.Font.Bold), int(df.IndentLevel)))
            text0 = values[i][keys[0]] if keys else ""
            pattern = re.sub(r"\d", "#", str(text0))[:30] if text0 not in (None, "") else "(empty)"
            key = tuple(sig)
            g = groups.setdefault(key, {"rows": [], "patterns": Counter(), "sample": [str(values[i][j])[:50] for j in keys]})
            g["rows"].append(rr)
            g["patterns"][pattern] += 1
        kinds_out = []
        for n, (sig, g) in enumerate(sorted(groups.items(), key=lambda kv: -len(kv[1]["rows"])), start=1):
            heights = sorted({float(ws.Cells(r, c1).RowHeight) for r in (g["rows"][:3] + g["rows"][-2:])})
            kinds_out.append({
                "row_height_sample": heights[:4],
                "kind": f"row_kind_{n}", "rows": len(g["rows"]), "first_rows": g["rows"][:4], "last_row": g["rows"][-1],
                "look": [{"column": col_letter(c1 + keys[k]), "fill": s[0], "bold": s[1], "indent": s[2]} for k, s in enumerate(sig)],
                "key_text_patterns": dict(g["patterns"].most_common(4)), "sample": g["sample"],
            })
        hierarchy = None
        if keys:
            kv = [str(values[i][keys[0]]) for i in row_idxs if values[i][keys[0]] not in (None, "")]
            if kv and sum(1 for t in kv if re.fullmatch(r"[\d]+(\.[\d]+)*", t.strip())) >= 0.8 * len(kv):
                depth = Counter(t.strip().count(".") for t in kv)
                hierarchy = {"column": col_letter(c1 + keys[0]), "meaning": "dotted code; level = number of dots", "levels": {str(k): v for k, v in sorted(depth.items())}}
        out["row_kinds"] = {
            "key_columns": [col_letter(c1 + j) for j in keys], "sampled_every": step, "hierarchy": hierarchy, "kinds": kinds_out, "partial": partial,
            "hidden_rows": hidden_rows[:60],
            "note": "Looks include conditional formatting (what the user sees). Kinds usually mirror a hierarchy encoded in a code column; see conditional_formats for the rules that drive the styling.",
        }
    # --- правила, имена, объекты
    try:
        out["conditional_formats"] = cf_rules_detail(app, wb, ws)
    except pywintypes.com_error:
        out["conditional_formats"] = []
    by_letter = {c["column"]: c for c in cols}
    for rule in out["conditional_formats"]:
        refs = []
        for m in re.finditer(r"\$?([A-Z]{1,3})\$?\d+", rule.get("formula") or ""):
            if m.group(1) in by_letter and m.group(1) not in refs:
                refs.append(m.group(1))
        if refs:
            rule["depends_on"] = [
                {"column": L, "header": by_letter[L]["header"], "hidden": by_letter[L]["hidden"], "values": by_letter[L].get("values") or by_letter[L].get("samples")}
                for L in refs
            ]
    hidden_deps = sorted({d["column"] for r in out["conditional_formats"] for d in r.get("depends_on", []) if d["hidden"]})
    if hidden_deps:
        out["hidden_columns_drive_formatting"] = {
            "columns": hidden_deps,
            "note": "Conditional-format rules read these HIDDEN columns. New rows must carry the same kind of values there (see each rule's depends_on.values) or the highlighting will differ from the sample.",
        }
    try:
        out["data_validation_areas"] = int(ws.Cells.SpecialCells(-4174).Areas.Count)  # xlCellTypeAllValidation
    except pywintypes.com_error:
        out["data_validation_areas"] = 0
    names = [wb.Names(i) for i in range(1, int(wb.Names.Count) + 1)]
    out["defined_names"] = {"visible": sum(1 for n in names if bool(n.Visible)), "hidden": sum(1 for n in names if not bool(n.Visible))}
    out["objects"] = {
        "tables": int(ws.ListObjects.Count), "charts": int(ws.ChartObjects().Count), "pivot_tables": int(ws.PivotTables().Count),
        "shapes": int(ws.Shapes.Count), "autofilter": bool(ws.AutoFilterMode),
    }
    try:
        out["body_has_merged_cells"] = used.MergeCells is not False
        out["has_vba_project"] = bool(wb.HasVBProject)
    except pywintypes.com_error:
        pass
    out["elapsed_s"] = round(time.monotonic() - started, 1)
    return out
