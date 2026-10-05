"""Excel: умные таблицы, сводные таблицы, диаграммы, профилирование данных, поиск проблем."""

import json
import datetime
import math
import os
import re
import shutil
import statistics
import tempfile
from collections import Counter

import pythoncom
import pywintypes
from mcp.server.mcpserver import Image

from . import com, config
from .errors import ToolError
from .registry import office_tool
from .safety import check_path
from .util import EXCEL_ERRORS, a1_cell, a1_range, col_letter, col_number, norm_value, parse_color, smart_number, split_sheet_ref, to_grid
from .excel_format import NUMBER_FORMATS
from .xl_common import (
    number_format_for_write, addr_of, bounded_range, bounds, clip_to_used, get_range, pick_sheet, pick_workbook, preview, sheet_is_empty, sheet_names,
    sub_range, validate_sheet_name,
)

MISSING = pythoncom.Missing


@office_tool("excel_analysis", "read", title="Trace formula")
def excel_trace_formula(workbook: str = "", sheet: str = "", cell: str = "", direction: str = "precedents",
                        depth: int = 1, max_nodes: int = 200) -> dict:
    """Trace formula references without selection changes, sheet activation, arrows or evaluation. Edges always point from input to consumer. Never treat complete=false as a full graph.

    Args:
        workbook, sheet: target workbook/sheet; empty selects the active one.
        cell: required single cell.
        direction: precedents (inputs) or dependents (consumers).
        depth: 1..5, default 1; an unexpanded edge sets depth_limit.
        max_nodes: 1..2000 including the root, default 200.

    Uses DirectPrecedents/DirectDependents on the already active sheet plus a conservative formula tokenizer.
    Supports A1 ranges, quoted sheets, range names and Table[Column]. Dependents scan Formula arrays once, at most 500000 cells across the workbook.
    A separate budget of 2000000 reference containment checks bounds large dependent graphs (reference_check_limit).
    Reports incomplete_reasons for inactive-sheet COM, ambiguous COM 1004, dynamic/external/3D/spill references, INDIRECT/OFFSET/INDEX/LET/LAMBDA,
    formula names, unsupported structured references/syntax and limits. No external workbooks are opened.
    """
    from .formula_trace import Trace, validate_trace

    validate_trace(direction, depth, max_nodes)
    app, wb = pick_workbook(workbook)
    _, original = get_range(wb, sheet, cell, empty_means_used=False)
    if int(original.Areas.Count) != 1 or int(original.Rows.Count) * int(original.Columns.Count) != 1:
        raise ToolError("cell must resolve to one cell (not a whole row/column).")
    _, rng = bounded_range(app, wb, sheet, cell, 1)
    if rng is None:
        raise ToolError("cell must resolve to one cell.")
    return {"workbook": wb.Name, "direction": direction, **Trace(app, wb, direction, depth, max_nodes).run(rng)}


def _compare_text(value, ignore_case, ignore_whitespace):
    if isinstance(value, str):
        if ignore_whitespace:
            value = " ".join(value.split())
        if ignore_case:
            value = value.casefold()
    return value


def _compare_type(value):
    if value is None or value == "":
        return "empty"
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, (float, int)):
        return "number"
    return "text"


def _compare_short(value):
    if isinstance(value, dict):
        return {k: _compare_short(v) for k, v in value.items()}
    return value[:185] + "... [truncated]" if isinstance(value, str) and len(value) > 200 else value


def _compare_values(rng):
    """Value2 сохраняет точность чисел, Value даёт тип даты, как excel_read_range."""
    values = to_grid(rng.Value2)
    raw = rng.Value
    dates = raw if isinstance(raw, tuple) and raw and isinstance(raw[0], tuple) else ((raw,),)
    for i, row in enumerate(dates):
        for j, value in enumerate(row):
            if isinstance(value, (datetime.date, datetime.time)):
                values[i][j] = norm_value(value)
    return values


@office_tool("excel_analysis", "read", title="Compare ranges")
def excel_compare_ranges(
    cells_a: str,
    workbook_a: str = "",
    sheet_a: str = "",
    workbook_b: str = "",
    sheet_b: str = "",
    cells_b: str = "",
    compare: str = "values",
    match: str = "position",
    key_column: str = "",
    ignore_case: bool = False,
    ignore_whitespace: bool = False,
    tolerance: float = 0,
    max_differences: int = 200,
) -> dict:
    """Compare two rectangles by position or unique row keys, across sheets, workbooks or Excel instances; never modifies either range. Reports values/types, formula text or five direct formats, with bounded differences.

    Args:
        cells_a: required A1 rectangle or name; whole rows/columns are clipped to the used range.
        workbook_a, sheet_a: first workbook and sheet (empty = active).
        workbook_b, sheet_b, cells_b: second target; omitted fields inherit A. Targets must differ.
        compare: values (Value2, dates as ISO), formulas (English Formula), or formats (NumberFormat, bold, italic, font/fill color).
        match: position (intersection plus extra rectangles), or key (first row contains unique headers).
        key_column: required for key matching: header text or absolute sheet column letter; resolves separately in each range.
        ignore_case, ignore_whitespace: normalize text, including keys/headers; whitespace trims and collapses NBSP and other whitespace.
        tolerance: finite nonnegative absolute numeric tolerance; keys always match exactly by type and value.
        max_differences: 1..2000, default 200. Texts over 200 characters are marked as truncated.

    Limits: 200000 cells per range for values/formulas; 5000 for formats. Empty keys are skipped and counted; duplicate keys/headers are refused.
    """
    if compare not in {"values", "formulas", "formats"} or match not in {"position", "key"}:
        raise ToolError("compare must be values, formulas or formats; match must be position or key.")
    if not math.isfinite(tolerance) or tolerance < 0:
        raise ToolError("tolerance must be finite and nonnegative.")
    if isinstance(max_differences, bool) or not isinstance(max_differences, int) or not 1 <= max_differences <= 2000:
        raise ToolError("max_differences must be an integer between 1 and 2000.")
    if match == "key" and not key_column.strip():
        raise ToolError("key_column is required for match='key'.")
    aa, wa = pick_workbook(workbook_a)
    ab, wb = pick_workbook(workbook_b) if workbook_b else (aa, wa)
    limit = 5000 if compare == "formats" else 200000
    sa, ra = bounded_range(aa, wa, sheet_a, cells_a, limit)
    second_cells = cells_b or split_sheet_ref(cells_a)[1]
    second_sheet = sheet_b or ("" if split_sheet_ref(second_cells)[0] else sa.Name)
    sb, rb = bounded_range(ab, wb, second_sheet, second_cells, limit)
    if (int(aa.Hwnd), wa.FullName.casefold(), sa.Name.casefold(), cells_a if ra is None else addr_of(ra)) == (
            int(ab.Hwnd), wb.FullName.casefold(), sb.Name.casefold(), second_cells if rb is None else addr_of(rb)):
        raise ToolError("Choose different workbooks, sheets or cell addresses.")

    def load(ws, rng):
        if rng is None:
            return {"values": [], "rows": 0, "cols": 0, "r": 1, "c": 1, "data": []}
        r, c, end, right = bounds(rng)
        values = _compare_values(rng) if compare != "formats" or match == "key" else []
        data = values
        if compare == "formulas":
            formulas = to_grid(rng.Formula)
            data = [[(f, True) if isinstance(f, str) and f.startswith("=") and bool(ws.Cells(r + i, c + j).HasFormula)
                     else (values[i][j], False) for j, f in enumerate(row)] for i, row in enumerate(formulas)]
        elif compare == "formats":
            data = []
            for row in range(r, end + 1):
                line = []
                for col in range(c, right + 1):
                    cell = ws.Cells(row, col)
                    line.append({"number_format": cell.NumberFormat, "bold": bool(cell.Font.Bold), "italic": bool(cell.Font.Italic),
                                 "font_color": cell.Font.Color, "fill_color": cell.Interior.Color})
                data.append(line)
        return {"values": values, "rows": end - r + 1, "cols": right - c + 1, "r": r, "c": c, "data": data}

    a, b = load(sa, ra), load(sb, rb)
    summary = {"compared_cells": 0, "compared_rows": 0, "differences": dict.fromkeys(("value", "type", "formula", "format"), 0),
               "rows_only_in_a": 0, "rows_only_in_b": 0, "skipped_empty_keys_a": 0, "skipped_empty_keys_b": 0}
    differences, only_a, only_b = [], [], []
    headers_only_a, headers_only_b = [], []

    def address(side, row, col):
        return a1_cell(side["r"] + row, side["c"] + col)

    def normalized(value):
        value = _compare_text(value, ignore_case, ignore_whitespace)
        return _compare_type(value), None if value is None or value == "" else value

    def compare_cell(ar, ac, br, bc, extra):
        x, y = a["data"][ar][ac], b["data"][br][bc]
        summary["compared_cells"] += 1
        kind = None
        if compare == "formats":
            kind = "format" if x != y else None
        else:
            formula = False
            if compare == "formulas":
                (x, xf), (y, yf) = x, y
                formula = xf or yf
                if xf != yf:
                    kind = "formula"
            tx, nx = normalized(x)
            ty, ny = normalized(y)
            if kind is None:
                if tx != ty:
                    kind = "type"
                elif tx == "number":
                    kind = "value" if abs(nx - ny) > tolerance else None
                elif nx != ny:
                    kind = "value"
            if formula and kind:
                kind = "formula"
        if kind:
            summary["differences"][kind] += 1
            if len(differences) < max_differences:
                differences.append({"cell_a": address(a, ar, ac), "cell_b": address(b, br, bc), "kind": kind,
                                    "a": _compare_short(x), "b": _compare_short(y), **extra})

    if match == "position":
        rows, cols = min(a["rows"], b["rows"]), min(a["cols"], b["cols"])
        for r in range(rows):
            for c in range(cols):
                compare_cell(r, c, r, c, {})
        summary["compared_rows"] = rows
        for name, side, output in (("a", a, only_a), ("b", b, only_b)):
            summary["rows_only_in_" + name] = side["rows"] - rows
            if side["rows"] > rows and side["cols"]:
                output.append(a1_range(side["r"] + rows, side["c"], side["r"] + side["rows"] - 1, side["c"] + side["cols"] - 1))
            if side["cols"] > cols and rows:
                output.append(a1_range(side["r"], side["c"] + cols, side["r"] + rows - 1, side["c"] + side["cols"] - 1))
    else:
        def index(side, label):
            if not side["values"]:
                raise ToolError(f"Range {label} needs a header row for key matching.")
            headers, keys, duplicates = {}, {}, []
            for i, value in enumerate(side["values"][0]):
                key = normalized(value)
                if key in headers:
                    raise ToolError(f"Duplicate header in {label}: {_compare_short(value)!r}")
                headers[key] = i
            needle = normalized(key_column)
            col = headers.get(needle)
            if col is None and re.fullmatch(r"[A-Za-z]{1,3}", key_column):
                candidate = col_number(key_column) - side["c"]
                if 0 <= candidate < side["cols"]:
                    col = candidate
            if col is None:
                raise ToolError(f"key_column {key_column!r} not found in {label}.")
            for i, row in enumerate(side["values"][1:], 1):
                key = normalized(row[col])
                if key[0] == "empty":
                    summary["skipped_empty_keys_" + label.lower()] += 1
                elif key in keys:
                    if len(duplicates) < 10:
                        duplicates.append(_compare_short(row[col]))
                else:
                    keys[key] = i
            if duplicates:
                raise ToolError(f"Duplicate keys in {label}: {duplicates}")
            return headers, keys, col

        ha, ka, ca = index(a, "A")
        hb, kb, _ = index(b, "B")
        headers_only_a = [_compare_short(a["values"][0][c]) for h, c in ha.items() if h not in hb]
        headers_only_b = [_compare_short(b["values"][0][c]) for h, c in hb.items() if h not in ha]
        for key, ar in ka.items():
            if key not in kb:
                only_a.append(a1_range(a["r"] + ar, a["c"], a["r"] + ar, a["c"] + a["cols"] - 1))
                continue
            br = kb[key]
            summary["compared_rows"] += 1
            for header, ac in ha.items():
                if header in hb:
                    compare_cell(ar, ac, br, hb[header], {"key": _compare_short(a["values"][ar][ca]), "column": _compare_short(a["values"][0][ac])})
        only_b = [a1_range(b["r"] + br, b["c"], b["r"] + br, b["c"] + b["cols"] - 1) for key, br in kb.items() if key not in ka]
        summary.update(rows_only_in_a=len(only_a), rows_only_in_b=len(only_b))
    total = sum(summary["differences"].values())
    # списки строк «только в A/B» при сравнении по ключу растут со строками — режем тем же лимитом
    return {"summary": summary, "differences": differences, "truncated": total > len(differences), "remaining_differences": total - len(differences),
            "target_a": {"workbook": wa.FullName, "sheet": sa.Name}, "target_b": {"workbook": wb.FullName, "sheet": sb.Name},
            "only_in_a": only_a[:max_differences], "only_in_b": only_b[:max_differences],
            "only_truncated": len(only_a) > max_differences or len(only_b) > max_differences,
            "headers_only_in_a": headers_only_a, "headers_only_in_b": headers_only_b}

# ================================================================== умные таблицы (ListObject)


def _find_table(wb, name: str):
    names = []
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        for j in range(1, int(ws.ListObjects.Count) + 1):
            lo = ws.ListObjects(j)
            names.append(lo.Name)
            if lo.Name.lower() == name.lower():
                return ws, lo
    raise ToolError(f"Table '{name}' not found. Tables in the workbook: {names}")


@office_tool("excel_analysis", "write", title="Excel tables", read_actions=("list",), destructive=True)
def excel_manage_tables(
    workbook: str,
    action: str = "list",
    sheet: str = "",
    cells: str = "",
    name: str = "",
    new_name: str = "",
    style: str = "",
    has_headers: bool = True,
    show_totals: bool | None = None,
    show_filter_buttons: bool | None = None,
    banded_rows: bool | None = None,
) -> dict:
    """Work with Excel tables (the structured 'Format as Table' objects with filter buttons and auto-growing ranges).

    Args:
        workbook: exact workbook name.
        action: 'list', 'create' (from `cells` on `sheet`), 'modify' (style/totals/filter/stripes/rename/resize by `name`), 'delete' (removes table AND its data), 'to_range' (keeps the data, removes the table object).
        sheet: sheet for 'create' ('' = active).
        cells: for create - the data block incl. header row; for modify - a new range to resize the table to.
        name: table name (create: desired name; others: existing table).
        new_name: rename on modify.
        style: table style such as 'TableStyleMedium2', 'TableStyleLight9', 'TableStyleMedium9' (Light1-21, Medium1-28, Dark1-11).
        has_headers: the first row of `cells` holds headers (create).
        show_totals, show_filter_buttons, banded_rows: switches for modify/create.
    """
    app, wb = pick_workbook(workbook)
    act = action.lower()
    if act == "list":
        out = []
        for i in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(i)
            for j in range(1, int(ws.ListObjects.Count) + 1):
                lo = ws.ListObjects(j)
                try:
                    st = lo.TableStyle.Name if lo.TableStyle is not None else None
                except (pywintypes.com_error, AttributeError):
                    st = None
                out.append({
                    "name": lo.Name, "sheet": ws.Name, "range": addr_of(lo.Range), "data_rows": int(lo.ListRows.Count),
                    "columns": [c.Name for c in lo.ListColumns], "style": st, "totals": bool(lo.ShowTotals),
                })
        return {"workbook": wb.Name, "tables": out}
    if act == "create":
        ws, rng = get_range(wb, sheet, cells, empty_means_used=False)
        if name:
            existing = [lo for i in range(1, wb.Worksheets.Count + 1) for lo in [wb.Worksheets(i).ListObjects]]
            for coll in existing:
                for j in range(1, int(coll.Count) + 1):
                    if coll(j).Name.lower() == name.lower():
                        raise ToolError(f"A table named '{name}' already exists.")
        # ListObjects.Add(SourceType=xlSrcRange(1), Source, LinkSource, XlListObjectHasHeaders)
        lo = ws.ListObjects.Add(1, rng, MISSING, 1 if has_headers else 2)
        if name:
            lo.Name = name
        lo.TableStyle = style or "TableStyleMedium2"
        if show_totals is not None:
            lo.ShowTotals = bool(show_totals)
        if show_filter_buttons is not None:
            lo.ShowAutoFilter = bool(show_filter_buttons)
        if banded_rows is not None:
            lo.ShowTableStyleRowStripes = bool(banded_rows)
        return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "table": lo.Name, "range": addr_of(lo.Range), "columns": [c.Name for c in lo.ListColumns]}
    if not name:
        raise ToolError("'name' of the table is required.")
    ws, lo = _find_table(wb, name)
    if act == "modify":
        applied = []
        if cells:
            _, new_rng = get_range(wb, ws.Name, cells, empty_means_used=False)
            lo.Resize(new_rng)
            applied.append(f"resized={addr_of(new_rng)}")
        if style:
            lo.TableStyle = style
            applied.append(f"style={style}")
        if show_totals is not None:
            lo.ShowTotals = bool(show_totals)
            applied.append(f"show_totals={show_totals}")
        if show_filter_buttons is not None:
            lo.ShowAutoFilter = bool(show_filter_buttons)
            applied.append(f"show_filter_buttons={show_filter_buttons}")
        if banded_rows is not None:
            lo.ShowTableStyleRowStripes = bool(banded_rows)
            applied.append(f"banded_rows={banded_rows}")
        if new_name:
            lo.Name = new_name
            applied.append(f"renamed={new_name}")
        if not applied:
            raise ToolError("Nothing to modify.")
        return {"ok": True, "workbook": wb.Name, "table": lo.Name, "range": addr_of(lo.Range), "applied": applied}
    if act == "delete":
        rng_addr = addr_of(lo.Range)
        lo.Delete()
        return {"ok": True, "workbook": wb.Name, "deleted_table": name, "range_cleared": rng_addr}
    if act == "to_range":
        lo.Unlist()
        return {"ok": True, "workbook": wb.Name, "table_converted_to_range": name}
    raise ToolError("action must be 'list', 'create', 'modify', 'delete' or 'to_range'.")


# ================================================================== сводные таблицы

PIVOT_FUNCS = {
    "sum": -4157, "count": -4112, "average": -4106, "avg": -4106, "max": -4136, "min": -4139,
    "product": -4149, "count_numbers": -4113, "stdev": -4155, "var": -4164,
}
PIVOT_SHOW_AS = {"percent_of_total": 8, "percent_of_row": 6, "percent_of_column": 7, "percent_of_parent": 12}
PIVOT_LAYOUTS = {"tabular": 1, "compact": 0, "outline": 2}


def _pivot_fields(pt, which) -> list[str]:
    try:
        return [f.Name for f in getattr(pt, which)]
    except (pywintypes.com_error, AttributeError):
        return []


def _pivot_field(pt, name: str):
    try:
        return pt.PivotFields(name)
    except pywintypes.com_error:
        available = [f.Name for f in pt.PivotFields()]
        raise ToolError(f"Field '{name}' is not a column of the source data. Available fields: {available}") from None


@office_tool("excel_analysis", "write", title="Create pivot table")
def excel_create_pivot_table(
    workbook: str,
    source: str,
    values: list[dict],
    rows: list[str] | None = None,
    columns: list[str] | None = None,
    filters: list[str] | None = None,
    source_sheet: str = "",
    dest_sheet: str = "",
    dest_cell: str = "A3",
    name: str = "",
    layout: str = "tabular",
    grand_totals: bool = True,
    style: str = "",
) -> dict:
    """Create a PivotTable ('сводная таблица') from a data block or an Excel table. The first row of the source must contain unique header names.

    Args:
        workbook: exact workbook name.
        source: the source data - a range with headers like 'A1:F500' (with `source_sheet`, or 'Data!A1:F500') or the name of an Excel table.
        values: aggregated fields, e.g. [{"field": "Revenue", "function": "sum", "number_format": "#,##0"}, {"field": "Order", "function": "count", "caption": "Orders"}]. function: sum|count|average|max|min|product|count_numbers|stdev|var; optional show_as: percent_of_total|percent_of_row|percent_of_column.
        rows: fields placed on the row axis, in order. columns: fields on the column axis. filters: report-filter fields.
        source_sheet: sheet of the source range ('' = active).
        dest_sheet: sheet for the pivot ('' = create a new sheet 'Pivot'; an unknown name creates that sheet).
        dest_cell: top-left cell of the pivot on the destination sheet.
        name: pivot table name. layout: tabular|compact|outline. grand_totals: show grand totals. style: e.g. 'PivotStyleMedium9'.
    """
    if not values:
        raise ToolError("'values' must contain at least one field to aggregate.")
    app, wb = pick_workbook(workbook)
    # источник: имя умной таблицы или диапазон
    src_obj = None
    for i in range(1, wb.Worksheets.Count + 1):
        coll = wb.Worksheets(i).ListObjects
        for j in range(1, int(coll.Count) + 1):
            if coll(j).Name.lower() == source.strip().lower():
                src_obj = coll(j).Range
    if src_obj is None:
        _, src_obj = get_range(wb, source_sheet, source, empty_means_used=False)
    # место назначения
    dest_name = dest_sheet or "Pivot"
    existing = {n.lower(): n for n in sheet_names(wb, any_type=True)}
    if dest_name.lower() in existing:
        if not dest_sheet:  # «Pivot» занят — берём свободное имя
            k = 2
            while f"{dest_name}{k}".lower() in existing:
                k += 1
            dest_name = f"{dest_name}{k}"
            validate_sheet_name(dest_name)
            dws = wb.Worksheets.Add(None, wb.Sheets(wb.Sheets.Count))
            dws.Name = dest_name
        else:
            dws = pick_sheet(wb, dest_sheet)
    else:
        validate_sheet_name(dest_name)
        dws = wb.Worksheets.Add(None, wb.Sheets(wb.Sheets.Count))
        dws.Name = dest_name
    dest = dws.Range(dest_cell)
    cache = wb.PivotCaches().Create(1, src_obj)  # Create(SourceType=xlDatabase, SourceData)
    try:
        pt = cache.CreatePivotTable(dest, name) if name else cache.CreatePivotTable(dest)
    except pywintypes.com_error as exc:
        raise ToolError("Excel could not create the pivot table (check that the source has unique, non-empty headers and that the destination is free): " + com.com_error_text(exc)) from None
    try:
        if layout.lower() in PIVOT_LAYOUTS:
            pt.RowAxisLayout(PIVOT_LAYOUTS[layout.lower()])
        for i, fname in enumerate(rows or [], start=1):
            pf = _pivot_field(pt, fname)
            pf.Orientation = 1
            pf.Position = i
        for i, fname in enumerate(columns or [], start=1):
            pf = _pivot_field(pt, fname)
            pf.Orientation = 2
            pf.Position = i
        for i, fname in enumerate(filters or [], start=1):
            pf = _pivot_field(pt, fname)
            pf.Orientation = 3
            pf.Position = i
        for v in values:
            fname = v.get("field")
            func = str(v.get("function", "sum")).lower()
            if func not in PIVOT_FUNCS:
                raise ToolError(f"Unknown function '{func}'. Use {sorted(PIVOT_FUNCS)}")
            pf = _pivot_field(pt, fname)
            caption = v.get("caption") or f"{func.capitalize()} of {fname}"
            df = pt.AddDataField(pf, caption, PIVOT_FUNCS[func])
            if v.get("number_format"):
                from .excel_pivot import _set_field_format

                _set_field_format(pt, df, v["number_format"])
            if v.get("show_as"):
                sa = str(v["show_as"]).lower()
                if sa not in PIVOT_SHOW_AS:
                    raise ToolError(f"show_as must be one of {sorted(PIVOT_SHOW_AS)}")
                df.Calculation = PIVOT_SHOW_AS[sa]
        pt.RowGrand = bool(grand_totals)
        pt.ColumnGrand = bool(grand_totals)
        if style:
            pt.TableStyle2 = style
    except ToolError:
        _drop_pivot(pt)
        raise
    except pywintypes.com_error as exc:
        _drop_pivot(pt)
        raise ToolError("Excel rejected a pivot setting: " + com.com_error_text(exc)) from None
    area = pt.TableRange1
    return {
        "ok": True, "workbook": wb.Name, "sheet": dws.Name, "pivot_table": pt.Name, "range": addr_of(area),
        "preview": preview(area, 120),
    }


def _drop_pivot(pt):
    try:
        pt.TableRange2.Clear()
    except pywintypes.com_error:
        pass


@office_tool("excel_analysis", "write", title="Manage pivot tables", read_actions=("list",), destructive=True)
def excel_manage_pivot_tables(workbook: str, action: str = "list", name: str = "", sheet: str = "") -> dict:
    """List, refresh or delete pivot tables.

    Args:
        workbook: exact workbook name.
        action: 'list', 'refresh' (one by name, or all when name is empty) or 'delete' (clears the pivot area).
        name: pivot table name for refresh/delete.
        sheet: restrict listing/search to one sheet.
    """
    app, wb = pick_workbook(workbook)
    act = action.lower()
    sheets = [pick_sheet(wb, sheet)] if sheet else [wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1)]
    found = []
    for ws in sheets:
        pts = ws.PivotTables()
        for j in range(1, int(pts.Count) + 1):
            pt = pts(j)
            found.append((ws, pt))
    if act == "list":
        items = []
        for ws, pt in found:
            try:
                src = pt.SourceData
            except pywintypes.com_error:
                src = None
            items.append({
                "name": pt.Name, "sheet": ws.Name, "range": addr_of(pt.TableRange1), "source": src,
                "row_fields": _pivot_fields(pt, "RowFields"), "column_fields": _pivot_fields(pt, "ColumnFields"),
                "page_fields": _pivot_fields(pt, "PageFields"), "data_fields": _pivot_fields(pt, "DataFields"),
            })
        return {"workbook": wb.Name, "pivot_tables": items}
    targets = [(ws, pt) for ws, pt in found if not name or pt.Name.lower() == name.lower()]
    if name and not targets:
        raise ToolError(f"Pivot table '{name}' not found. Existing: {[pt.Name for _, pt in found]}")
    if act == "refresh":
        for _, pt in targets:
            pt.RefreshTable()
        return {"ok": True, "workbook": wb.Name, "refreshed": [pt.Name for _, pt in targets]}
    if act == "delete":
        if not name:
            raise ToolError("'name' is required for delete.")
        _, pt = targets[0]
        area = addr_of(pt.TableRange2)
        pt.TableRange2.Clear()
        return {"ok": True, "workbook": wb.Name, "deleted": name, "cleared_range": area}
    raise ToolError("action must be 'list', 'refresh' or 'delete'.")


# ================================================================== диаграммы

CHART_TYPES = {
    "column": 51, "column_stacked": 52, "column_100": 53, "bar": 57, "bar_stacked": 58, "bar_100": 59,
    "line": 4, "line_markers": 65, "line_stacked": 63, "pie": 5, "pie_exploded": 69, "doughnut": -4120,
    "area": 1, "area_stacked": 76, "scatter": -4169, "scatter_lines": 74, "scatter_smooth": 72,
    "radar": -4151, "radar_markers": 81, "bubble": 15, "surface": 83,
}
LEGEND_POS = {"bottom": -4107, "top": -4160, "left": -4131, "right": -4152}
NO_AXES = {"pie", "pie_exploded", "doughnut"}


def _chart_type_code(chart_type: str) -> int:
    t = chart_type.lower().replace("-", "_").replace(" ", "_")
    if t not in CHART_TYPES:
        raise ToolError(f"chart_type must be one of {sorted(CHART_TYPES)}")
    return CHART_TYPES[t]


def _find_chart(wb, chart_name: str, sheet: str = ""):
    sheets = [pick_sheet(wb, sheet)] if sheet else [wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1)]
    seen = []
    for ws in sheets:
        cos = ws.ChartObjects()
        for j in range(1, int(cos.Count) + 1):
            co = cos(j)
            seen.append(co.Name)
            if co.Name.lower() == chart_name.lower():
                return ws, co
    raise ToolError(f"Chart '{chart_name}' not found. Charts: {seen}")


def _chart_summary(co) -> dict:
    ch = co.Chart
    info = {"name": co.Name, "chart_type": int(ch.ChartType), "left": float(co.Left), "top": float(co.Top), "width": float(co.Width), "height": float(co.Height)}
    info["title"] = ch.ChartTitle.Text if ch.HasTitle else None
    series = []
    sc = ch.SeriesCollection()
    for k in range(1, int(sc.Count) + 1):
        s = sc(k)
        series.append({"index": k, "name": s.Name, "formula": s.Formula})
    info["series"] = series
    return info


def _style_chart(ch, kind_key: str, title, x_title, y_title, legend, data_labels, y_format=None, y_min=None, y_max=None, colors=None, label_format=None):
    if title is not None:
        if title == "":
            ch.HasTitle = False
        else:
            ch.HasTitle = True
            ch.ChartTitle.Text = title
    if kind_key not in NO_AXES:
        if x_title is not None:
            ax = ch.Axes(1)
            ax.HasTitle = bool(x_title)
            if x_title:
                ax.AxisTitle.Text = x_title
        if y_title is not None:
            ax = ch.Axes(2)
            ax.HasTitle = bool(y_title)
            if y_title:
                ax.AxisTitle.Text = y_title
    if legend is not None:
        if legend.lower() == "none":
            ch.HasLegend = False
        else:
            if legend.lower() not in LEGEND_POS:
                raise ToolError("legend must be bottom|top|left|right|none")
            ch.HasLegend = True
            ch.Legend.Position = LEGEND_POS[legend.lower()]
    if data_labels is not None:
        sc = ch.SeriesCollection()
        for k in range(1, int(sc.Count) + 1):
            sc(k).HasDataLabels = bool(data_labels)
    app = ch.Application
    if kind_key not in NO_AXES and (y_format or y_min is not None or y_max is not None):
        ax = ch.Axes(2)
        if y_format:
            ax.TickLabels.NumberFormat = number_format_for_write(app, y_format, NUMBER_FORMATS)
        if y_min is not None:
            ax.MinimumScale = float(y_min)
        if y_max is not None:
            ax.MaximumScale = float(y_max)
    if label_format:
        sc = ch.SeriesCollection()
        for k in range(1, int(sc.Count) + 1):
            try:
                sc(k).DataLabels().NumberFormat = number_format_for_write(app, label_format, NUMBER_FORMATS)
            except pywintypes.com_error:
                pass  # у ряда нет подписей данных
    if colors:
        sc = ch.SeriesCollection()
        for k, col in enumerate(colors, start=1):
            if k > int(sc.Count):
                break
            bgr = parse_color(col)
            s_k = sc(k)
            try:
                s_k.Format.Fill.ForeColor.RGB = bgr
            except pywintypes.com_error:
                pass
            try:
                s_k.Format.Line.ForeColor.RGB = bgr
            except pywintypes.com_error:
                pass


@office_tool("excel_analysis", "write", title="Create chart")
def excel_create_chart(
    workbook: str,
    source: str,
    sheet: str = "",
    chart_type: str = "column",
    title: str = "",
    anchor_cell: str = "",
    width: float = 480,
    height: float = 288,
    x_axis_title: str = "",
    y_axis_title: str = "",
    legend: str = "bottom",
    data_labels: bool = False,
    series_in_rows: bool | None = None,
    name: str = "",
    value_axis_number_format: str = "",
    value_axis_min: float | None = None,
    value_axis_max: float | None = None,
    series_colors: list[str] | None = None,
    data_label_number_format: str = "",
) -> dict:
    """Create a chart from a data block. Include the header row and the label column in `source`: the first column becomes the category axis and each other column a series (use series_in_rows=true if the series are laid out in rows).

    Args:
        workbook: exact workbook name.
        source: data range such as 'A1:D6' (or 'Sheet2!A1:D6'), or the name of a pivot table to get a pivot chart.
        sheet: sheet of the source and where the chart is placed ('' = active).
        chart_type: column | column_stacked | column_100 | bar | bar_stacked | line | line_markers | pie | doughnut | area | area_stacked | scatter | scatter_lines | scatter_smooth | radar | bubble.
        title: chart title ('' = no title).
        anchor_cell: top-left cell for the chart (default: two columns right of the data).
        width, height: size in points.
        x_axis_title, y_axis_title: axis captions ('' = none).
        legend: bottom | top | left | right | none.
        data_labels: show value labels.
        series_in_rows: force series orientation (null = automatic).
        name: chart object name.
        value_axis_number_format: number format of the value axis, e.g. '#,##0' (no decimals). value_axis_min, value_axis_max: axis limits.
        series_colors: one color per series, e.g. ['#1F4E78', '#C00000']. data_label_number_format: format of the value labels.
    """
    app, wb = pick_workbook(workbook)
    pivot_hit = None
    for pws in [wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1)]:
        pts = pws.PivotTables()
        for j in range(1, int(pts.Count) + 1):
            if pts(j).Name.lower() == source.strip().lower():
                pivot_hit = (pws, pts(j))
    if pivot_hit:  # источник — сводная таблица: получится сводная диаграмма
        ws, src = pivot_hit[0], pivot_hit[1].TableRange1
    else:
        ws, src = get_range(wb, sheet, source, empty_means_used=False)
    r1, c1, r2, c2 = bounds(src)
    key = chart_type.lower().replace("-", "_").replace(" ", "_")
    code = _chart_type_code(key)
    anchor = ws.Range(anchor_cell) if anchor_cell else ws.Range(a1_cell(r1, min(c2 + 2, 16384)))
    left, top = float(anchor.Left), float(anchor.Top)
    try:
        shape = ws.Shapes.AddChart2(-1, code, left, top, float(width), float(height))  # AddChart2(Style, Type, Left, Top, W, H)
        ch = shape.Chart
    except (pywintypes.com_error, AttributeError):
        co = ws.ChartObjects().Add(left, top, float(width), float(height))
        shape = co
        ch = co.Chart
        ch.ChartType = code
    if series_in_rows is None:
        ch.SetSourceData(src)
    else:
        ch.SetSourceData(src, 1 if series_in_rows else 2)  # PlotBy: xlRows=1, xlColumns=2
    if name:
        shape.Name = name
    _style_chart(ch, key, title, x_axis_title, y_axis_title, legend, data_labels, value_axis_number_format or None, value_axis_min, value_axis_max, series_colors, data_label_number_format or None)
    series = []
    sc = ch.SeriesCollection()
    for k in range(1, int(sc.Count) + 1):
        series.append(sc(k).Name)
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "chart": shape.Name, "type": key, "series": series, "position": {"left": left, "top": top, "width": float(width), "height": float(height)}}


@office_tool("excel_analysis", "write", title="Manage charts", unstructured=True, read_actions=("list", "export_image"), destructive=True, file_args=("export_path",))
def excel_manage_charts(
    workbook: str,
    action: str = "list",
    chart_name: str = "",
    sheet: str = "",
    chart_type: str = "",
    title: str | None = None,
    x_axis_title: str | None = None,
    y_axis_title: str | None = None,
    legend: str | None = None,
    data_labels: bool | None = None,
    source: str = "",
    series_overrides: list[dict] | None = None,
    new_series: dict | None = None,
    width: float | None = None,
    height: float | None = None,
    anchor_cell: str = "",
    export_path: str = "",
    overwrite: bool = False,
    value_axis_number_format: str | None = None,
    value_axis_min: float | None = None,
    value_axis_max: float | None = None,
    series_colors: list[str] | None = None,
    data_label_number_format: str | None = None,
) -> list | dict:
    """List, restyle, move, delete or export charts.

    Args:
        workbook: exact workbook name.
        action: 'list' | 'update' | 'delete' | 'export_image' (returns the chart as a PNG image; also saved to export_path when given).
        chart_name: chart name from 'list' (required except for list).
        sheet: sheet to search ('' = all sheets).
        chart_type, title, x_axis_title, y_axis_title, legend, data_labels: new settings for 'update' ('' title removes it).
        source: new data range for 'update'.
        series_overrides: combo-chart control, e.g. [{"index": 2, "chart_type": "line", "secondary_axis": true}].
        new_series: add a series, e.g. {"name": "Plan", "values": "B2:B6", "x_values": "A2:A6"} (ranges on the chart's sheet or 'Sheet!A1:A5').
        width, height, anchor_cell: resize/move the chart.
        export_path: optional PNG path for export_image (refuses to replace an existing file unless overwrite=true; not available in read-only mode).
        overwrite: allow export_path to replace an existing file.
        value_axis_number_format ('#,##0'), value_axis_min, value_axis_max, series_colors (one per series), data_label_number_format: axis and look settings for 'update'.
    """
    app, wb = pick_workbook(workbook)
    act = action.lower()
    if act == "list":
        sheets = [pick_sheet(wb, sheet)] if sheet else [wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1)]
        items = []
        for ws in sheets:
            cos = ws.ChartObjects()
            for j in range(1, int(cos.Count) + 1):
                s = _chart_summary(cos(j))
                s["sheet"] = ws.Name
                items.append(s)
        return {"workbook": wb.Name, "charts": items}
    if not chart_name:
        raise ToolError("'chart_name' is required.")
    ws, co = _find_chart(wb, chart_name, sheet)
    ch = co.Chart
    if act == "delete":
        co.Delete()
        return {"ok": True, "workbook": wb.Name, "deleted": chart_name}
    if act == "export_image":
        dest = None
        if export_path:
            if config.SETTINGS.readonly:
                raise ToolError("export_path writes a file, which read-only mode does not allow; omit it to just get the image back.")
            dest = check_path(export_path, "write", {"png"})
            if os.path.exists(dest) and not overwrite:
                raise ToolError(f"File already exists: {dest}. Pass overwrite=true to replace it.")
        tmpdir = tempfile.mkdtemp(prefix="office_live_")
        png = os.path.join(tmpdir, "chart.png")
        try:
            ch.Export(png, "PNG")
            with open(png, "rb") as f:
                data = f.read()
            if dest:
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                shutil.copyfile(png, dest)
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
        return [json.dumps({"ok": True, "chart": co.Name, "png_bytes": len(data), "saved_to": export_path or None}, ensure_ascii=False), Image(data=data, format="png")]
    if act != "update":
        raise ToolError("action must be 'list', 'update', 'delete' or 'export_image'.")
    applied = []
    key = None
    if chart_type:
        key = chart_type.lower().replace("-", "_").replace(" ", "_")
        ch.ChartType = _chart_type_code(key)
        applied.append(f"chart_type={key}")
    if source:
        _, src = get_range(wb, ws.Name, source, empty_means_used=False, allow_other_sheet=True)
        ch.SetSourceData(src)
        applied.append(f"source={addr_of(src)}")
    _style_chart(ch, key or "", title, x_axis_title, y_axis_title, legend, data_labels, value_axis_number_format, value_axis_min, value_axis_max, series_colors, data_label_number_format)
    for k, v in (("title", title), ("x_axis_title", x_axis_title), ("y_axis_title", y_axis_title), ("legend", legend), ("data_labels", data_labels)):
        if v is not None:
            applied.append(f"{k}={v}")
    for k, v in (("value_axis_number_format", value_axis_number_format), ("value_axis_min", value_axis_min), ("value_axis_max", value_axis_max), ("series_colors", series_colors), ("data_label_number_format", data_label_number_format)):
        if v not in (None, ""):
            applied.append(f"{k}={v}")
    for ov in series_overrides or []:
        idx = int(ov.get("index", 0))
        s = ch.SeriesCollection(idx)
        if ov.get("chart_type"):
            s.ChartType = _chart_type_code(ov["chart_type"])
        if ov.get("secondary_axis") is not None:
            s.AxisGroup = 2 if ov["secondary_axis"] else 1
        applied.append(f"series{idx}")
    if new_series:
        s = ch.SeriesCollection().NewSeries()
        if new_series.get("name"):
            s.Name = new_series["name"]
        if new_series.get("values"):
            s.Values = get_range(wb, ws.Name, new_series["values"], empty_means_used=False, allow_other_sheet=True)[1]
        if new_series.get("x_values"):
            s.XValues = get_range(wb, ws.Name, new_series["x_values"], empty_means_used=False, allow_other_sheet=True)[1]
        applied.append("new_series")
    if width:
        co.Width = float(width)
        applied.append(f"width={width}")
    if height:
        co.Height = float(height)
        applied.append(f"height={height}")
    if anchor_cell:
        a = ws.Range(anchor_cell)
        co.Left, co.Top = float(a.Left), float(a.Top)
        applied.append(f"moved_to={anchor_cell}")
    if not applied:
        raise ToolError("Nothing to update.")
    return {"ok": True, "workbook": wb.Name, "chart": co.Name, "applied": applied, "now": _chart_summary(co)}


# ================================================================== профилирование


_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}( \d{2}:\d{2}:\d{2})?$")
_TIME_RE = re.compile(r"^\d{2}:\d{2}:\d{2}$")
_ERRORS = frozenset(EXCEL_ERRORS.values())


def _kind(v) -> str:
    if v is None or v == "":
        return "empty"
    if isinstance(v, bool):
        return "bool"
    if isinstance(v, (int, float)):
        return "number"
    if isinstance(v, str):
        if v in _ERRORS:
            return "error"
        if _DATE_RE.match(v):
            return "date"
        if _TIME_RE.match(v):
            return "time"
        return "text"
    return "other"


def _read_block(app, ws, rng, max_cells: int):
    """Блок значений (list[list]) с обрезкой по max_cells."""
    clipped = clip_to_used(app, ws, rng)
    if clipped is None:
        return None, None, False
    r1, c1, r2, c2 = bounds(clipped)
    cols = c2 - c1 + 1
    if cols > max_cells:
        raise ToolError(f"The range has {cols} columns, more than max_cells={max_cells}. Narrow the columns or raise max_cells.")
    keep_rows = max(1, max_cells // cols)
    truncated = (r2 - r1 + 1) > keep_rows
    if truncated:
        clipped = sub_range(ws, r1, c1, r1 + keep_rows - 1, c2)
    return clipped, (r1, c1), truncated


@office_tool("excel_analysis", "read", title="Profile data")
def excel_profile_range(
    workbook: str = "",
    sheet: str = "",
    cells: str = "",
    has_header: bool = True,
    top_values: int = 5,
    max_cells: int = 200000,
) -> dict:
    """Statistical profile of a data block, column by column: types, empty/unique counts, min/max/mean/median/sum for numbers, date range, most frequent values, duplicates, formula counts. Use it to understand unfamiliar data before analysing or cleaning it.

    Args:
        workbook: exact name or '' for the active workbook.
        sheet: sheet name or '' for active.
        cells: data block ('' = the whole used range).
        has_header: first row holds column names.
        top_values: how many most-frequent values to list per column.
        max_cells: safety cap; larger blocks are profiled on their first rows.
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells)
    block, origin, truncated = _read_block(app, ws, rng, int(max_cells))
    if block is None or sheet_is_empty(app, ws):
        return {"workbook": wb.Name, "sheet": ws.Name, "rows": 0, "columns": []}
    r1, c1 = origin
    grid = to_grid(block.Value)
    formulas = to_grid(block.Formula)
    header = grid[0] if has_header else None
    body = grid[1:] if has_header else grid
    fbody = formulas[1:] if has_header else formulas
    ncols = len(grid[0])
    columns = []
    for j in range(ncols):
        vals = [row[j] for row in body]
        kinds = Counter(_kind(v) for v in vals)
        present = [v for v in vals if _kind(v) != "empty"]
        label = header[j] if header and header[j] not in (None, "") else col_letter(c1 + j)
        info = {
            "column": col_letter(c1 + j), "name": label, "non_empty": len(present), "empty": kinds.get("empty", 0),
            "unique": len({str(v).lower() if isinstance(v, str) else v for v in present}),
            "types": {k: n for k, n in kinds.items() if k != "empty"},
            "formula_cells": sum(1 for row in fbody if isinstance(row[j], str) and row[j].startswith("=")),
        }
        nums = [v for v in present if _kind(v) == "number"]
        if nums:
            info["numeric"] = {
                "min": min(nums), "max": max(nums), "sum": round(sum(nums), 6), "mean": round(statistics.fmean(nums), 6),
                "median": statistics.median(nums), "zeros": sum(1 for v in nums if v == 0), "negatives": sum(1 for v in nums if v < 0),
            }
            if len(nums) > 1:
                info["numeric"]["stdev"] = round(statistics.pstdev(nums), 6)
        dates = sorted(str(v) for v in present if _kind(v) in ("date", "time"))
        if dates:
            info["dates"] = {"min": dates[0], "max": dates[-1]}
        texts = [v for v in present if _kind(v) == "text"]
        if texts:
            info["text"] = {"min_length": min(len(t) for t in texts), "max_length": max(len(t) for t in texts),
                            "with_edge_spaces": sum(1 for t in texts if t != t.strip())}
        counts = Counter(str(v) for v in present)
        dup = sum(1 for n in counts.values() if n > 1)
        info["duplicated_values"] = dup
        info["top_values"] = [{"value": v, "count": n} for v, n in counts.most_common(max(0, int(top_values)))]
        columns.append(info)
    seen, dup_rows, empty_rows = set(), 0, 0
    for row in body:
        if all(_kind(v) == "empty" for v in row):
            empty_rows += 1
            continue
        key = tuple(row)
        if key in seen:
            dup_rows += 1
        seen.add(key)
    return {
        "workbook": wb.Name, "sheet": ws.Name, "range": addr_of(block), "header": bool(has_header),
        "data_rows": len(body), "columns_count": ncols, "empty_rows": empty_rows, "duplicate_rows": dup_rows,
        "truncated": truncated, "columns": columns,
    }


@office_tool("excel_analysis", "read", title="Find data issues")
def excel_find_issues(workbook: str = "", sheet: str = "", cells: str = "", max_cells: int = 200000, limit: int = 15, has_header: bool = True) -> dict:
    """Audit a sheet/range for problems: error values, numbers stored as text, inconsistent formulas down a column, hard-coded numbers among formulas, stray spaces, blank rows/headers, duplicate headers, mixed types in a column, merged cells.

    Args:
        workbook: exact name or '' for the active workbook.
        sheet: sheet name or '' for active.
        cells: range to audit ('' = whole used range).
        max_cells: safety cap on analysed cells.
        limit: max examples listed per issue type.
        has_header: the first row of the audited range is a header row (false when `cells` starts at the data: header checks are skipped).
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells)
    block, origin, truncated = _read_block(app, ws, rng, int(max_cells))
    if block is None or sheet_is_empty(app, ws):
        return {"workbook": wb.Name, "sheet": ws.Name, "issues": [], "summary": "The range is empty."}
    r1, c1 = origin
    values = to_grid(block.Value)
    r1c1 = to_grid(block.FormulaR1C1)
    nrows, ncols = len(values), len(values[0])
    issues: dict[str, dict] = {}

    def add(kind, severity, cell, detail):
        entry = issues.setdefault(kind, {"type": kind, "severity": severity, "count": 0, "examples": []})
        entry["count"] += 1
        if len(entry["examples"]) < limit:
            entry["examples"].append({"cell": cell, "detail": detail})

    for j in range(ncols):
        col_cells = [(i, values[i][j], r1c1[i][j]) for i in range(nrows)]
        formula_cells = [(i, f) for i, _, f in col_cells if isinstance(f, str) and f.startswith("=")]
        for i, v, f in col_cells:
            cell = a1_cell(r1 + i, c1 + j)
            if isinstance(v, str):
                if v in _ERRORS:
                    add("error_value", "high", cell, v)
                elif v != v.strip() and v.strip():
                    add("stray_spaces", "low", cell, repr(v[:40]))
                elif smart_number(v) is not None and not (isinstance(f, str) and f.startswith("=")):
                    add("number_stored_as_text", "medium", cell, repr(v[:40]))
        if len(formula_cells) >= 3:
            counts = Counter(f for _, f in formula_cells)
            main, main_n = counts.most_common(1)[0]
            if main_n < len(formula_cells):
                for i, f in formula_cells:
                    if f != main:
                        add("inconsistent_formula", "high", a1_cell(r1 + i, c1 + j), f"{f}  (column majority: {main})")
            if len(formula_cells) / max(1, sum(1 for _, v, _ in col_cells if v not in (None, ""))) >= 0.6:
                for i, v, f in col_cells:
                    if isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(f, str) and f.startswith("=")):
                        add("hardcoded_number_among_formulas", "medium", a1_cell(r1 + i, c1 + j), f"value {v} in a column of formulas")
        kinds = Counter(_kind(v) for _, v, _ in col_cells[1 if has_header else 0:] if _kind(v) not in ("empty", "error"))
        if len(kinds) > 1 and min(kinds.values()) <= max(2, 0.1 * sum(kinds.values())):
            add("mixed_types_in_column", "medium", col_letter(c1 + j), dict(kinds))
    # отдельные блоки данных рядом (разделены пустыми столбцами) проверяем независимо: заголовки одного блока не дублируют другой
    has_data = [any(values[i][j] not in (None, "") for i in range(nrows)) for j in range(ncols)]
    blocks, start = [], None
    for j, flag in enumerate(has_data + [False]):
        if flag and start is None:
            start = j
        elif not flag and start is not None:
            blocks.append((start, j - 1))
            start = None
    for b0, b1 in blocks if has_header else []:
        headers = values[0][b0:b1 + 1]
        named = [str(h).strip().lower() for h in headers if h not in (None, "")]
        for h, n in Counter(named).items():
            if n > 1:
                add("duplicate_header", "medium", col_letter(c1 + b0), f"'{h}' appears {n} times")
        for jj, h in enumerate(headers):
            j = b0 + jj
            if h in (None, "") and any(values[i][j] not in (None, "") for i in range(1, nrows)):
                add("blank_header", "medium", a1_cell(r1, c1 + j), "column has data but no header")
    data_rows = [i for i in range(nrows) if any(v not in (None, "") for v in values[i])]
    if data_rows:
        filled = set(data_rows)  # множество: проверка членства за O(1), а не O(n) на каждую строку
        for i in range(data_rows[0], data_rows[-1] + 1):
            if i not in filled:
                add("blank_row_inside_data", "low", f"{r1 + i}:{r1 + i}", "empty row between data rows")
    try:
        merged = block.MergeCells
        if merged is None or merged is True:
            add("merged_cells", "low", addr_of(block), "range contains merged cells (they break sorting, filtering and pivots)")
    except pywintypes.com_error:
        pass
    ordered = sorted(issues.values(), key=lambda e: ({"high": 0, "medium": 1, "low": 2}[e["severity"]], -e["count"]))
    return {
        "workbook": wb.Name, "sheet": ws.Name, "range": addr_of(block), "truncated": truncated,
        "summary": f"{sum(e['count'] for e in ordered)} findings of {len(ordered)} types" if ordered else "No issues found.",
        "issues": ordered,
    }
