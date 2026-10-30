"""Excel: умные таблицы, сводные таблицы, диаграммы, профилирование данных, поиск проблем."""

import datetime
import math
import re
import statistics
from collections import Counter

import pythoncom
import pywintypes

from . import com, undo
from . import xl_common as xl
from .errors import PartialChangeError, ToolError
from .registry import office_tool
from .util import EXCEL_ERRORS, a1_cell, a1_range, col_letter, col_number, norm_value, smart_number, split_sheet_ref, to_grid
from .xl_common import (
    addr_of,
    bounded_range,
    bounds,
    clip_to_used,
    get_range,
    pick_sheet,
    pick_workbook,
    preview,
    sheet_is_empty,
    sheet_names,
    sub_range,
    validate_sheet_name,
)

MISSING = pythoncom.Missing


@office_tool("excel_analysis", "read", title="Trace formula")
def excel_trace_formula(workbook: str = "", sheet: str = "", cell: str = "", direction: str = "precedents",
                        depth: int = 1, max_nodes: int = 200, per_match_links: bool = False) -> dict:
    """Trace formula references without selection changes, sheet activation, arrows or evaluation. Edges always point from input to consumer. Never treat complete=false as a full graph.

    Args:
        workbook, sheet: target workbook/sheet; empty selects the active one.
        cell: required single cell.
        direction: precedents (inputs) or dependents (consumers).
        depth: 1..5, default 1; an unexpanded edge sets depth_limit.
        max_nodes: 1..2000 including the root, default 200.
        per_match_links: true returns links for individual nodes (up to 20); default returns one workbook link.

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
    per_match_links: bool = False,
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
        per_match_links: true returns links for individual differences (up to 20); default returns one link to the first workbook with a difference.

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

from .excel_charts import (
    CHART_TYPES as CHART_TYPES,
    LEGEND_POS as LEGEND_POS,
    NO_AXES as NO_AXES,
    _chart_type_code as _chart_type_code,
    _find_chart as _find_chart,
    _chart_summary as _chart_summary,
    _style_chart as _style_chart,
    _chart_pivot as _chart_pivot,
    _chart_selection as _chart_selection,
    excel_create_chart as excel_create_chart,
    _PNG_SIGNATURE as _PNG_SIGNATURE,
    _export_chart_png as _export_chart_png,
    excel_manage_charts as excel_manage_charts,
)


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
def excel_find_issues(workbook: str = "", sheet: str = "", cells: str = "", max_cells: int = 200000, limit: int = 15, has_header: bool = True,
                      per_match_links: bool = False) -> dict:
    """Audit a sheet/range for problems: error values, numbers stored as text, inconsistent formulas down a column, hard-coded numbers among formulas, stray spaces, blank rows/headers, duplicate headers, mixed types in a column, merged cells.

    Args:
        workbook: exact name or '' for the active workbook.
        sheet: sheet name or '' for active.
        cells: range to audit ('' = whole used range).
        max_cells: safety cap on analysed cells.
        limit: max examples listed per issue type.
        has_header: the first row of the audited range is a header row (false when `cells` starts at the data: header checks are skipped).
        per_match_links: true returns links for individual issue examples (up to 20); default returns one workbook link.
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


def _finite(value, name):
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise ToolError(f"{name} must be a finite number.")


def scenario_state(item):
    values = item.Values  # optional-index PROPERTYGET, not a method
    return {"name": item.Name, "changing_cells": xl.addr_of(item.ChangingCells),
            "values": list(values) if isinstance(values, (tuple, list)) else [values], "comment": item.Comment,
            "locked": bool(item.Locked), "hidden": bool(item.Hidden)}


def scenario_items(ws):
    collection = ws.Scenarios()
    return [collection.Item(i) for i in range(1, int(collection.Count) + 1)]


def _what_if_union(app, wb, ws, address):
    """changing_cells/result_cells: one range or a comma-separated union on this sheet (get_range takes one area)."""
    parts = [part.strip() for part in str(address).split(",") if part.strip()]
    if not parts:
        raise ToolError("Cells are required.")
    result = None
    for part in parts:
        _, rng = xl.get_range(wb, ws.Name, part, empty_means_used=False)
        if rng.Worksheet.Name != ws.Name:
            raise ToolError("All cells must be on the scenario sheet.")
        result = rng if result is None else app.Union(result, rng)
    return result


def _what_if_ranges(rng, ws, wb, limit):
    if int(rng.Count) > limit:
        raise ToolError(f"Select at most {limit} cells.")
    areas = [rng.Areas.Item(i) for i in range(1, int(rng.Areas.Count) + 1)]
    for area in areas:
        xl.validate_rectangle(area, ws, wb)
    return areas


def _what_if_writable(app, ws, rng):
    if rng.MergeCells is not False or rng.HasArray is not False or rng.HasSpill is not False:
        raise ToolError("Merged cells, arrays and spills are not supported.")
    xl.reject_pivots(app, ws, rng)
    for table in ws.ListObjects:
        if app.Intersect(rng, table.Range) is not None:
            raise ToolError("Range intersects an Excel table.")


@office_tool("excel_analysis", "write", title="What-if analysis", read_actions=("list_scenarios",), destructive=True)
def excel_what_if(workbook: str, sheet: str, action: str = "list_scenarios", cells: str = "",
                  row_input: str = "", column_input: str = "", name: str = "", changing_cells: str = "",
                  values: list | None = None, comment: str = "", result_cells: str = "") -> dict:
    """Build one/two-variable data tables or add/show/delete/list scenarios and create a standard summary sheet.

    cells: entire data-table rectangle including headers; inputs must be outside it on the same sheet.
    One-variable tables have formulas down the first column (row_input) or across the first row (column_input).
    Two-variable tables have a formula in the upper-left corner. Results must be empty; office_undo clears the array.
    Scenarios accept up to 32 changing_cells (including unions), one finite number or text value per cell.
    result_cells: required formula cells for scenario_summary. Only list_scenarios is available in readonly.
    """
    if action not in {"data_table", "list_scenarios", "add_scenario", "show_scenario", "delete_scenario", "scenario_summary"}:
        raise ToolError("Unknown what-if action.")
    app, wb = xl.pick_workbook(workbook)
    ws = xl.pick_sheet(wb, sheet)
    base = {"workbook": wb.Name, "sheet": ws.Name, "action": action}
    if action == "list_scenarios":
        return {**base, "scenarios": [scenario_state(s) for s in scenario_items(ws)]}
    if bool(ws.ProtectContents) or bool(ws.ProtectScenarios):
        raise ToolError("Unprotect the sheet before changing what-if analysis.")
    plan = {"_what_sheet": ws.Name}
    if action == "data_table":
        from .formula_trace import same_sheet_dependency

        _, block = xl.get_range(wb, ws.Name, cells, empty_means_used=False)
        xl.validate_rectangle(block, ws, wb)
        _what_if_ranges(block, ws, wb, 20000)
        _what_if_writable(app, ws, block)
        r, c, end, right = xl.bounds(block)
        if r == end or c == right or not (row_input or column_input):
            raise ToolError("A data table needs headers, results and at least one input cell.")
        inputs = []
        for address in (row_input, column_input):
            if not address:
                inputs.append(None)  # None, not Missing: Missing in the first position drops the next argument (com.py)
                continue
            _, cell = xl.get_range(wb, ws.Name, address, empty_means_used=False)
            xl.validate_rectangle(cell, ws, wb)
            if int(cell.Count) != 1 or app.Intersect(block, cell) is not None:
                raise ToolError("Each input must be one cell outside the data table on the same sheet.")
            _what_if_writable(app, ws, cell)
            inputs.append(cell)
        if row_input and column_input and xl.addr_of(inputs[0]) == xl.addr_of(inputs[1]):
            raise ToolError("Row and column inputs must differ.")
        formulas = ([ws.Cells(r, c)] if row_input and column_input else
                    [ws.Cells(i, c) for i in range(r + 1, end + 1)] if row_input else
                    [ws.Cells(r, j) for j in range(c + 1, right + 1)])
        verified, reasons = True, set()
        for formula in formulas:
            if not bool(formula.HasFormula):
                raise ToolError("Every result header must contain a formula.")
            for input_cell in [inputs[i] for i, a in enumerate((row_input, column_input)) if a]:
                depends, complete, why = same_sheet_dependency(app, wb, formula, input_cell)
                if complete and not depends:
                    raise ToolError("Result formula does not depend on an input cell.")
                verified = verified and complete and depends
                reasons.update(why)
        result = xl.sub_range(ws, r + 1, c + 1, end, right)
        if any(v is not None for row in to_grid(result.Value2) for v in row) or result.HasFormula is not False:
            raise ToolError("The result area must be empty.")
        mode = int(app.Calculation)
        note = ("Use F9 or excel_calculate to recalculate data tables." if mode in {-4135, 2} else "")  # xlCalculationManual, xlCalculationSemiautomatic
        plan.update(_what_results=xl.addr_of(result))
        base.update(rows=end - r, columns=right - c, results=xl.addr_of(result), dependency_verified=bool(verified),
                    dependency_reasons=sorted(reasons), calculation_mode=mode, calculation_note=note)
    else:
        scenarios = scenario_items(ws)
        existing = next((s for s in scenarios if s.Name.casefold() == name.casefold()), None)
        if action == "add_scenario":
            if not name.strip() or len(name) > 255 or existing is not None:
                raise ToolError("Scenario name must be nonempty, unique and at most 255 characters.")
            changing = _what_if_union(app, wb, ws, changing_cells)
            _what_if_ranges(changing, ws, wb, 32)
            if values is None or len(values) != int(changing.Count):
                raise ToolError("Supply one value per changing cell (at most 32).")
            for value in values:
                if not isinstance(value, str):
                    _finite(value, "scenario value")
                elif len(value) > 255:
                    raise ToolError("Scenario text values must be at most 255 characters.")
            plan.update(_what_name=name, _what_before=None)
        elif action in {"show_scenario", "delete_scenario"}:
            if existing is None:
                raise ToolError("Scenario not found.")
            before = scenario_state(existing)
            changing = existing.ChangingCells
            areas = _what_if_ranges(changing, ws, wb, 32)
            if action == "show_scenario":
                for area in areas:
                    _what_if_writable(app, ws, area)
                plan.update(_what_cells=[xl.addr_of(a) for a in areas])
            else:
                plan.update(_what_name=existing.Name, _what_before=before)
        else:
            if not scenarios or bool(wb.ProtectStructure):
                raise ToolError("Summary needs scenarios and an unprotected workbook structure.")
            result = _what_if_union(app, wb, ws, result_cells)
            areas = _what_if_ranges(result, ws, wb, 32)
            if any(a.HasFormula is not True for a in areas):
                raise ToolError("result_cells must contain formulas.")
            sheets_before = xl.sheet_names(wb, any_type=True)
        if action != "scenario_summary":
            base["name"] = name
    undo.require_undo("workbook", app, wb, plan=plan)
    try:
        if action == "data_table":
            block.Table(*inputs)
            base["preview"] = xl.preview(result, 20)
        elif action == "add_scenario":
            ws.Scenarios().Add(name, changing, values, comment, False, False)
        elif action == "show_scenario":
            existing.Show()
        elif action == "delete_scenario":
            existing.Delete()
        else:
            ws.Scenarios().CreateSummary(1, result)  # xlStandardSummary
            base["created_sheets"] = [n for n in xl.sheet_names(wb, any_type=True) if n not in sheets_before]
        return base
    except (pywintypes.com_error, ToolError) as exc:
        location = plan.get("_what_results") or name or "new summary sheet"
        raise PartialChangeError(f"{action} interrupted; {base['sheet']}: {location} may have changed. Use office_undo. {com.translate(exc)}") from None


@office_tool("excel_analysis", "write", title="Goal seek")
def excel_goal_seek(workbook: str, sheet: str, set_cell: str, target_value: float, changing_cell: str,
                    max_iterations: int | None = None, max_change: float | None = None,
                    keep_if_not_converged: bool = False) -> dict:
    """Find a numeric constant making a formula reach target_value. Snapshots only changing_cell for office_undo.

    Args:
        workbook, sheet: exact workbook and sheet.
        set_cell: one formula cell; parsed same-sheet dependency, no active sheet required.
        Incomplete parsing (dynamic references or node limit) allows GoalSeek with dependency_verified=false.
        target_value: finite numeric goal.
        changing_cell: one numeric or empty constant on the same sheet.
        max_iterations, max_change: optional positive limits, restored in finally.
        keep_if_not_converged: false restores and verifies the original value with no undo entry; true keeps the attempt.
    """
    _finite(target_value, "target_value")
    if max_iterations is not None and (type(max_iterations) is not int or not 1 <= max_iterations <= 2147483647):
        raise ToolError("max_iterations must be a positive integer.")
    if max_change is not None:
        _finite(max_change, "max_change")
        if max_change <= 0:
            raise ToolError("max_change must be positive.")
    app, wb = xl.pick_workbook(workbook)
    ws, target = xl.get_range(wb, sheet, set_cell, empty_means_used=False)
    other_ws, changing = xl.get_range(wb, ws.Name, changing_cell, empty_means_used=False)
    for rng in (target, changing):
        xl.validate_rectangle(rng, ws, wb)
        if int(rng.Count) != 1 or rng.MergeCells is not False or bool(rng.HasSpill):
            raise ToolError("Select one unmerged, non-spill cell.")
        xl.reject_pivots(app, ws, rng)
    if other_ws.Name != ws.Name:
        raise ToolError("changing_cell must be on the same sheet.")
    if not bool(target.HasFormula):
        raise ToolError("set_cell must contain a formula.")
    original = changing.Value2
    if bool(changing.HasFormula) or (original is not None and (isinstance(original, bool) or not isinstance(original, (int, float)))):
        raise ToolError("changing_cell must be a numeric or empty constant, not a formula/text/error.")
    if original is not None:
        _finite(original, "changing_cell")
        if original in EXCEL_ERRORS:
            raise ToolError("changing_cell contains an Excel error.")
    if bool(ws.ProtectContents) and bool(changing.Locked):
        raise ToolError("changing_cell is locked on a protected sheet.")
    from .formula_trace import same_sheet_dependency

    depends, complete, dependency_reasons = same_sheet_dependency(app, wb, target, changing)
    if complete and not depends:
        raise ToolError("Формула не зависит от изменяемой ячейки (полный разбор ссылок).")
    iterations, change = app.MaxIterations, app.MaxChange
    manual = int(app.Calculation) == -4135
    undo.require_undo("workbook", app, wb, plan={"_goal_sheet": ws.Name, "_goal_cell": xl.addr_of(changing)})
    try:
        try:
            if max_iterations is not None:
                app.MaxIterations = max_iterations
            if max_change is not None:
                app.MaxChange = max_change
            converged = bool(target.GoalSeek(target_value, changing))
            new, achieved = changing.Value2, target.Value2
            kept = converged or keep_if_not_converged
            if not kept:
                changing.Value2 = original
                if changing.Value2 != original:
                    raise ToolError("Original value could not be verified after rollback.")
                undo.cancel_prepared("workbook", app, wb)
            return {"workbook": wb.Name, "sheet": ws.Name, "converged": converged, "original_value": original,
                    "new_value": new if kept else original, "attempted_value": new, "achieved_value": EXCEL_ERRORS.get(achieved, achieved),
                    "difference": achieved - target_value if isinstance(achieved, (int, float)) and achieved not in EXCEL_ERRORS else None,
                    "kept": bool(kept), "manual_calculation": manual,
                    "dependency_verified": complete and depends, "dependency_reasons": dependency_reasons,
                    "parameters": {"max_iterations": max_iterations if max_iterations is not None else iterations,
                                   "max_change": max_change if max_change is not None else change}}
        finally:
            try:
                app.MaxIterations = iterations
            finally:
                app.MaxChange = change
    except (pywintypes.com_error, ToolError) as exc:
        try:
            applied = int(changing.Value2 != original)
        except (pywintypes.com_error, ToolError):
            applied = "unknown"
        raise PartialChangeError(f"GoalSeek interrupted: applied={applied} cell. {com.translate(exc)}") from None
