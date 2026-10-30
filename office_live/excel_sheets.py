"""Excel worksheets, rows, columns and range operations."""

import math
import re
import time
from collections import Counter
import pywintypes
from . import com, config, undo
from . import xl_common as xl
from .errors import PartialChangeError, ToolError
from .registry import office_tool
from .util import (
    MAX_COLS,
    MAX_ROWS,
    a1_cell,
    col_letter,
    col_number,
    parse_a1,
    parse_color,
    quote_sheet,
    to_grid,
)
from .xl_common import (
    addr_of,
    bounds,
    clip_to_used,
    count_nonempty,
    get_range,
    localize_formula,
    pick_sheet,
    pick_workbook,
    ref_label,
    sub_range,
    validate_sheet_name,
)
from .excel_workbooks import _restore_alerts, _sheet_list


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
    column_widths: dict[str, float] | None = None,
    row_heights: dict[str, float] | None = None,
) -> dict:
    """Set column widths / row heights individually or in maps, or autofit the used range with size/time guards.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        cells: range whose columns/rows are affected ('' = used range for autofit; entire sheet for scalar sizes).
        column_width: width in characters (e.g. 12).
        row_height: height in points (e.g. 18).
        autofit: 'columns', 'rows' or 'both' - fit to content; at most 100000 cells and 1000 dimensions, with a 5 s budget checked between COM calls.
        column_widths: map of individual column letters or 1-based indices to widths, e.g. {'A': 12, 'B': 24}.
        row_heights: map of individual 1-based row indices to heights, e.g. {'1': 24, '3': 18}.

    Scalar sizes are applied first, then maps, then autofit. One office_undo entry restores all dimensions.
    A synchronous Excel COM call cannot be interrupted; the time budget stops subsequent calls and retains undo for partial changes.
    """
    deadline = time.monotonic() + 5.0
    options = dict(cells=cells, column_width=column_width, row_height=row_height,
                   autofit=autofit, column_widths=column_widths, row_heights=row_heights)
    app, wb = pick_workbook(workbook)
    targets, fit_range, mode = _dimension_targets(wb, sheet, options)
    ws = fit_range.Worksheet
    snapshots = [{"cells": ref_label(rng.Worksheet, rng), "property": prop} for rng, prop, _ in targets]
    for axis, prop in (("columns", "ColumnWidth"), ("rows", "RowHeight")):
        if mode in {axis, "both"}:
            snapshots.append({"cells": ref_label(ws, fit_range), "property": prop})
    if mode or column_widths is not None or row_heights is not None:
        undo.require_undo("workbook", app, wb, plan={"sheet": ws.Name, "_dimension_targets": snapshots})
    applied = []
    attempted = False
    try:
        attempted = bool(targets)
        _apply_dimension_targets(targets, applied)
        if mode:
            r1, c1, r2, c2 = bounds(fit_range)
            def check_time():
                if time.monotonic() >= deadline:
                    raise ToolError("Autofit time limit (5 s) exceeded; use a smaller cells range or explicit dimension maps.")
            # Small strips keep each call bounded in work and permit deadline checks.
            axes = ["columns", "rows"] if mode == "both" else [mode]
            for axis in axes:
                for i in (range(c1, c2 + 1) if axis == "columns" else range(r1, r2 + 1)):
                    check_time()
                    strip = sub_range(ws, r1, i, r2, i) if axis == "columns" else sub_range(ws, i, c1, i, c2)
                    attempted = True
                    (strip.Columns if axis == "columns" else strip.Rows).AutoFit()
                    if axis != axes[-1] or i != (c2 if axis == "columns" else r2):
                        check_time()
            applied.append(f"autofit={mode}")
    except (pywintypes.com_error, ToolError) as exc:
        if attempted:
            raise PartialChangeError(f"Dimensions interrupted. {com.translate(exc)}") from None
        undo.cancel_prepared("workbook", app, wb)
        raise
    address = addr_of(fit_range) if mode else cells or ("mapped dimensions" if column_widths or row_heights else "entire sheet")
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cells": address, "applied": applied}


def _apply_dimension_targets(targets, applied):
    for rng, prop, value in targets:
        setattr(rng, prop, value)
        applied.append(f"{addr_of(rng)}: {prop}={value}")


def _dimension_targets(wb, sheet, args):
    """Validate the complete layout before snapshotting or changing dimensions."""
    mode = args.get("autofit", "").lower()
    if mode and mode not in {"columns", "rows", "both"}:
        raise ToolError("autofit must be 'columns', 'rows' or 'both'.")
    _, used = get_range(wb, sheet, args.get("cells", ""))
    ws = used.Worksheet
    # The undo caller reuses these arguments to resolve unqualified snapshots.
    args["sheet"] = ws.Name
    targets = []
    def size(value, prop):
        limit = 255 if prop == "ColumnWidth" else 409.5
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or not 0 <= value <= limit:
            raise ToolError(f"{prop} must be a finite number between 0 and {limit}.")
        return float(value)
    for key, prop in (("column_width", "ColumnWidth"), ("row_height", "RowHeight")):
        if args.get(key) is not None:
            scalar_range = used if args.get("cells") else ws.Cells
            targets.append((scalar_range, prop, size(args[key], prop)))
    for key, prop in (("column_widths", "ColumnWidth"), ("row_heights", "RowHeight")):
        mapping = args.get(key)
        if mapping is None:
            continue
        if not isinstance(mapping, dict):
            raise ToolError(f"{key} must be a map of individual columns/rows to sizes.")
        for label, value in mapping.items():
            label = str(label)
            if prop == "ColumnWidth" and re.fullmatch(r"[A-Za-z]{1,3}", label):
                index = col_number(label)
            elif re.fullmatch(r"[1-9][0-9]*", label):
                index = int(label)
            else:
                raise ToolError(f"Invalid {key} index: {label!r}.")
            if index > (MAX_COLS if prop == "ColumnWidth" else MAX_ROWS):
                raise ToolError(f"Invalid {key} index: {label!r}.")
            # A one-cell range changes its whole column/row without huge range snapshots.
            rng = ws.Range(a1_cell(1, index) if prop == "ColumnWidth" else a1_cell(index, 1))
            targets.append((rng, prop, size(value, prop)))
    if mode:
        rows, cols = int(used.Rows.Count), int(used.Columns.Count)
        count = (cols if mode in {"columns", "both"} else 0) + (rows if mode in {"rows", "both"} else 0)
        if rows * cols > 100000 or count > 1000:
            raise ToolError("Autofit size limit exceeded (100000 cells / 1000 dimensions); use a smaller cells range or explicit dimension maps.")
    if not targets and not mode:
        raise ToolError("Nothing to do: pass column_width, row_height, column_widths, row_heights or autofit.")
    return targets, used, mode


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


FUNCTIONS = {"sum": -4157, "count": -4112, "average": -4106, "max": -4136, "min": -4139,
             "product": -4149, "count_numbers": -4113, "stdev": -4155, "stdevp": -4156, "var": -4164, "varp": -4165}
SUBTOTAL_BARRIER = "Existing subtotals/outline cannot be reconstructed safely; this operation is an undo barrier."
REMOVE_BARRIER = "Removing subtotals has no safe reconstruction strategy; this operation is an undo barrier."


def _column(rng, headers, key):
    matches = [i + 1 for i, h in enumerate(headers) if h == key]
    if not matches:
        matches = [i + 1 for i in range(len(headers)) if col_letter(int(rng.Column) + i) == key]
    if len(matches) != 1:
        raise ToolError(f"Column {key!r} was not found or is ambiguous.")
    return matches[0]


def outline_levels(ws, rng):
    return [int(ws.Rows(i).OutlineLevel) for i in range(int(rng.Row), int(rng.Row) + int(rng.Rows.Count))]


@office_tool("excel_core", "write", title="Subtotals")
def excel_subtotals(workbook: str, sheet: str, cells: str, action: str = "add", group_by: str = "",
                    function: str = "sum", columns: list[str] | None = None, replace_existing: bool = True,
                    page_breaks: bool = False, summary_below: bool = True, sort_first: bool = False) -> dict:
    """Add/remove Excel subtotals with structural undo. Existing outline/subtotals and remove create explicit barriers.

    Args:
        workbook, sheet, cells: exact workbook, sheet and required header/data rectangle within the undo limit.
        action: add or remove.
        group_by, columns: exact headers or absolute column letters inside the block; columns required for add.
        function: sum/count/average/max/min/product/count_numbers/stdev/stdevp/var/varp.
        replace_existing, page_breaks, summary_below: Subtotal options.
        sort_first: sort by group_by, with header, before adding subtotals. Undo restores the original order.
    """
    if action not in {"add", "remove"} or function not in FUNCTIONS:
        raise ToolError("Invalid action or subtotal function.")
    if action == "add" and (not group_by or not isinstance(columns, list) or not columns or any(not isinstance(c, str) for c in columns)):
        raise ToolError("group_by and nonempty columns are required for add.")
    app, wb = xl.pick_workbook(workbook)
    ws, rng = xl.get_range(wb, sheet, cells, empty_means_used=False)
    xl.validate_rectangle(rng, ws, wb)
    size = int(rng.Count)
    if xl.bounds(rng)[2] >= xl.MAX_ROWS:
        raise ToolError("Block must leave a row below it for structural undo verification.")
    if size > config.SETTINGS.undo_max_cells:
        raise ToolError("Block exceeds OFFICE_LIVE_UNDO_MAX_CELLS; narrow the block.")
    if bool(ws.ProtectContents) or rng.MergeCells is not False:
        raise ToolError("Protected sheets and merged cells are not supported.")
    xl.reject_pivots(app, ws, rng)
    for table in ws.ListObjects:
        if app.Intersect(rng, table.Range) is not None:
            raise ToolError("Block intersects an Excel table.")
    if bool(ws.AutoFilterMode):
        af = ws.AutoFilter
        if app.Intersect(rng, af.Range) is not None and any(bool(f.On) for f in af.Filters):
            raise ToolError("Block has active AutoFilter conditions.")
    levels = outline_levels(ws, rng)
    formulas = to_grid(rng.Formula)
    existing = any(level != 1 for level in levels) or any(isinstance(f, str) and "SUBTOTAL(" in f.upper() for row in formulas for f in row)
    groups = unsorted = 0
    if action == "add":
        if int(rng.Rows.Count) < 2:
            raise ToolError("Block must contain a header and data rows.")
        values = to_grid(rng.Value2)
        headers = values[0]
        if any(n > 1 for h, n in Counter(h for h in headers if h not in (None, "")).items()):
            raise ToolError("Duplicate headers are ambiguous.")
        key = _column(rng, headers, group_by)
        totals = [_column(rng, headers, c) for c in columns]
        if len(set(totals)) != len(totals):
            raise ToolError("Duplicate subtotal columns.")
        series = [row[key - 1] for row in values[1:]]
        runs = [v for i, v in enumerate(series) if i == 0 or v != series[i - 1]]
        counts = Counter(runs)
        unsorted = sum(n > 1 for n in counts.values())
        groups = len(counts) if sort_first else len(runs)
        if not config.SETTINGS.undo:
            raise ToolError("Subtotals add requires OFFICE_LIVE_UNDO to be enabled.")
        # The enlarged block must also fit the fingerprint limit, before sorting or inserting any rows.
        if (int(rng.Rows.Count) + groups + 1) * int(rng.Columns.Count) > config.SETTINGS.undo_max_cells:
            raise ToolError("Subtotal result exceeds OFFICE_LIVE_UNDO_MAX_CELLS.")
        if int(rng.Row) + int(rng.Rows.Count) + groups + 1 > xl.MAX_ROWS:
            raise ToolError("Subtotal result would exceed worksheet row boundaries.")
        if xl.bounds(ws.UsedRange)[2] + groups + 1 >= xl.MAX_ROWS:
            raise ToolError("Subtotal insertion could shift used cells past worksheet row boundaries.")
    barrier = REMOVE_BARRIER if action == "remove" else SUBTOTAL_BARRIER if existing else ""
    undo.require_undo("workbook", app, wb, plan={"_subtotal_sheet": ws.Name, "_subtotal_range": xl.addr_of(rng),
                                                "_operation_barrier": barrier}, barrier_reason=barrier)
    old_rows = int(rng.Rows.Count)
    below = ws.Cells(int(rng.Row) + old_rows, int(rng.Column))
    applied = 0
    try:
        if action == "add":
            if sort_first:
                rng.Sort(ws.Cells(int(rng.Row) + 1, int(rng.Column) + key - 1), 1, None, None, 1, None, 1, 1,
                         1, False, 1, 1, 0, 0, 0, None)
                applied = old_rows - 1
            rng.Subtotal(key, FUNCTIONS[function], tuple(totals), bool(replace_existing), bool(page_breaks), 1 if summary_below else 0)
        else:
            rng.RemoveSubtotal()
        current = subtotal_extent(app, wb, ws, rng, below)
        return {"workbook": wb.Name, "sheet": ws.Name, "action": action, "groups": groups,
                "unsorted_groups": 0 if sort_first else unsorted, "inserted_rows": int(current.Rows.Count) - old_rows,
                "range": xl.addr_of(current), "outline_levels": sorted(set(outline_levels(ws, current)))}
    except (pywintypes.com_error, ToolError) as exc:
        try:
            current = subtotal_extent(app, wb, ws, rng, below)
            delta = int(current.Rows.Count) - old_rows
        except (pywintypes.com_error, ToolError):
            delta = "unknown"
            if undo._active is not None:
                entry = undo._active.pending.get(undo.stack_key("workbook", app, wb))
                if entry is not None:
                    entry.barrier("Subtotal block extent is unreadable; structural undo cannot safely remove rows.")
        raise PartialChangeError(f"Subtotal interrupted: applied={applied} sorted rows, row_delta={delta}. {com.translate(exc)}") from None


def subtotal_extent(app, wb, ws, rng, below):
    rows = int(below.Row) - int(rng.Row)
    if rows < 1:
        raise ToolError("Cannot determine subtotal block extent.")
    r, c, _, c2 = xl.bounds(rng)
    current = xl.sub_range(ws, r, c, r + rows - 1, c2)
    if undo._active is not None:
        entry = undo._active.pending.get(undo.stack_key("workbook", app, wb))
        if entry is not None and entry.undoable:
            entry.areas = [undo._area(current)]
            entry.ops[0]["current_address"] = xl.addr_of(current)
    return current
