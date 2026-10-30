"""Excel: книги, листы, чтение/запись, формулы, структура, сортировка, фильтры, имена, выделение."""

import re
import unicodedata

import pywintypes

from . import com, excel_import
from .errors import PartialChangeError, ToolError
from .registry import office_tool
from .util import (
    DATE_FORMATS,
    MAX_COLS,
    MAX_ROWS,
    a1_cell,
    a1_range,
    cell_kind,
    excel_date_serial,
    parse_a1,
    smart_number,
    to_com_grid,
    to_grid,
)
from .xl_common import (
    addr_of,
    bounded_range,
    bounds,
    clip_to_used,
    count_nonempty,
    error_cells,
    formula_cells,
    get_range,
    pick_sheet,
    pick_workbook,
    preview,
    read_grid,
    set_number_format,
    sheet_is_empty,
    sub_range,
    workbook_allowed,
)

Cell = str | int | float | bool | None
GRID_HINT = "2-D array: a list of rows, each row a list of cell values (a flat list is treated as one row)."

# коды форматов SaveAs
from .excel_workbooks import (
    _SAVE_FORMATS as _SAVE_FORMATS,
    _LEGACY_FORMATS as _LEGACY_FORMATS,
    _LEGACY_EXTENSIONS as _LEGACY_EXTENSIONS,
    _COMPATIBILITY_NOTE as _COMPATIBILITY_NOTE,
    _save_workbook as _save_workbook,
    _autosave as _autosave,
    _pid as _pid,
    _is_visible as _is_visible,
    _used_address as _used_address,
    _sheet_list as _sheet_list,
    _restore_alerts as _restore_alerts,
    excel_list_workbooks as excel_list_workbooks,
    excel_workbook_info as excel_workbook_info,
    excel_open_workbook as excel_open_workbook,
    _excel_running as _excel_running,
    excel_new_workbook as excel_new_workbook,
    excel_close_workbook as excel_close_workbook,
    excel_save as excel_save,
    excel_save_as as excel_save_as,
    excel_export_pdf as excel_export_pdf,
)


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
    per_match_links: bool = False,
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
        per_match_links: true returns a link per match (up to 20); default returns one workbook link for the result.
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
    date_style = (date_number_format or date_format.lower()) if dates else None
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
                    set_number_format(app, change["cell"], date_style if dates else "General")
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
                    set_number_format(app, change["cell"], "General")
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
    readback: str = "full",
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
        readback: 'full' (default) keeps the existing value preview; 'summary' returns target shape, counts and up to five sample cells;
            'none' omits readback and readback_types. All modes report errors (up to 20). Summary/none scan the entire written target.
    """
    if value_mode not in ("auto", "text"):
        raise ToolError("value_mode must be 'auto' or 'text'.")
    if readback not in ("full", "summary", "none"):
        raise ToolError("readback must be 'full', 'summary' or 'none'.")
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
    try:
        _write_values(target, grid, value_mode, fill=fill)
    except pywintypes.com_error as exc:
        raise ToolError("Excel refused the write (protected sheet, merged cells in the way, or invalid data): " + com.com_error_text(exc)) from None
    r1, c1, r2, c2 = bounds(target)
    out = {
        "ok": True, "workbook": wb.Name, "sheet": ws.Name, "cells": addr_of(target),
        "written_shape": [rows, cols], "overwritten_nonempty_cells": max(before, 0),
    }
    if readback != "full":
        types, samples, errors = {}, [], []
        target_rows, target_cols = r2 - r1 + 1, c2 - c1 + 1
        step = max(1, 20000 // target_cols)
        for start in range(r1, r2 + 1, step):
            chunk = read_grid(sub_range(ws, start, c1, min(r2, start + step - 1), c2))
            for i, row in enumerate(chunk):
                for j, value in enumerate(row):
                    kind = cell_kind(value)
                    if kind != "empty":
                        types[kind] = types.get(kind, 0) + 1
                    if readback == "summary" and len(samples) < 5:
                        samples.append({"cell": a1_cell(start + i, c1 + j), "value": value})
            if len(errors) < 20:
                errors.extend(error_cells(chunk, start, c1, limit=20 - len(errors)))
        out["errors"] = errors
        if readback == "summary":
            total = target_rows * target_cols
            nonempty = sum(types.values())
            out.update(readback={"shape": [target_rows, target_cols], "total_cells": total,
                                 "nonempty_cells": nonempty, "empty_cells": total - nonempty, "samples": samples},
                       readback_types=types)
        return out
    back = preview(target)
    types: dict[str, int] = {}
    for row in back:
        for v in row:
            k = cell_kind(v)
            if k != "empty":
                types[k] = types.get(k, 0) + 1
    out.update(readback=back, readback_types=types, errors=error_cells(back, r1, c1))
    return out


def _write_values(target, grid, value_mode, *, fill=False, text_columns=()):
    """Общий путь ввода write_range/import: Excel сам разбирает числа и даты."""
    grid = tuple(tuple(("'" + v) if isinstance(v, str) and v != "" and (value_mode == "text" or j in text_columns)
                       else v for j, v in enumerate(row)) for row in grid)
    target.Value = grid[0][0] if fill else grid  # Value, not Value2: unchanged write_range semantics (dates)


# Code lives in excel_import; registering here keeps the catalog order.
excel_import_table = office_tool("excel_core", "write", title="Import table")(excel_import.excel_import_table)


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


from .excel_sheets import (
    _position_args as _position_args,
    excel_add_worksheet as excel_add_worksheet,
    excel_manage_sheet as excel_manage_sheet,
    excel_delete_sheet as excel_delete_sheet,
    _rows_cols_range as _rows_cols_range,
    excel_insert_rows_columns as excel_insert_rows_columns,
    excel_delete_rows_columns as excel_delete_rows_columns,
    excel_set_dimensions as excel_set_dimensions,
    _apply_dimension_targets as _apply_dimension_targets,
    _dimension_targets as _dimension_targets,
    _runs as _runs,
    _line_spec as _line_spec,
    _where_match as _where_match,
    excel_hide_rows_columns as excel_hide_rows_columns,
    excel_group_rows_columns as excel_group_rows_columns,
    excel_copy_range as excel_copy_range,
    _resolve_column as _resolve_column,
    excel_sort_range as excel_sort_range,
    _FILTER_OPS as _FILTER_OPS,
    excel_filter_range as excel_filter_range,
    excel_remove_duplicates as excel_remove_duplicates,
    excel_merge_cells as excel_merge_cells,
    excel_manage_names as excel_manage_names,
    excel_calculate as excel_calculate,
    FUNCTIONS as FUNCTIONS,
    SUBTOTAL_BARRIER as SUBTOTAL_BARRIER,
    REMOVE_BARRIER as REMOVE_BARRIER,
    _column as _column,
    outline_levels as outline_levels,
    excel_subtotals as excel_subtotals,
    subtotal_extent as subtotal_extent,
)
