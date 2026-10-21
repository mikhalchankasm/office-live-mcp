"""Excel: импорт CSV/TSV/TXT/JSON одним блоком (регистрирует excel_core в прежнем порядке каталога)."""

import csv
import io
import json
import math
import re

from . import com, config, undo
from . import xl_common as xl
from .errors import PartialChangeError, ToolError
from .safety import check_path
from .util import MAX_COLS, MAX_ROWS, a1_cell, a1_range, col_letter, parse_a1, to_grid
from .xl_common import count_nonempty, pick_workbook, sheet_names, sub_range, suspend_events, validate_sheet_name


def _import_rows(path, delimiter, encoding, header, max_rows):
    full = check_path(path, "read", {"csv", "tsv", "txt", "json"})
    try:
        with open(full, "rb") as source:
            raw = source.read()
        if encoding == "auto":
            encoding = "utf-8-sig" if raw.startswith(b"\xef\xbb\xbf") else "utf-16" if raw.startswith((b"\xff\xfe", b"\xfe\xff")) else "utf-8"
            try:
                raw.decode(encoding)
            except UnicodeDecodeError:
                if encoding != "utf-8":
                    raise
                encoding = "cp1251"
        text = raw.decode(encoding)
        keys = []
        if full.lower().endswith(".json"):
            rows = json.loads(text)
            if not isinstance(rows, list) or not rows:
                raise ToolError("JSON must be a nonempty array of objects or arrays.")
            objects = isinstance(rows[0], dict)
            for i, row in enumerate(rows):
                if not isinstance(row, dict if objects else list):
                    raise ToolError(f"Mixed JSON row types at $[{i}].")
                for key, value in (row.items() if objects else enumerate(row)):
                    if isinstance(value, (dict, list)):
                        raise ToolError(f"Nested JSON value at $[{i}][{json.dumps(key, ensure_ascii=False)}].")
                    if objects and key not in keys:
                        keys.append(key)
                if (len(keys) if objects else len(row)) > 200 or i + 1 + int(objects and header) > max_rows:
                    raise ToolError("Import exceeds max_rows / 200 columns; split the file (Power Query is outside this tool).")
            if objects:
                rows = ([keys] if header else []) + [[row.get(key) for key in keys] for row in rows]
            delimiter = None
        else:
            if delimiter == "auto":
                try:
                    delimiter = csv.Sniffer().sniff(text[:65536], delimiters=",;\t|").delimiter
                except csv.Error:
                    if any(char in text[:65536] for char in ",;\t|"):
                        raise ToolError("Cannot detect delimiter; pass an explicit delimiter.") from None
                    delimiter = "\t" if full.lower().endswith(".tsv") else ","
            rows = []
            for row in csv.reader(io.StringIO(text, newline=""), delimiter=delimiter, strict=True):
                rows.append(row)
                if len(rows) > max_rows or len(row) > 200:
                    raise ToolError("Import exceeds max_rows / 200 columns; split the file (Power Query is outside this tool).")
    except (OSError, UnicodeError, LookupError, ValueError, csv.Error) as exc:
        raise ToolError(f"Cannot read import file: {exc}") from None
    width = max(map(len, rows), default=0)
    if not width:
        raise ToolError("Import file has no cells.")
    warnings = []
    if any(len(row) != width for row in rows):
        warnings.append("Uneven rows padded with blanks; no data truncated.")
    rows = [row + [None] * (width - len(row)) for row in rows]
    if header:
        keys = rows[0]
        if any(v is None or str(v).strip() == "" for v in keys):
            warnings.append("Empty headers.")
        if len(set(keys)) != len(keys):
            warnings.append("Duplicate headers; select text columns by 1-based number.")
    for row in rows:
        for value in row:
            if isinstance(value, str) and len(value) > 32767:
                raise ToolError("A cell value exceeds 32767 characters.")
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                try:
                    valid = math.isfinite(value)
                except OverflowError:
                    valid = False
                if not valid:
                    raise ToolError("JSON numbers must be finite Excel numbers.")
    return rows, keys, delimiter, encoding, warnings


_NUMBER_LIKE = re.compile(r"[+-]\s*\d[\d\s.,]*(?:[eE][+-]?\d+)?%?")


def _formula_like(value) -> bool:
    """A string Excel's typed-input parser would turn into a formula (or a broken one)."""
    if not isinstance(value, str) or not value:
        return False
    if value[0] in "=@":
        return True
    return value[0] in "+-" and not _NUMBER_LIKE.fullmatch(value.strip())


def _formula_like_cells(rows):
    """{row_index: {col_index, ...}} of cells that must stay literal text."""
    found = {}
    for i, row in enumerate(rows):
        cols = {j for j, value in enumerate(row) if _formula_like(value)}
        if cols:
            found[i] = cols
    return found


def _neutralize(rows, literal, text_columns):
    """Prefix formula-like strings with Excel's literal-text apostrophe; text columns get it in _write_values."""
    if not literal:
        return rows
    return [[("'" + v) if j in literal.get(i, ()) and j not in text_columns else v for j, v in enumerate(row)]
            for i, row in enumerate(rows)]


def _compact_import_result(result):
    """Лимит включает добавленные обёрткой undo и ссылки; данные на листе не обрезаются."""
    from .excel_core import _runs

    if "sample" not in result:
        return
    if len(json.dumps(result).encode("utf-8")) > 2048 and len(result["text_columns"]) > 32 and all(isinstance(i, int) for i in result["text_columns"]):
        result["text_columns"] = [a if a == b else f"{a}-{b}" for a, b in _runs(result["text_columns"])]
    while len(json.dumps(result).encode("utf-8")) > 2048:
        if result.get("links"):
            result.pop("links")
        else:
            longest = max(range(len(result["sample"])), key=lambda i: len(json.dumps(result["sample"][i])))
            row = result["sample"][longest]
            if not row:
                break
            column = max(range(len(row)), key=lambda j: len(json.dumps(row[j])))
            if isinstance(row[column], str) and len(row[column]) > 8:
                row[column] = row[column][:len(row[column]) // 2] + "…"
            else:
                row.pop()
            if "Sample truncated to fit 2 KB." not in result["warnings"]:
                result["warnings"].append("Sample truncated to fit 2 KB.")


def excel_import_table(
    path: str, workbook: str = "", sheet: str = "Data", start_cell: str = "A1",
    delimiter: str = "auto", encoding: str = "auto", value_mode: str = "auto",
    text_columns: list[str | int] | None = None, header: bool = True, layout: bool = True,
    overwrite: bool = False, max_rows: int = 200000, formulas: bool = False,
) -> dict:
    """Import CSV/TSV/TXT or flat JSON in one Value2 block, with optional layout and one undo entry.

    File text is untrusted: with formulas=false (default) strings starting with = or @, and + or - strings that are
    not numbers (phone numbers like +7 (999) ...), are stored as literal text, so a file cannot plant formulas
    (WEBSERVICE, HYPERLINK, DDE). Pass formulas=true only for a file the user trusts to contain formulas.

    Empty workbook creates a new unsaved book (undo barrier); sheet defaults to Data and is created if absent.
    start_cell is one A1 cell. delimiter='auto' sniffs , ; tab | in the first 64 KiB; encoding='auto' uses BOM,
    strict UTF-8, then cp1251. value_mode matches excel_write_range; text_columns are exact unique headers or
    1-based numbers, formatted as text within the imported block. JSON objects use the ordered union of keys;
    header=false omits those generated headers. Nested values are refused. max_rows includes the header,
    is at most 200000; at most 200 columns. Larger files must be split (no Power Query).
    layout applies bold/wrapped headers and content widths capped at 60 characters. Freeze only on an active
    sheet; otherwise the reply asks to apply excel_sheet_view after activation. Existing books require a safe
    undo snapshot within OFFICE_LIVE_UNDO_MAX_CELLS. No occupied cells are replaced without overwrite=true.
    Reply includes shape, detected encoding/delimiter, 1-based text columns (compact a-b runs if needed)
    and up to three source rows (2 KB).
    """
    from .excel_core import _apply_dimension_targets, _dimension_targets, _write_values
    from .excel_format import _freeze_panes, _prepare_format

    validate_sheet_name(sheet)
    if value_mode not in {"auto", "text"} or (delimiter != "auto" and (len(delimiter) != 1 or delimiter in "\r\n\0")):
        raise ToolError("Use value_mode auto/text and delimiter auto or one non-newline character.")
    if isinstance(max_rows, bool) or not isinstance(max_rows, int) or not 1 <= max_rows <= 200000:
        raise ToolError("max_rows must be 1..200000 (including headers).")
    if not re.fullmatch(r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]*", start_cell):
        raise ToolError("start_cell must be one A1 cell.")
    r, c, _, _ = parse_a1(start_cell)
    rows, keys, delimiter, encoding, warnings = _import_rows(path, delimiter, encoding, header, max_rows)
    height, width = len(rows), len(rows[0])
    if r + height - 1 > MAX_ROWS or c + width - 1 > MAX_COLS:
        raise ToolError("Import exceeds sheet boundaries.")
    selected = set(range(width)) if value_mode == "text" else set()
    for item in text_columns or []:
        if isinstance(item, str):
            if keys.count(item) != 1:
                raise ToolError(f"Unknown or duplicate text column header: {item!r}.")
            index = keys.index(item)
        elif isinstance(item, int) and not isinstance(item, bool) and 1 <= item <= width:
            index = item - 1
        else:
            raise ToolError("text_columns must contain unique headers or 1-based column numbers.")
        selected.add(index)
    cells = a1_range(r, c, r + height - 1, c + width - 1)
    ws = None
    if workbook:
        app, wb = pick_workbook(workbook)
        if wb.ReadOnly:
            raise ToolError("Workbook is read-only.")
        if sheet.casefold() in {name.casefold() for name in sheet_names(wb, any_type=True)}:
            ws = xl.pick_sheet(wb, sheet, exact_only=True)
            target = ws.Range(cells)
            if ws.ProtectContents or target.MergeCells is not False or target.HasArray is not False or target.HasSpill is not False:
                raise ToolError("Import refuses protected sheets, merges, array formulas and spills.")
            xl.reject_pivots(app, ws, target)
            if any(app.Intersect(target, table.Range) is not None for table in ws.ListObjects):
                raise ToolError("Import refuses existing Excel tables; use a plain range.")
            if not overwrite and count_nonempty(app, target):
                occupied = [a1_cell(r + i, c + j) for i, row in enumerate(to_grid(target.Formula))
                            for j, v in enumerate(row) if v is not None][:20]
                raise ToolError(f"Occupied cells in {cells}: {', '.join(occupied)}; pass overwrite=true.")
        elif wb.ProtectStructure:
            raise ToolError("Workbook structure is protected; cannot create sheet.")
        if config.SETTINGS.undo and height * width > config.SETTINGS.undo_max_cells:
            raise ToolError("Import exceeds OFFICE_LIVE_UNDO_MAX_CELLS; split the file or import into a new workbook.")
        undo.require_undo("workbook", app, wb, plan={"sheet": ws.Name if ws else sheet, "cells": cells, "_import_created": ws is None})
    else:
        app = com.primary_app("excel", launch=True)
        suspend_events(app)
    completed = []
    try:
        if not workbook:
            wb = app.Workbooks.Add(-4167)  # xlWBATWorksheet: one plain worksheet
            completed.append("workbook created")
            undo.selected("workbook", app, wb)
            undo.require_undo("workbook", app, wb, barrier_reason=undo.IMPORT_BOOK_BARRIER)
            ws = wb.Worksheets(1)
            ws.Name = sheet
        elif ws is None:
            ws = wb.Worksheets.Add(None, wb.Sheets(wb.Sheets.Count))
            completed.append("sheet created")
            ws.Name = sheet
        target = ws.Range(cells)
        for index in sorted(selected):
            sub_range(ws, r, c + index, r + height - 1, c + index).NumberFormat = "@"
        completed.append("text formats applied")
        literal = _formula_like_cells(rows) if value_mode == "auto" and not formulas else {}
        _write_values(target, _neutralize(rows, literal, selected), value_mode, text_columns=selected)
        if literal:
            warnings.append(f"{sum(map(len, literal.values()))} cells starting with = @ + - stored as text "
                            "(formulas=false protects against formulas in untrusted files).")
        completed.append("values written")
        if layout:
            if header:
                _prepare_format(app, wb, ws, sub_range(ws, r, c, r, c + width - 1), bold=True, wrap_text=True)()
            widths = {col_letter(c + j): min(60, max(8, max(len(str(row[j] or "")) for row in rows) + 2)) for j in range(width)}
            targets, _, _ = _dimension_targets(wb, ws.Name, {"cells": cells, "column_widths": widths})
            _apply_dimension_targets(targets, [])
            completed.append("layout applied")
            if header:
                if app.ActiveWorkbook.FullName == wb.FullName and app.ActiveSheet.Name == ws.Name:
                    _freeze_panes(app.ActiveWindow, r + 1, 1)
                else:
                    warnings.append(f"Freeze skipped: activate {ws.Name}, then call excel_sheet_view(freeze_at='A{r + 1}').")
        result = {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cells": cells, "rows": height, "columns": width,
                  "delimiter": delimiter, "encoding": encoding, "text_columns": [i + 1 for i in sorted(selected)],
                  "sample": [row[:] for row in rows[:3]], "warnings": warnings}
        _compact_import_result(result)
        return result
    except Exception as exc:
        raise PartialChangeError(f"Import interrupted at {cells}; completed: {', '.join(completed) or 'none confirmed'}; "
                                 f"the current step may be partial. {exc}") from None
