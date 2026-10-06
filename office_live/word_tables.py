"""Word: таблицы — чтение, создание, запись ячеек, строки/столбцы, объединение, оформление."""


import ctypes
import re

import pywintypes

from . import com, undo
from . import wd_common as wd
from .errors import PartialChangeError, ToolError
from .registry import office_tool
from .util import clean_word_text, cm_to_points, parse_color, to_com_grid, to_word_text
from .wd_common import (
    ALIGNMENTS,
    apply_table_style,
    ensure_free_anchor,
    paragraph_index_at,
    pick_document,
)

_BORDER_IDS = {"top": -1, "left": -2, "bottom": -3, "right": -4, "inside_horizontal": -5, "inside_vertical": -6}
_LINE_STYLES = {"single": 1, "dotted": 2, "dashed": 3, "double": 7, "none": 0}
_VALIGN = {"top": 0, "center": 1, "bottom": 3}
_AUTOFIT = {"fixed": 0, "content": 1, "window": 2}


def get_table(doc, index: int):
    n = int(doc.Tables.Count)
    if n == 0:
        raise ToolError("The document has no tables.")
    if not 1 <= int(index) <= n:
        raise ToolError(f"Table {index} does not exist; the document has {n} table(s).")
    return doc.Tables(int(index))


def _dims(tbl) -> tuple[int, int | None]:
    rows = int(tbl.Rows.Count)
    try:
        return rows, int(tbl.Columns.Count)
    except pywintypes.com_error:
        return rows, None  # объединённые ячейки: столбцы неоднородны


def table_grid(tbl) -> dict:
    """Содержимое таблицы: быстрый путь для обычных таблиц, ячейка-за-ячейкой для неоднородных."""
    rows, cols = _dims(tbl)
    nested = int(tbl.Tables.Count) > 0
    if cols is not None and not nested and bool(tbl.Uniform):
        tokens = tbl.Range.Text.split("\r\x07")
        out, i = [], 0
        for _ in range(rows):
            row = tokens[i:i + cols]
            if len(row) < cols:
                break
            out.append([t.replace("\r", "\n").replace("\x0b", "\n").replace("\x07", "") for t in row])
            i += cols + 1  # +1 — пустой токен маркера конца строки
        if len(out) == rows:
            return {"uniform": True, "rows": rows, "columns": cols, "values": out}
    cells = []
    for cell in tbl.Range.Cells:
        cells.append({"row": int(cell.RowIndex), "col": int(cell.ColumnIndex), "text": clean_word_text(cell.Range.Text)})
    return {"uniform": False, "rows": rows, "columns": cols, "cells": cells}


# ================================================================== чтение


@office_tool("word_tables", "read", title="List tables")
def word_list_tables(document: str = "") -> dict:
    """List the tables of a document: index, size, style, first cell, and the paragraph number where each table starts.

    Args:
        document: exact name or '' for the active document.
    """
    app, doc = pick_document(document)
    out = []
    for i in range(1, int(doc.Tables.Count) + 1):
        t = doc.Tables(i)
        rows, cols = _dims(t)
        try:
            style = t.Style.NameLocal
        except (pywintypes.com_error, AttributeError):
            style = None
        out.append({
            "index": i, "rows": rows, "columns": cols, "uniform": bool(t.Uniform), "style": style,
            "first_cell": clean_word_text(t.Cell(1, 1).Range.Text)[:60], "start_paragraph": paragraph_index_at(doc, int(t.Range.Start)),
            "nested": int(t.Tables.Count),
        })
    return {"document": doc.Name, "tables": out}


@office_tool("word_tables", "read", title="Read table")
def word_read_table(document: str = "", table_index: int = 1) -> dict:
    """Read a table as a grid of cell texts (rows x columns). Tables with merged cells come back as a list of {row, col, text} cells instead.

    Args:
        document: exact name or '' for the active document.
        table_index: 1-based table number (see word_list_tables).
    """
    app, doc = pick_document(document)
    tbl = get_table(doc, table_index)
    return {"document": doc.Name, "table": int(table_index), **table_grid(tbl)}


# ================================================================== создание и запись


def _cell_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v == int(v) and abs(v) < 1e15:
        return str(int(v))
    return str(v)


def build_table(doc, data, position="end", paragraph=0, style="Table Grid", header_row=True, header_fill="#D9E2F3",
                autofit="window", column_widths_cm=None, font_size=None, alignment=None, right_align_cols=None, table_index=0) -> dict:
    """Создаёт таблицу из двумерного массива (общий код word_create_table и мостов Excel->Word)."""
    grid = to_com_grid(data, pad=None)
    rows, cols = len(grid), len(grid[0])
    if rows * cols > 5000:
        raise ToolError("Table is too large for one call (5000 cells max); create it in parts with word_write_table.")
    # --- все параметры проверяем ДО вставки: ошибка не должна оставлять в документе полуготовую таблицу
    if autofit not in _AUTOFIT:
        raise ToolError("autofit must be 'window', 'content' or 'fixed'.")
    if alignment and alignment.lower() not in ALIGNMENTS:
        raise ToolError("alignment must be left, center, right or justify.")
    if header_row and header_fill:
        parse_color(header_fill)
    if column_widths_cm and len(column_widths_cm) > cols:
        raise ToolError(f"column_widths_cm has {len(column_widths_cm)} widths but the table has {cols} columns.")
    for j in right_align_cols or []:
        if not 1 <= int(j) <= cols:
            raise ToolError(f"right-aligned column {j} is outside 1..{cols}.")
    anchor = ensure_free_anchor(doc, position, paragraph, separate_from_table=True, table_index=table_index)
    # текст -> строки с табами -> ConvertToTable: на порядок быстрее, чем заполнять ячейки по одной
    # переводы строк внутри ячейки ConvertToTable принимает за границу строки таблицы — прячем их за маркером и возвращаем после
    mark = "\ue000"
    multiline = []
    lines = []
    for i, row in enumerate(grid):
        cells = []
        for j, v in enumerate(row):
            t = _cell_text(v).replace("\t", " ").replace("\r\n", "\n").replace("\r", "\n")
            if "\n" in t:
                multiline.append((i + 1, j + 1, t))
                t = t.replace("\n", mark)
            cells.append(t)
        lines.append("\t".join(cells))
    anchor.Text = "\r".join(lines)
    try:
        tbl = anchor.ConvertToTable(1, rows, cols)  # wdSeparateByTabs
    except Exception:
        try:
            anchor.Text = ""  # убрать вставленный текст, если таблица не получилась
        except Exception:  # noqa: BLE001
            pass
        raise
    t_start = int(tbl.Range.Start)
    try:
        for r, c, t in multiline:
            tbl.Cell(r, c).Range.Text = t.replace("\n", "\r")  # абзацы внутри ячейки
        used_style = apply_table_style(doc, tbl, style) if style else None
        tbl.AutoFitBehavior(_AUTOFIT[autofit])
        if column_widths_cm:
            for j, w in enumerate(column_widths_cm, start=1):
                tbl.Columns(j).Width = cm_to_points(w)
        if font_size is not None:
            tbl.Range.Font.Size = float(font_size)
        if alignment:
            tbl.Range.ParagraphFormat.Alignment = ALIGNMENTS[alignment.lower()]
        for j in right_align_cols or []:
            for r in range(2 if header_row else 1, rows + 1):
                tbl.Cell(r, j).Range.ParagraphFormat.Alignment = 2  # числа — по правому краю
        if header_row:
            hr = tbl.Rows(1)
            hr.HeadingFormat = True
            hr.Range.Font.Bold = True
            if header_fill:
                hr.Shading.BackgroundPatternColor = parse_color(header_fill)
    except Exception as exc:  # noqa: BLE001 — таблица уже в документе: говорим, что именно осталось сделанным
        index = next((i for i in range(1, int(doc.Tables.Count) + 1) if int(doc.Tables(i).Range.Start) == t_start), int(doc.Tables.Count))
        raise ToolError(
            f"The table was inserted (table {index}, {rows}x{cols}) but a later formatting step failed: {com.translate(exc)}. "
            "Fix it with word_format_table / word_modify_table or delete the table."
        ) from None
    index = next((i for i in range(1, int(doc.Tables.Count) + 1) if int(doc.Tables(i).Range.Start) == t_start), int(doc.Tables.Count))
    return {"ok": True, "document": doc.Name, "table": index, "rows": rows, "columns": cols, "style_used": used_style}


@office_tool("word_tables", "write", title="Create table")
def word_create_table(
    document: str,
    data: list[list[str | int | float | bool | None]],
    position: str = "end",
    paragraph: int = 0,
    style: str = "Table Grid",
    header_row: bool = True,
    header_fill: str = "#D9E2F3",
    autofit: str = "window",
    column_widths_cm: list[float] | None = None,
    font_size: float | None = None,
    alignment: str | None = None,
    table_index: int = 0,
) -> dict:
    """Create a table from a 2-D array of values (the first row is the header by default) and return its index.

    Args:
        document: exact document name (required).
        data: rows of cells; shorter rows are padded; numbers/booleans are converted to text.
        position: 'end' (default), 'start', 'after_paragraph' / 'before_paragraph' (with `paragraph`) or 'after_table' / 'before_table' (with `table_index` - no need to count paragraphs).
        paragraph: paragraph number for the *_paragraph positions.
        style: table style name; 'Table Grid' (plain borders) works in every UI language.
        header_row: format row 1 as a header (bold, shaded, repeats on each page).
        header_fill: header background color ('' = none).
        autofit: 'window' (full width, default), 'content' or 'fixed'.
        column_widths_cm: optional width of each column in centimetres.
        font_size: font size in the whole table.
        alignment: horizontal alignment of the table text: left|center|right|justify.
    """
    app, doc = pick_document(document)
    return build_table(doc, data, position, paragraph, style, header_row, header_fill, autofit, column_widths_cm, font_size, alignment, None, table_index)


@office_tool("word_tables", "write", title="Write table cells")
def word_write_table(
    document: str,
    table_index: int,
    values: list[list[str | int | float | bool | None]],
    start_row: int = 1,
    start_col: int = 1,
    add_rows: bool = True,
) -> dict:
    """Write a block of values into an existing table starting at (start_row, start_col). Rows are appended when the data runs past the last row (add_rows=true). Not supported for tables with merged cells in the target area.

    Args:
        document: exact document name (required).
        table_index: 1-based table number.
        values: 2-D array of cell values.
        start_row, start_col: 1-based top-left target cell.
        add_rows: append rows when needed.
    """
    app, doc = pick_document(document)
    tbl = get_table(doc, table_index)
    grid = to_com_grid(values, pad=None)
    rows, cols = _dims(tbl)
    if cols is None:
        raise ToolError("The table has merged cells with uneven columns; write single cells via word_modify_table/cell edits or use a simpler table.")
    if int(start_row) < 1 or int(start_col) < 1:
        raise ToolError("start_row and start_col are 1-based.")
    if int(start_col) - 1 + len(grid[0]) > cols:
        raise ToolError(f"Data is {len(grid[0])} columns wide from column {start_col}, but the table has {cols} columns.")
    need_rows = int(start_row) - 1 + len(grid)
    if need_rows > rows:
        if not add_rows:
            raise ToolError(f"Data needs {need_rows} rows but the table has {rows}; pass add_rows=true.")
        for _ in range(need_rows - rows):
            tbl.Rows.Add()
    written = 0
    for i, row in enumerate(grid):
        for j, v in enumerate(row):
            if v is None:
                continue
            tbl.Cell(int(start_row) + i, int(start_col) + j).Range.Text = to_word_text(_cell_text(v))
            written += 1
    return {"ok": True, "document": doc.Name, "table": int(table_index), "cells_written": written, "rows_now": int(tbl.Rows.Count)}


# ================================================================== структура


def _sort_value(text, kind, language):
    """Разбор Windows с тем же LCID, который явно передаём Word.Sort."""
    text = text.strip().replace("\u00a0", " ").replace("\u202f", " ")
    if not text or "\n" in text or (kind == "date" and not re.search(r"\b\d{4}\b", text)):
        return None
    library = ctypes.WinDLL("oleaut32")
    parse = library.VarDateFromStr if kind == "date" else library.VarR8FromStr
    parse.argtypes = (ctypes.c_wchar_p, ctypes.c_ulong, ctypes.c_ulong, ctypes.POINTER(ctypes.c_double))
    parse.restype = ctypes.c_long
    result = ctypes.c_double()
    if parse(text, language, 2 if kind == "date" else 0, ctypes.byref(result)) < 0:
        return None
    return result.value


def _sort_rows(tbl, rows, cols):
    return [[clean_word_text(tbl.Cell(r, c).Range.Text) for c in range(1, cols + 1)] for r in range(1, rows + 1)]


@office_tool("word_tables", "write", title="Sort table")
def word_sort_table(document: str, table_index: int, keys: list[dict], header: bool = True, case_sensitive: bool = False) -> dict:
    """Sort a plain Word table by 1..3 columns, preserving its header by default; supports office_undo and refuses later user edits on undo. Returns the first 20 rows.

    Args:
        document: exact open document name.
        table_index: 1-based table index.
        keys: 1..3 unique keys {column: 1-based integer, order: asc/desc (default asc), type: text/number/date (default text)}.
        header: true keeps the first row in place and requires at least one data row.
        case_sensitive: case-sensitive text sorting, false by default.

    Refuses merged/nested tables, tracking or document protection, and unparseable numeric/date keys before any edit.
    Windows parses typed keys using the table's language (Windows locale for mixed language); dates require a four-digit year. This same language is passed to Word.Sort.
    """
    if not isinstance(keys, list) or not 1 <= len(keys) <= 3:
        raise ToolError("keys must contain 1..3 sort keys.")
    plan, columns = [], set()
    for key in keys:
        if not isinstance(key, dict):
            raise ToolError("Each key must be an object with column, order and type.")
        column, order, kind = key.get("column"), key.get("order", "asc"), key.get("type", "text")
        if isinstance(column, bool) or not isinstance(column, int) or column < 1:
            raise ToolError("column must be a positive 1-based integer.")
        if column in columns:
            raise ToolError("Duplicate sort column.")
        if order not in {"asc", "desc"} or kind not in {"text", "number", "date"}:
            raise ToolError("order must be asc/desc; type must be text/number/date.")
        columns.add(column)
        plan.append((column, kind, order))
    if isinstance(table_index, bool) or not isinstance(table_index, int) or table_index < 1:
        raise ToolError("table_index must be a positive 1-based integer.")
    app, doc = pick_document(document)
    tbl = get_table(doc, table_index)
    if bool(doc.TrackRevisions) and int(doc.ProtectionType) != 0:
        raise ToolError("Turn off Track Changes first: sorting would create many tracked insertions/deletions.")
    if int(doc.ProtectionType) not in {-1, 0}:
        raise ToolError("The document is protected; no rows were changed.")
    if not bool(tbl.Uniform):
        raise ToolError("Merged or nonuniform tables cannot be sorted.")
    try:
        rows, cols = int(tbl.Rows.Count), int(tbl.Columns.Count)
    except pywintypes.com_error:
        raise ToolError("Merged or nonuniform tables cannot be sorted.") from None
    if int(tbl.Range.Cells.Count) != rows * cols:
        raise ToolError("Merged cells cannot be sorted.")
    if int(tbl.Tables.Count) or int(tbl.NestingLevel) != 1:
        raise ToolError("Nested tables cannot be sorted.")
    if rows < (2 if header else 1):
        raise ToolError("The table has no data rows to sort with this header setting.")
    if max(columns) > cols:
        raise ToolError(f"Sort column is outside the table (1..{cols}).")
    language = int(tbl.Range.LanguageID)
    if language in {0, 1024, 9999999}:
        language = int(ctypes.WinDLL("kernel32").GetUserDefaultLCID())
    before = _sort_rows(tbl, rows, cols)
    invalid = []
    for col, kind, _ in plan:
        if kind == "text":
            continue
        for r in range(1 if header else 0, rows):
            text = before[r][col - 1]
            if _sort_value(text, kind, language) is None:
                invalid.append({"row": r + 1, "column": col, "type": kind, "text": text[:200]})
    if invalid:
        raise ToolError(f"Unparseable sort cells ({len(invalid)}): {invalid}; no rows were changed.")
    from .undo import require_undo

    require_undo("document", app, doc)
    args = [bool(header)]
    for col, kind, order in plan:
        args.extend((str(col), {"text": 0, "number": 1, "date": 2}[kind], 0 if order == "asc" else 1))
    for _ in range(3 - len(plan)):
        args.extend(("", 0, 0))
    args.extend((bool(case_sensitive), False, False, False, False, False, language))
    try:
        tbl.Sort(*args)
        after = _sort_rows(tbl, rows, cols)
    except (pywintypes.com_error, ToolError) as exc:
        try:
            current = _sort_rows(tbl, rows, cols)
            applied = sum(a != b for a, b in zip(before, current, strict=True))
            progress = f"applied={applied} row positions changed, unchanged={rows - applied}"
        except (pywintypes.com_error, ToolError):
            progress = "applied=unknown (Word failed to read back row positions)"
        raise PartialChangeError(f"Sort interrupted: {progress}. {com.translate(exc)}") from None
    return {"document": doc.Name, "table": table_index, "rows": rows, "columns": cols, "uniform": True,
            "sorted_rows": rows - int(header), "values": after[:20], "truncated": rows > 20, "language_id": language}


@office_tool("word_tables", "write", title="Modify table structure", destructive=True)
def word_modify_table(
    document: str,
    table_index: int,
    action: str,
    position: int = 0,
    count: int = 1,
    row: int = 0,
    col: int = 0,
    to_row: int = 0,
    to_col: int = 0,
    autofit: str = "",
    widths_cm: list[float] | None = None,
) -> dict:
    """Change a table's structure.

    Args:
        document: exact document name (required).
        table_index: 1-based table number.
        action: 'add_rows' (insert `count` rows before row `position`, 0 = append at the end) | 'add_columns' (same for columns) | 'delete_rows' (delete `count` rows starting at `position`) | 'delete_columns' | 'merge_cells' (from cell row,col to cell to_row,to_col) | 'split_cell' (row,col into `count` columns... uses to_row as number of rows) | 'autofit' (autofit: window|content|fixed) | 'set_widths' (widths_cm per column) | 'header_row' (repeat row 1 on each page) | 'delete_table' | 'to_text' (convert the table to tab-separated text).
        position, count: for the add/delete actions.
        row, col, to_row, to_col: cell coordinates for merge_cells/split_cell.
        autofit, widths_cm: for 'autofit' / 'set_widths'.
    """
    app, doc = pick_document(document)
    tbl = get_table(doc, table_index)
    act = action.lower()
    n_rows, n_cols = _dims(tbl)
    if act in ("add_rows", "add_columns", "delete_rows", "delete_columns") and not 1 <= int(count) <= 1000:
        raise ToolError("count must be between 1 and 1000.")
    if act == "add_rows" and position and not 1 <= int(position) <= n_rows:
        raise ToolError(f"position {position} is outside the table (1..{n_rows} rows).")
    if act == "add_columns" and position and (n_cols is None or not 1 <= int(position) <= n_cols):
        raise ToolError(f"position {position} is outside the table ({n_cols} columns; tables with merged cells are not supported here).")
    if act == "delete_columns":
        if n_cols is None:
            raise ToolError("The table has merged cells with uneven columns; deleting columns is not supported for it.")
        if position and int(position) + int(count) - 1 > n_cols:
            raise ToolError(f"The table has only {n_cols} columns.")
        if position and int(position) == 1 and int(count) >= n_cols:
            raise ToolError("That would delete every column; use action='delete_table' instead.")
    if act == "delete_rows" and position and int(position) == 1 and int(count) >= n_rows:
        raise ToolError("That would delete every row; use action='delete_table' instead.")
    if act in ("merge_cells", "split_cell"):
        for name, v, limit in (("row", row, n_rows), ("to_row", to_row if act == "merge_cells" else 0, n_rows)):
            if v and not 1 <= int(v) <= limit:
                raise ToolError(f"{name}={v} is outside the table ({n_rows} rows).")
        if n_cols is not None:
            for name, v in (("col", col), ("to_col", to_col if act == "merge_cells" else 0)):
                if v and not 1 <= int(v) <= n_cols:
                    raise ToolError(f"{name}={v} is outside the table ({n_cols} columns).")
    if act == "add_rows":
        for _ in range(int(count)):
            if position:
                tbl.Rows.Add(tbl.Rows(int(position)))
            else:
                tbl.Rows.Add()
    elif act == "add_columns":
        for _ in range(int(count)):
            if position:
                tbl.Columns.Add(tbl.Columns(int(position)))
            else:
                tbl.Columns.Add()
    elif act == "delete_rows":
        if not position:
            raise ToolError("'position' (first row to delete) is required.")
        total = int(tbl.Rows.Count)
        if int(position) + int(count) - 1 > total:
            raise ToolError(f"The table has only {total} rows.")
        for r in range(int(position) + int(count) - 1, int(position) - 1, -1):
            tbl.Rows(r).Delete()
    elif act == "delete_columns":
        if not position:
            raise ToolError("'position' (first column to delete) is required.")
        for c in range(int(position) + int(count) - 1, int(position) - 1, -1):
            tbl.Columns(c).Delete()
    elif act == "merge_cells":
        if not (row and col and to_row and to_col):
            raise ToolError("Pass row, col, to_row and to_col.")
        tbl.Cell(int(row), int(col)).Merge(tbl.Cell(int(to_row), int(to_col)))
    elif act == "split_cell":
        if not (row and col):
            raise ToolError("Pass row and col.")
        tbl.Cell(int(row), int(col)).Split(int(to_row) or 1, int(count))  # Split(NumRows, NumColumns)
    elif act == "autofit":
        if autofit not in _AUTOFIT:
            raise ToolError("autofit must be 'window', 'content' or 'fixed'.")
        tbl.AutoFitBehavior(_AUTOFIT[autofit])
    elif act == "set_widths":
        if not widths_cm:
            raise ToolError("'widths_cm' is required.")
        tbl.AllowAutoFit = False
        for j, w in enumerate(widths_cm, start=1):
            tbl.Columns(j).Width = cm_to_points(w)
    elif act == "header_row":
        tbl.Rows(1).HeadingFormat = True
    elif act == "delete_table":
        tbl.Delete()
        return {"ok": True, "document": doc.Name, "deleted_table": int(table_index), "tables_left": int(doc.Tables.Count)}
    elif act == "to_text":
        tbl.ConvertToText(1)  # wdSeparateByTabs
        return {"ok": True, "document": doc.Name, "converted_table": int(table_index), "tables_left": int(doc.Tables.Count)}
    else:
        raise ToolError("action must be one of: add_rows, add_columns, delete_rows, delete_columns, merge_cells, split_cell, autofit, set_widths, header_row, delete_table, to_text")
    rows, cols = _dims(tbl)
    return {"ok": True, "document": doc.Name, "table": int(table_index), "action": act, "rows": rows, "columns": cols}


# ================================================================== оформление


@office_tool("word_tables", "write", title="Format table")
def word_format_table(
    document: str,
    table_index: int,
    row_from: int = 0,
    row_to: int = 0,
    col_from: int = 0,
    col_to: int = 0,
    style: str | None = None,
    fill_color: str | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    font_color: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    alignment: str | None = None,
    vertical_alignment: str | None = None,
    borders: str | None = None,
    border_style: str = "single",
    border_color: str = "#000000",
    border_width_pt: float = 0.5,
    cell_padding_cm: float | None = None,
) -> dict:
    """Format a table or a block of its cells. Only the properties you pass are changed. Cell block = row_from..row_to x col_from..col_to (all zero = the whole table).

    Args:
        document: exact document name (required).
        table_index: 1-based table number.
        row_from, row_to, col_from, col_to: the block of cells (0 = from the start / to the end).
        style: table style name (e.g. 'Table Grid', or a UI-language style name from word_list_styles(kind='table')).
        fill_color: cell background '#RRGGBB' or name ('none' clears it).
        font_name, font_size, font_color, bold, italic: text formatting in the block.
        alignment: left|center|right|justify. vertical_alignment: top|center|bottom.
        borders: all|outline|inside|top|bottom|left|right|none (applied to the whole table, or to the block's outer edges for 'outline'-style choices when a block is given).
        border_style: single|dotted|dashed|double. border_color, border_width_pt: border look.
        cell_padding_cm: inner margin of all cells.
    """
    app, doc = pick_document(document)
    tbl = get_table(doc, table_index)
    rows, cols = _dims(tbl)
    cols = cols or 1
    r1, r2 = int(row_from) or 1, int(row_to) or rows
    c1, c2 = int(col_from) or 1, int(col_to) or cols
    if not (1 <= r1 <= r2 <= rows) or not (1 <= c1 <= c2 <= cols):
        raise ToolError(f"Cell block out of range; the table is {rows}x{cols}.")
    # --- все параметры проверяем и переводим ДО первого изменения
    if alignment and alignment.lower() not in ALIGNMENTS:
        raise ToolError("alignment must be left, center, right or justify.")
    if vertical_alignment and vertical_alignment.lower() not in _VALIGN:
        raise ToolError("vertical_alignment must be top, center or bottom.")
    if font_size is not None and not 1 <= float(font_size) <= 1638:
        raise ToolError("font_size must be between 1 and 1638 points.")
    font_rgb = parse_color(font_color) if font_color else None
    fill_value = None if fill_color is None else (-16777216 if fill_color.lower() == "none" else parse_color(fill_color))  # wdColorAutomatic
    border_choice = line_style = line_width = border_rgb = None
    if borders is not None:
        border_choice = borders.lower()
        if border_choice not in _BORDER_CHOICES:
            raise ToolError(f"borders must be one of {sorted(_BORDER_CHOICES)}")
        if border_style.lower() not in _LINE_STYLES or _LINE_STYLES[border_style.lower()] == 0:
            raise ToolError("border_style must be single, dotted, dashed or double.")
        line_style, line_width, border_rgb = _LINE_STYLES[border_style.lower()], _line_width(border_width_pt), parse_color(border_color)
    wants_text = any(v is not None for v in (font_name, font_size, font_rgb, bold, italic)) or bool(alignment)
    if not (style or wants_text or fill_value is not None or vertical_alignment or cell_padding_cm is not None or borders is not None):
        raise ToolError("Nothing to apply: pass at least one formatting parameter.")

    whole = (r1, c1, r2, c2) == (1, 1, rows, cols)
    applied = []
    if style:
        apply_table_style(doc, tbl, style)
        applied.append(f"style={style}")
    if wants_text:
        # Непрерывный Range от первой до последней ячейки блока захватил бы и соседние ячейки строк — форматируем ячейка за ячейкой
        targets = [tbl.Range] if whole else [tbl.Cell(r, c).Range for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
        for rg in targets:
            font = rg.Font
            if font_name:
                font.Name = font_name
            if font_size is not None:
                font.Size = float(font_size)
            if font_rgb is not None:
                font.Color = font_rgb
            if bold is not None:
                font.Bold = bool(bold)
            if italic is not None:
                font.Italic = bool(italic)
            if alignment:
                rg.ParagraphFormat.Alignment = ALIGNMENTS[alignment.lower()]
        applied += [n for n, v in (("font_name", font_name), ("font_size", font_size), ("font_color", font_rgb), ("bold", bold), ("italic", italic), ("alignment", alignment)) if v not in (None, "")]
    if fill_value is not None or vertical_alignment:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = tbl.Cell(r, c)
                if fill_value is not None:
                    cell.Shading.BackgroundPatternColor = fill_value
                if vertical_alignment:
                    cell.VerticalAlignment = _VALIGN[vertical_alignment.lower()]
        if fill_value is not None:
            applied.append("fill_color")
        if vertical_alignment:
            applied.append("vertical_alignment")
    if cell_padding_cm is not None:
        pad = cm_to_points(cell_padding_cm)
        tbl.TopPadding = tbl.BottomPadding = tbl.LeftPadding = tbl.RightPadding = pad
        applied.append("cell_padding_cm")
    if borders is not None:
        _apply_borders(tbl, border_choice, (r1, c1, r2, c2), whole, line_style, line_width, border_rgb)
        applied.append(f"borders={border_choice}")
    return {"ok": True, "document": doc.Name, "table": int(table_index), "block": [r1, c1, r2, c2], "applied": applied}


_BORDER_CHOICES = {"all", "outline", "inside", "top", "bottom", "left", "right", "none"}


def _block_edges(choice: str, r: int, c: int, block: tuple) -> list[str]:
    """Какие границы ячейки (r, c) относятся к выбору `choice` для блока (r1, c1, r2, c2)."""
    r1, c1, r2, c2 = block
    if choice in ("all", "none"):
        return ["top", "left", "bottom", "right"]
    edges = []
    if choice == "outline":
        for name, hit in (("top", r == r1), ("left", c == c1), ("bottom", r == r2), ("right", c == c2)):
            if hit:
                edges.append(name)
    elif choice == "inside":
        for name, hit in (("top", r > r1), ("left", c > c1), ("bottom", r < r2), ("right", c < c2)):
            if hit:
                edges.append(name)
    elif choice == "top" and r == r1:
        edges.append("top")
    elif choice == "bottom" and r == r2:
        edges.append("bottom")
    elif choice == "left" and c == c1:
        edges.append("left")
    elif choice == "right" and c == c2:
        edges.append("right")
    return edges


def _apply_borders(tbl, choice: str, block: tuple, whole: bool, line_style: int | None, line_width: int | None, rgb: int | None):
    """Границы только там, где просили: для всего стола — через tbl.Borders (без Enable=True, который включил бы и внутренние),
    для блока — по границам его ячеек, не затрагивая остальную таблицу."""
    r1, c1, r2, c2 = block
    if whole:
        if choice == "none":
            tbl.Borders.Enable = False
            return
        names = {
            "all": list(_BORDER_IDS), "outline": ["top", "left", "bottom", "right"], "inside": ["inside_horizontal", "inside_vertical"],
            "top": ["top"], "bottom": ["bottom"], "left": ["left"], "right": ["right"],
        }[choice]
        for name in names:
            br = tbl.Borders(_BORDER_IDS[name])
            br.LineStyle = line_style
            br.LineWidth = line_width
            br.Color = rgb
        return
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = tbl.Cell(r, c)
            for name in _block_edges(choice, r, c, block):
                br = cell.Borders(_BORDER_IDS[name])
                if choice == "none":
                    br.LineStyle = 0
                else:
                    br.LineStyle = line_style
                    br.LineWidth = line_width
                    br.Color = rgb


def _line_width(points: float) -> int:
    """pt -> WdLineWidth (константы фиксированного набора)."""
    table = [(0.25, 2), (0.5, 4), (0.75, 6), (1.0, 8), (1.5, 12), (2.25, 18), (3.0, 24), (4.5, 36), (6.0, 48)]
    return min(table, key=lambda t: abs(t[0] - points))[1]


def _separator(separator, action):
    if separator == "paragraph":
        if action != "table_to_text":
            raise ToolError("paragraph separator is only supported for table_to_text.")
        return 0
    value = {"tab": 1, "comma": 2, "semicolon": ";"}.get(separator, separator)
    if isinstance(value, str) and (len(value) != 1 or value in "\r\n\x07"):
        raise ToolError("separator must be tab/comma/semicolon or one non-paragraph character.")
    return value


@office_tool("word_tables", "write", title="Convert table/text")
def word_table_text(document: str, action: str, table: int = 1, start_paragraph: int = 0, end_paragraph: int = 0,
                    separator: str = "tab", columns: int = 0, header_row: bool = True, style: str = "") -> dict:
    """Convert a plain table to text or a paragraph range to a table, with one Word UndoRecord.

    Args:
        document: exact document name.
        action: table_to_text/text_to_table.
        table: 1-based table index for table_to_text (nested/merged tables refused).
        start_paragraph, end_paragraph: inclusive nonempty paragraph range outside tables for text_to_table.
        separator: tab/comma/semicolon/one character; paragraph only for table_to_text.
        columns: 0 infers the maximum fields per line; otherwise positive.
        header_row: mark first row as repeating heading. style: optional existing table style.
    """
    if action not in {"table_to_text", "text_to_table"}:
        raise ToolError("Invalid conversion action.")
    sep = _separator(separator, action)
    if type(columns) is not int or not 0 <= columns <= 63:
        raise ToolError("columns must be an integer between 0 and 63.")
    app, doc = wd.pick_document(document)
    if action == "table_to_text":
        if type(table) is not int or table < 1:
            raise ToolError("table must be a positive integer.")
        tbl = get_table(doc, table)
        if not bool(tbl.Uniform) or int(tbl.Tables.Count) or int(tbl.NestingLevel) != 1:
            raise ToolError("Nested or merged tables cannot be converted.")
        try:
            rows, cols = int(tbl.Rows.Count), int(tbl.Columns.Count)
        except pywintypes.com_error:
            raise ToolError("Merged tables cannot be converted.") from None
        if int(tbl.Range.Cells.Count) != rows * cols:
            raise ToolError("Merged tables cannot be converted.")
        before_tables = int(doc.Tables.Count)
        undo.require_undo("document", app, doc)
        try:
            result = tbl.ConvertToText(sep, False)
            return {"document": doc.Name, "start_paragraph": wd.paragraph_index_at(doc, int(result.Start)),
                    "end_paragraph": wd.paragraph_index_at(doc, max(int(result.Start), int(result.End) - 1)),
                    "start": int(result.Start), "end": int(result.End)}
        except pywintypes.com_error as exc:
            try:
                applied = rows if int(doc.Tables.Count) < before_tables else "unknown"
            except pywintypes.com_error:
                applied = "unknown"
            raise PartialChangeError(f"ConvertToText interrupted: applied={applied} converted rows. {com.translate(exc)}") from None
    if any(type(n) is not int for n in (start_paragraph, end_paragraph)) or not 1 <= start_paragraph <= end_paragraph <= int(doc.Paragraphs.Count):
        raise ToolError("Paragraph range is outside the document or reversed.")
    rng = wd.paragraphs_range(doc, start_paragraph, end_paragraph)
    if int(rng.Tables.Count) or wd._in_table(doc, int(rng.Start)) or wd._in_table(doc, max(int(rng.Start), int(rng.End) - 1)):
        raise ToolError("Paragraph range intersects a table.")
    text = rng.Text.removesuffix("\r")  # remove the final paragraph marker, preserving selected empty paragraphs
    if not text.strip():
        raise ToolError("Paragraph range is empty.")
    literal = {1: "\t", 2: ","}.get(sep, sep)
    fields = [line.count(literal) + 1 for line in text.split("\r")]
    num_columns = columns or max(fields)
    if num_columns > 63:
        raise ToolError("The inferred table exceeds Word's 63-column limit.")
    style_obj = None
    if style:
        for candidate in wd.TABLE_STYLE_ALIASES.get(style.strip().lower(), [style]):
            try:
                item = doc.Styles(candidate)
                if int(item.Type) == 3:
                    style_obj = item
                    break
            except pywintypes.com_error:
                continue
        if style_obj is None:
            raise ToolError("Table style was not found; no changes made.")
    undo.require_undo("document", app, doc)
    anchor = int(rng.Start)
    try:
        # live probe: NumRows=None reaches Word as 0 ("must be from 1 to 32767"); pass the row count explicitly
        tbl = rng.ConvertToTable(sep, len(fields), num_columns)
        tbl.Rows(1).HeadingFormat = bool(header_row)
        if style_obj is not None:
            tbl.Style = style_obj
        index = next(i for i in range(1, int(doc.Tables.Count) + 1) if int(doc.Tables(i).Range.Start) == int(tbl.Range.Start))
        return {"document": doc.Name, "table": index, "rows": int(tbl.Rows.Count), "columns": int(tbl.Columns.Count),
                "uneven_rows": sum(n != max(fields) for n in fields) if columns == 0 else 0}
    except (pywintypes.com_error, StopIteration) as exc:
        try:
            applied = next((int(doc.Tables(i).Rows.Count) for i in range(1, int(doc.Tables.Count) + 1)
                            if int(doc.Tables(i).Range.Start) == anchor), 0)
        except pywintypes.com_error:
            applied = "unknown"
        raise PartialChangeError(f"ConvertToTable interrupted: applied={applied} converted rows. {com.translate(exc)}") from None
