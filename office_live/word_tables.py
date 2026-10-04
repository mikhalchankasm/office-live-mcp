"""Word: таблицы — чтение, создание, запись ячеек, строки/столбцы, объединение, оформление."""

import pywintypes

from .errors import ToolError
from .registry import office_tool
from .util import clean_word_text, cm_to_points, parse_color, to_com_grid, to_word_text
from .wd_common import (
    ALIGNMENTS, apply_table_style, ensure_free_anchor, paragraph_index_at, pick_document,
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
    if autofit not in _AUTOFIT:
        raise ToolError("autofit must be 'window', 'content' or 'fixed'.")
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
    tbl = anchor.ConvertToTable(1, rows, cols)  # wdSeparateByTabs
    for r, c, t in multiline:
        tbl.Cell(r, c).Range.Text = t.replace("\n", "\r")  # абзацы внутри ячейки
    used_style = apply_table_style(doc, tbl, style) if style else None
    tbl.AutoFitBehavior(_AUTOFIT[autofit])
    if column_widths_cm:
        for j, w in enumerate(column_widths_cm, start=1):
            if j <= cols:
                tbl.Columns(j).Width = cm_to_points(w)
    if font_size is not None:
        tbl.Range.Font.Size = float(font_size)
    if alignment:
        if alignment.lower() not in ALIGNMENTS:
            raise ToolError("alignment must be left, center, right or justify.")
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
    t_start = int(tbl.Range.Start)
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
    need_rows = int(start_row) - 1 + len(grid)
    if need_rows > rows:
        if not add_rows:
            raise ToolError(f"Data needs {need_rows} rows but the table has {rows}; pass add_rows=true.")
        for _ in range(need_rows - rows):
            tbl.Rows.Add()
    if int(start_col) - 1 + len(grid[0]) > cols:
        raise ToolError(f"Data is {len(grid[0])} columns wide from column {start_col}, but the table has {cols} columns.")
    written = 0
    for i, row in enumerate(grid):
        for j, v in enumerate(row):
            if v is None:
                continue
            tbl.Cell(int(start_row) + i, int(start_col) + j).Range.Text = to_word_text(_cell_text(v))
            written += 1
    return {"ok": True, "document": doc.Name, "table": int(table_index), "cells_written": written, "rows_now": int(tbl.Rows.Count)}


# ================================================================== структура


@office_tool("word_tables", "write", title="Modify table structure")
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
    if act == "add_rows":
        for k in range(int(count)):
            if position:
                tbl.Rows.Add(tbl.Rows(int(position)))
            else:
                tbl.Rows.Add()
    elif act == "add_columns":
        for k in range(int(count)):
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
    applied = []
    if style:
        apply_table_style(doc, tbl, style)
        applied.append(f"style={style}")
    block = doc.Range(int(tbl.Cell(r1, c1).Range.Start), int(tbl.Cell(r2, c2).Range.End))
    font = block.Font
    if font_name:
        font.Name = font_name
        applied.append("font_name")
    if font_size is not None:
        font.Size = float(font_size)
        applied.append("font_size")
    if font_color:
        font.Color = parse_color(font_color)
        applied.append("font_color")
    if bold is not None:
        font.Bold = bool(bold)
        applied.append("bold")
    if italic is not None:
        font.Italic = bool(italic)
        applied.append("italic")
    if alignment:
        if alignment.lower() not in ALIGNMENTS:
            raise ToolError("alignment must be left, center, right or justify.")
        block.ParagraphFormat.Alignment = ALIGNMENTS[alignment.lower()]
        applied.append("alignment")
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = tbl.Cell(r, c)
            if fill_color is not None:
                cell.Shading.BackgroundPatternColor = -16777216 if fill_color.lower() == "none" else parse_color(fill_color)  # wdColorAutomatic
            if vertical_alignment:
                if vertical_alignment.lower() not in _VALIGN:
                    raise ToolError("vertical_alignment must be top, center or bottom.")
                cell.VerticalAlignment = _VALIGN[vertical_alignment.lower()]
    if fill_color is not None:
        applied.append("fill_color")
    if vertical_alignment:
        applied.append("vertical_alignment")
    if cell_padding_cm is not None:
        pad = cm_to_points(cell_padding_cm)
        tbl.TopPadding = tbl.BottomPadding = tbl.LeftPadding = tbl.RightPadding = pad
        applied.append("cell_padding_cm")
    if borders is not None:
        b = borders.lower()
        if border_style.lower() not in _LINE_STYLES or _LINE_STYLES[border_style.lower()] == 0:
            raise ToolError("border_style must be single, dotted, dashed or double.")
        edges = {
            "all": list(_BORDER_IDS), "outline": ["top", "left", "bottom", "right"], "inside": ["inside_horizontal", "inside_vertical"],
            "top": ["top"], "bottom": ["bottom"], "left": ["left"], "right": ["right"], "none": list(_BORDER_IDS),
        }
        if b not in edges:
            raise ToolError(f"borders must be one of {sorted(edges)}")
        if b == "none":
            tbl.Borders.Enable = False
        else:
            tbl.Borders.Enable = True
            width = _line_width(border_width_pt)
            for name in edges[b]:
                br = tbl.Borders(_BORDER_IDS[name])
                br.LineStyle = _LINE_STYLES[border_style.lower()]
                br.LineWidth = width
                br.Color = parse_color(border_color)
        applied.append(f"borders={b}")
    if not applied:
        raise ToolError("Nothing to apply: pass at least one formatting parameter.")
    return {"ok": True, "document": doc.Name, "table": int(table_index), "block": [r1, c1, r2, c2], "applied": applied}


def _line_width(points: float) -> int:
    """pt -> WdLineWidth (константы фиксированного набора)."""
    table = [(0.25, 2), (0.5, 4), (0.75, 6), (1.0, 8), (1.5, 12), (2.25, 18), (3.0, 24), (4.5, 36), (6.0, 48)]
    return min(table, key=lambda t: abs(t[0] - points))[1]

