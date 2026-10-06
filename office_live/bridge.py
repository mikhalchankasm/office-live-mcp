"""Мосты между приложениями: таблица Word <-> диапазон Excel, диаграммы и картинки из Excel в Word, слияние Excel -> Word."""

import contextlib
import os
import re
import shutil
import tempfile

import pywintypes

from . import com
from .errors import ToolError
from .excel_core import _prepare_target
from .excel_format import render_range_png
from .registry import office_tool
from .safety import WORD_EXTS, check_path
from .util import a1_cell, clean_word_text, smart_number, to_com_grid, to_grid
from .wd_common import all_documents, pick_document
from .win32 import Win32
from .word_core import _word_running, fill_placeholders
from .word_layout import insert_picture
from .word_tables import build_table, get_table, table_grid
from .xl_common import (
    _guard_workbook, addr_of, bounds, count_nonempty, error_cells, get_range, pick_sheet, pick_workbook, preview, read_grid, sheet_names,
    suspend_events, validate_sheet_name,
)


def _target_sheet(wb, sheet: str, create: bool):
    if sheet and sheet.lower() not in [n.lower() for n in sheet_names(wb, any_type=True)]:
        if not create:
            raise ToolError(f"Sheet '{sheet}' does not exist in {wb.Name}; pass create_sheet=true to add it.")
        validate_sheet_name(sheet)
        ws = wb.Worksheets.Add(None, wb.Sheets(wb.Sheets.Count))
        ws.Name = sheet
        return ws
    return pick_sheet(wb, sheet)


def _excel_for_write(workbook: str, create_workbook: bool):
    try:
        return pick_workbook(workbook)
    except ToolError:
        if not create_workbook:
            raise
        app = com.primary_app("excel", launch=True)
        suspend_events(app)
        return _guard_workbook(app, app.Workbooks.Add())


# ================================================================== Word -> Excel


@office_tool("bridge", "write", title="Word table -> Excel")
def bridge_word_table_to_excel(
    document: str,
    table_index: int,
    workbook: str,
    sheet: str = "",
    top_left: str = "A1",
    convert: str = "numbers",
    decimal_separator: str = "auto",
    header_bold: bool = True,
    create_sheet: bool = False,
    create_workbook: bool = False,
    overwrite: bool = False,
) -> dict:
    """Copy a table from a Word document into an Excel sheet. Numbers typed as text in Word ('1 234,50', '12%') become real Excel numbers, everything else stays text (IDs like 007 keep their zeros). Works across Word/Excel instances.

    Args:
        document: Word document holding the table (name or '' for active).
        table_index: 1-based table number (see word_list_tables).
        workbook: target Excel workbook ('' = active).
        sheet: target sheet ('' = active; unknown name + create_sheet=true adds it).
        top_left: cell where the table's top-left corner goes.
        convert: 'numbers' (default: numeric-looking text -> numbers), 'text' (everything stays text), 'excel' (let Excel interpret strings as if typed - dates, formulas).
        decimal_separator: ',' or '.' used to read numbers in Word text; 'auto' = Excel's regional setting.
        header_bold: bold the first row.
        create_sheet, create_workbook: create the target sheet / a new workbook when missing.
        overwrite: allow replacing non-empty cells.
    """
    _, doc = pick_document(document)
    tbl = get_table(doc, table_index)
    g = table_grid(tbl)
    if g["uniform"]:
        text_grid = g["values"]
    else:
        rows = g["rows"]
        cols = max(c["col"] for c in g["cells"])
        text_grid = [[None] * cols for _ in range(rows)]
        for c in g["cells"]:
            text_grid[c["row"] - 1][c["col"] - 1] = c["text"]
    app, wb = _excel_for_write(workbook, create_workbook)
    ws = _target_sheet(wb, sheet, create_sheet)
    dec = decimal_separator if decimal_separator in (",", ".") else app.International[2]
    mode = convert.lower()
    if mode not in ("numbers", "text", "excel"):
        raise ToolError("convert must be 'numbers', 'text' or 'excel'.")
    numbers = texts = 0
    pct_cols = {}
    out = []
    for row in text_grid:
        out_row = []
        for j, cell in enumerate(row):
            s = "" if cell is None else str(cell)
            if s.strip() == "":
                out_row.append(None)
            elif mode == "numbers" and (num := smart_number(s, dec)) is not None:
                out_row.append(num)
                numbers += 1
                if s.strip().endswith("%"):
                    pct_cols.setdefault(j, True)
                else:
                    pct_cols[j] = False
            elif mode == "excel":
                out_row.append(s)
            else:
                out_row.append("'" + s)  # литерал: Excel не станет «угадывать» число/дату/формулу
                texts += 1
        out.append(out_row)
    grid = to_com_grid(out)
    rows, cols = len(grid), len(grid[0])
    anchor = ws.Range(top_left)
    target, _ = _prepare_target(ws, anchor, rows, cols)
    existing = count_nonempty(app, target)
    if existing > 0 and not overwrite:
        raise ToolError(f"Target {addr_of(target)} already contains {existing} non-empty cells; pass overwrite=true or choose another top_left.")
    target.Value = grid
    r1, c1, _, _ = bounds(target)
    for j, is_pct in pct_cols.items():
        if is_pct:
            ws.Range(f"{a1_cell(r1, c1 + j)}:{a1_cell(r1 + rows - 1, c1 + j)}").NumberFormat = "0%"
    if header_bold and rows:
        ws.Range(f"{a1_cell(r1, c1)}:{a1_cell(r1, c1 + cols - 1)}").Font.Bold = True
    back = preview(target)
    return {
        "ok": True, "workbook": wb.Name, "sheet": ws.Name, "cells": addr_of(target), "shape": [rows, cols],
        "numbers_converted": numbers, "text_cells": texts, "source_uniform": g["uniform"],
        "readback": back, "errors": error_cells(back, r1, c1),
    }


@office_tool("bridge", "write", title="Word text -> Excel")
def bridge_word_text_to_excel(
    document: str,
    workbook: str,
    sheet: str = "",
    top_left: str = "A1",
    start_paragraph: int = 1,
    end_paragraph: int = 0,
    headings_only: bool = False,
    create_sheet: bool = False,
    create_workbook: bool = False,
    overwrite: bool = False,
) -> dict:
    """List the paragraphs of a Word document in an Excel sheet as rows [paragraph no, style, text] - handy for analysing, tagging or restructuring a document's text.

    Args:
        document: Word document (name or '' for active).
        workbook / sheet / top_left: target in Excel (see bridge_word_table_to_excel).
        start_paragraph, end_paragraph: paragraph range (end 0 = to the end; at most 5000 paragraphs).
        headings_only: only paragraphs with an outline level (headings).
        create_sheet, create_workbook, overwrite: as in bridge_word_table_to_excel.
    """
    _, doc = pick_document(document)
    n = int(doc.Paragraphs.Count)
    first = max(1, int(start_paragraph))
    last = min(n, int(end_paragraph) or n, first + 4999)
    rows = [["Paragraph", "Style", "Text"]]
    for i in range(first, last + 1):
        p = doc.Paragraphs(i)
        if headings_only and int(p.OutlineLevel) >= 10:
            continue
        txt = clean_word_text(p.Range.Text)
        if not txt.strip():
            continue
        rows.append([i, p.Style.NameLocal, "'" + txt])
    app, wb = _excel_for_write(workbook, create_workbook)
    ws = _target_sheet(wb, sheet, create_sheet)
    grid = to_com_grid(rows)
    target, _ = _prepare_target(ws, ws.Range(top_left), len(grid), 3)
    existing = count_nonempty(app, target)
    if existing > 0 and not overwrite:
        raise ToolError(f"Target {addr_of(target)} already contains {existing} non-empty cells; pass overwrite=true or choose another top_left.")
    target.Value = grid
    r1, c1, _, _ = bounds(target)
    ws.Range(f"{a1_cell(r1, c1)}:{a1_cell(r1, c1 + 2)}").Font.Bold = True
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cells": addr_of(target), "paragraphs_written": len(rows) - 1}


# ================================================================== Excel -> Word


def _hidden_lines(ws, r1, c1, r2, c2):
    rows = {r for r in range(r1, r2 + 1) if bool(ws.Cells(r, c1).EntireRow.Hidden)}
    cols = {c for c in range(c1, c2 + 1) if bool(ws.Cells(r1, c).EntireColumn.Hidden)}
    return rows, cols


def _display_look(df, *, row=False):
    """(заливка, цвет шрифта, жирный, курсив) по DisplayFormat. Для строки — None, если вид неоднороден (COM Null);
    у ячейки Null бывает только от смешанного шрифта в тексте: такой цвет не переносим, а жирный/курсив считаем выключенными."""
    interior, font = df.Interior, df.Font
    look = (interior.ColorIndex, interior.Color, font.Color, font.Bold, font.Italic)
    if row and any(v is None for v in look):
        return None
    index, fill, color, bold, italic = look
    fill = None if index is None or int(index) == -4142 else int(fill)
    return fill, -1 if color is None else int(color), bool(bold), bool(italic)


def _has_color(look) -> bool:
    fill, color = look[0], look[1]
    return color not in (0, 16777215 if fill is None else -1) and color >= 0


def _plain(look) -> bool:
    return look[0] is None and not _has_color(look) and not look[2] and not look[3]


def _apply_look(target, look):
    """target — строка или ячейка таблицы Word (у обеих есть Shading и Range)."""
    fill, color, bold, italic = look
    has_color = _has_color(look)
    if fill is not None and fill != 16777215:
        target.Shading.BackgroundPatternColor = fill
    if has_color or bold or italic:
        rng = target.Range
        if has_color:
            rng.Font.Color = color
        if bold:
            rng.Font.Bold = True
        if italic:
            rng.Font.Italic = True


def _copy_cell_looks(app, ws, tbl, r1, c1, keep_r, keep_c, header_row) -> int:
    """Переносит вид ячеек Excel (как на экране, с условным форматированием) в таблицу Word: заливка, цвет шрифта, жирный, курсив.

    Каждое COM-обращение дорого (44x5 ячеек по одной — ~40 с в демо 2026-10-05): строку с одинаковым видом читаем одним
    DisplayFormat и оформляем одной строкой Word; по ячейкам — только неоднородные строки.
    """
    total = len(keep_r) * len(keep_c)
    rows = keep_r if total <= 1500 else keep_r[:1]
    contiguous = keep_c == list(range(keep_c[0], keep_c[-1] + 1))
    n = 0
    for ii, i in enumerate(rows, start=1):
        look = None
        if contiguous and len(keep_c) > 1:
            look = _display_look(ws.Range(f"{a1_cell(r1 + i, c1 + keep_c[0])}:{a1_cell(r1 + i, c1 + keep_c[-1])}").DisplayFormat, row=True)
        if look is not None and _plain(look):
            continue
        if look is not None:
            try:
                _apply_look(tbl.Rows(ii), look)
                n += len(keep_c)
                continue
            except pywintypes.com_error:
                pass  # строка Word с объединёнными ячейками: оформляем по ячейкам
        for jj, j in enumerate(keep_c, start=1):
            cell_look = look or _display_look(ws.Cells(r1 + i, c1 + j).DisplayFormat)
            if not _plain(cell_look):
                _apply_look(tbl.Cell(ii, jj), cell_look)
                n += 1
    return n


@contextlib.contextmanager
def _frozen_screen(app):
    """Без перерисовки Word построение и оформление таблицы заметно быстрее; прежнее значение возвращается всегда."""
    try:
        updating = app.ScreenUpdating
        app.ScreenUpdating = False
    except pywintypes.com_error:
        yield
        return
    try:
        yield
    finally:
        app.ScreenUpdating = updating


@office_tool("bridge", "write", title="Excel range -> Word table")
def bridge_excel_range_to_word_table(
    workbook: str,
    cells: str,
    document: str,
    sheet: str = "",
    position: str = "end",
    paragraph: int = 0,
    style: str = "Table Grid",
    header_row: bool = True,
    header_fill: str = "#D9E2F3",
    display_text: bool = True,
    skip_hidden: bool = True,
    autofit: str = "window",
    font_size: float | None = None,
    max_cells: int = 2500,
    keep_formatting: bool = True,
    table_index: int = 0,
) -> dict:
    """Copy an Excel range into a Word document as a real, editable Word table. By default the text is exactly what Excel shows (number formats, dates, percentages), hidden rows/columns are skipped, and numeric columns are right-aligned.

    Args:
        workbook: Excel workbook ('' = active).
        cells: range to copy, e.g. 'A1:F20' or a defined name.
        document: target Word document (required).
        sheet: sheet of the range ('' = active).
        position: where the table goes - 'after_paragraph' / 'before_paragraph' (with `paragraph`) or 'after_table' / 'before_table' (with `table_index` - no need to count paragraphs).
        paragraph: paragraph number for the *_paragraph positions.
        style: table style ('Table Grid' by default). header_row: first row is a header (bold, shaded).
        header_fill: header color. display_text: use displayed text (true) or raw values (false).
        skip_hidden: leave out hidden rows and columns. autofit: window|content|fixed. font_size: points.
        max_cells: safety cap.
        keep_formatting: copy the look of every cell as shown in Excel (fill incl. conditional formatting, font color, bold, italic) - up to 1500 cells; larger blocks copy the header row only.
        table_index: with position after_table/before_table.
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells)
    r1, c1, r2, c2 = bounds(rng)
    if (r2 - r1 + 1) * (c2 - c1 + 1) > max_cells:
        raise ToolError(f"Range has more than {max_cells} cells; copy it in parts.")
    _, doc = pick_document(document)
    values = to_grid(rng.Value)
    text = read_grid(rng, "text") if display_text else [["" if v is None else str(v) for v in row] for row in values]
    hidden_rows, hidden_cols = _hidden_lines(ws, r1, c1, r2, c2) if skip_hidden else (set(), set())
    keep_r = [i for i in range(r2 - r1 + 1) if (r1 + i) not in hidden_rows]
    keep_c = [j for j in range(c2 - c1 + 1) if (c1 + j) not in hidden_cols]
    if not keep_r or not keep_c:
        raise ToolError("Nothing to copy: all rows or columns of the range are hidden.")
    grid = [[text[i][j] for j in keep_c] for i in keep_r]
    start = 1 if header_row else 0
    right = []
    for jj, j in enumerate(keep_c, start=1):
        col_vals = [values[i][j] for i in keep_r[start:] if values[i][j] not in (None, "")]
        if col_vals and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in col_vals):
            right.append(jj)
    with _frozen_screen(doc.Application):
        info = build_table(doc, grid, position, paragraph, style, header_row, header_fill, autofit, None, font_size, None, right, table_index)
        formatted = 0
        if keep_formatting:
            formatted = _copy_cell_looks(app, ws, doc.Tables(info["table"]), r1, c1, keep_r, keep_c, header_row)
    return {**info, "source": f"{wb.Name}!{ws.Name}!{addr_of(rng)}", "rows_copied": len(keep_r), "columns_copied": len(keep_c), "hidden_skipped": {"rows": len(hidden_rows), "columns": len(hidden_cols)}, "cells_with_copied_look": formatted}


def _insert_png(doc, data: bytes, position, paragraph, width_cm, caption, alignment, alt_text, table_index=0) -> dict:
    tmpdir = tempfile.mkdtemp(prefix="office_live_")
    try:
        png = os.path.join(tmpdir, "image.png")
        with open(png, "wb") as f:
            f.write(data)
        return insert_picture(doc, png, position, paragraph, width_cm, alt_text, alignment, caption, trusted=True, table_index=table_index)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


@office_tool("bridge", "write", title="Excel range -> Word picture")
def bridge_excel_range_to_word_picture(
    workbook: str,
    cells: str,
    document: str,
    sheet: str = "",
    position: str = "end",
    paragraph: int = 0,
    width_cm: float | None = None,
    caption: str = "",
    alignment: str = "center",
    table_index: int = 0,
) -> dict:
    """Paste an Excel range into Word as a PICTURE that looks exactly as on the Excel screen (fills, borders, conditional formats, fonts). Not editable text - use bridge_excel_range_to_word_table for that. Briefly uses the Windows clipboard.

    Args:
        workbook: Excel workbook ('' = active).
        cells: range to render.
        document: target Word document (required).
        sheet: sheet of the range ('' = active).
        position: 'after_paragraph' / 'before_paragraph' (with `paragraph`) or 'after_table' / 'before_table' (with `table_index` - no need to count paragraphs).
        paragraph: paragraph number for the *_paragraph positions.
        width_cm: picture width (default: fit the page width). caption: optional caption. alignment: left|center|right.
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells)
    data = render_range_png(app, wb, ws, rng)
    _, doc = pick_document(document)
    info = _insert_png(doc, data, position, paragraph, width_cm, caption, alignment, f"{ws.Name}!{addr_of(rng)}", table_index)
    return {**info, "source": f"{wb.Name}!{ws.Name}!{addr_of(rng)}"}


@office_tool("bridge", "write", title="Excel chart -> Word")
def bridge_excel_chart_to_word(
    workbook: str,
    chart_name: str,
    document: str,
    sheet: str = "",
    position: str = "end",
    paragraph: int = 0,
    width_cm: float | None = 15,
    caption: str = "",
    alignment: str = "center",
    table_index: int = 0,
) -> dict:
    """Insert an Excel chart into a Word document as a picture.

    Args:
        workbook: Excel workbook ('' = active).
        chart_name: chart name (see excel_manage_charts action='list').
        document: target Word document (required).
        sheet: sheet of the chart ('' = search all sheets).
        position: 'after_paragraph' / 'before_paragraph' (with `paragraph`) or 'after_table' / 'before_table' (with `table_index` - no need to count paragraphs).
        paragraph: paragraph number for the *_paragraph positions.
        width_cm: picture width. caption: optional caption. alignment: left|center|right.
    """
    from .excel_analysis import _find_chart

    app, wb = pick_workbook(workbook)
    ws, co = _find_chart(wb, chart_name, sheet)
    tmpdir = tempfile.mkdtemp(prefix="office_live_")
    try:
        png = os.path.join(tmpdir, "chart.png")
        try:
            co.Chart.Export(png, "PNG")
        except pywintypes.com_error as exc:
            raise ToolError("Excel could not export the chart: " + com.com_error_text(exc)) from None
        with open(png, "rb") as f:
            data = f.read()
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    _, doc = pick_document(document)
    info = _insert_png(doc, data, position, paragraph, width_cm, caption, alignment, chart_name, table_index)
    return {**info, "source": f"{wb.Name}!{ws.Name}!{co.Name}"}


# ================================================================== Excel -> набор документов Word


def _safe_filename(s: str, fallback: str) -> str:
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", s).strip(" .")
    return s[:80] or fallback


def _unique_stem(stem: str, idx: int, used: set, renamed: list) -> str:
    """Имя файла без коллизий: совпавшее имя различаем номером строки, а если и оно занято — добавляем счётчик.

    Сравнение без учёта регистра (Windows). `used` пополняется окончательным именем; renamed — отчёт о переименованиях.
    """
    wanted = stem
    if stem.lower() in used:
        stem = f"{wanted}_{idx:03d}"
        counter = 2
        while stem.lower() in used:
            stem = f"{wanted}_{idx:03d}_{counter}"
            counter += 1
        renamed.append({"row": idx, "wanted": wanted, "used": stem})
    used.add(stem.lower())
    return stem


def _quit_private_word(wapp) -> None:
    """Закрывает фоновый Word слияния вместе со всеми его документами — и после сбоя тоже."""
    with contextlib.suppress(pywintypes.com_error, ToolError, AttributeError):
        wapp.NormalTemplate.Saved = True  # Normal.dotm обычно занят Word пользователя: выход без вопроса о сохранении
    with contextlib.suppress(pywintypes.com_error, ToolError):
        wapp.Quit(0)


@contextlib.contextmanager
def _foreground_kept():
    """Окно переднего плана (обычно Excel пользователя) возвращается, если за время слияния его забрал Word. По
    возможности: Windows может отказать, а результат слияния от этого не зависит. Сам переключился — не мешаем."""
    try:
        api = Win32()
        foreground = api.foreground()
    except (ToolError, OSError, AttributeError):
        api, foreground = None, 0
    try:
        yield
    finally:
        if api is not None and foreground:
            with contextlib.suppress(ToolError, OSError):
                now = api.foreground()
                if now and now != foreground and api.processes().get(api.info(now)["pid"], {}).get("name", "").lower() == "winword.exe":
                    api.request_foreground(foreground, api.info(foreground)["pid"])


@office_tool("bridge", "save", title="Excel rows -> Word documents (mail merge)")
def bridge_excel_to_word_documents(
    workbook: str,
    cells: str,
    template: str,
    output_dir: str,
    filename_column: str = "",
    sheet: str = "",
    left_delimiter: str = "{{",
    right_delimiter: str = "}}",
    max_documents: int = 100,
    overwrite: bool = False,
    export_pdf: bool = False,
) -> dict:
    """Generate one Word document per Excel row from a Word template containing {{Header}} placeholders (the header names are the first row of `cells`). Typical use: letters, contracts, certificates, acts. Each document is saved into output_dir (and optionally as PDF); the template is never changed.

    Args:
        workbook: Excel workbook ('' = active).
        cells: data block with the header row first, e.g. 'A1:F40'.
        template: full path of the .docx/.dotx template, or the name of an open saved document.
        output_dir: folder for the results (created if missing).
        filename_column: header of the column used for file names (default: numbered files).
        sheet: sheet of the data ('' = active).
        left_delimiter, right_delimiter: placeholder brackets.
        max_documents: safety limit. overwrite: replace existing files. export_pdf: also write a PDF next to each document.

    Documents are generated in a separate hidden Word that quits afterwards (also after a failure): no windows appear, and your Word, its documents and the active window are not touched.
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells)
    n_rows = int(rng.Rows.Count)
    if n_rows < 2:
        raise ToolError("The data block has no rows below the header.")
    if n_rows - 1 > int(max_documents):  # до чтения данных
        raise ToolError(f"{n_rows - 1} rows exceed max_documents={max_documents}. Raise the limit or narrow the range.")
    text = read_grid(rng, "text")
    headers = [str(h).strip() for h in text[0]]
    body = text[1:]
    if filename_column and filename_column not in headers:
        raise ToolError(f"filename_column '{filename_column}' is not a header. Headers: {headers}")
    out_dir = check_path(output_dir, "write")
    if os.path.isfile(template):
        tpl_path = check_path(template, "read")
    else:
        _, tdoc = pick_document(template)
        if not tdoc.Path:
            raise ToolError("The template document has never been saved; save it with word_save_as first.")
        tpl_path = tdoc.FullName

    # --- предварительная проверка ВСЕХ результатов: имена, коллизии, перезапись, открытые файлы — до создания первого документа
    plan, used, renamed, blank_rows = [], set(), [], 0
    for idx, row in enumerate(body, start=1):
        if not any(str(c).strip() for c in row):
            blank_rows += 1
            continue
        values = {h: row[j] for j, h in enumerate(headers) if h}
        stem = _safe_filename(values.get(filename_column, ""), f"document_{idx:03d}") if filename_column else f"document_{idx:03d}"
        stem = _unique_stem(stem, idx, used, renamed)
        docx = check_path(os.path.join(out_dir, stem + ".docx"), "write", WORD_EXTS)
        pdf = check_path(os.path.splitext(docx)[0] + ".pdf", "write", {"pdf"}) if export_pdf else None
        plan.append((idx, values, docx, pdf))
    if not plan:
        raise ToolError("Every row below the header is empty; nothing to generate.")
    existing = [p for _, _, d, pdf in plan for p in (d, pdf) if p and os.path.exists(p)]
    if existing and not overwrite:
        raise ToolError(f"{len(existing)} output file(s) already exist (e.g. {existing[:3]}). Pass overwrite=true or choose another output_dir. Nothing was created.")
    if existing:
        open_now = {os.path.normcase(os.path.abspath(d.FullName)) for _, d in all_documents() if d.Path} if _word_running() else set()
        busy = [p for p in existing if os.path.normcase(os.path.abspath(p)) in open_now]
        if busy:
            raise ToolError(f"These files are open in Word and cannot be overwritten: {busy[:5]}. Close them first. Nothing was created.")
    os.makedirs(out_dir, exist_ok=True)

    created, pdfs, leftovers = [], [], set()
    with _foreground_kept():
        # Отдельный невидимый Word, а не Word пользователя: документы без окон в его Word ведут себя ненадёжно — после
        # того как пользователь закрыл видимый документ, Close каждого второго скрытого документа молча не срабатывает
        # (живая проверка), файлы остаются открытыми. А видимые документы мелькали поверх Excel и забирали фокус (демо).
        wapp = com.private_app("word")
        try:
            wapp.DisplayAlerts = 0
            for idx, values, docx, pdf in plan:
                try:
                    with com.macros_disabled(wapp):
                        d = wapp.Documents.Add(tpl_path, False, 0, False)  # Add(Template, NewTemplate, DocumentType, Visible)
                    try:
                        res = fill_placeholders(d, values, left_delimiter, right_delimiter, "all")
                        leftovers.update(res["unfilled_placeholders"])
                        d.SaveAs2(docx, 16)
                        created.append(os.path.basename(docx))
                        if pdf:
                            d.ExportAsFixedFormat(pdf, 17, False, 0, 0, 0, 0, 0)
                            pdfs.append(os.path.basename(pdf))
                    finally:
                        d.Close(0)
                except Exception as exc:  # noqa: BLE001 — говорим точно, что успело появиться на диске
                    raise ToolError(
                        f"Stopped at data row {idx} ({os.path.basename(docx)}): {com.translate(exc)}. "
                        f"Created before the failure: {len(created)} document(s) {created[:20]}" + (f" and {len(pdfs)} PDF(s)" if pdf else "") + "."
                    ) from None
        finally:
            _quit_private_word(wapp)
    out = {
        "ok": True, "output_dir": out_dir, "documents_created": len(created), "files": created[:50],
        "unfilled_placeholders": sorted(leftovers),
    }
    if pdfs:
        out["pdf_files"] = pdfs[:50]
    if renamed:
        out["renamed_for_uniqueness"] = renamed[:50]
    if blank_rows:
        out["skipped_blank_rows"] = blank_rows
    return out
