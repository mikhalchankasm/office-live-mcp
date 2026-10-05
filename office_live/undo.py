"""Сеансовая отмена: Word UndoRecord и собственные снимки Excel. В стеке нет COM-прокси."""

import contextlib
import ctypes
import hashlib
import json
import os
import re
import time
import uuid
from dataclasses import dataclass, field

import pywintypes

from . import com, config, journal, wd_common as wd, xl_common as xl
from .errors import PartialChangeError, ToolError
from .privacy import redact
from .registry import office_tool
from .util import clean_word_text, parse_a1, to_com_grid, to_grid

MARKER = "__OfficeLiveUndo"
SESSION = uuid.uuid4().hex
STACKS: dict[tuple, list] = {}
_active = None  # доступ сериализован com.LOCK; живёт только внутри run_com


def identity(kind, obj):
    return (xl.workbook_path(obj) if kind == "workbook" else wd.document_path(obj)) or obj.Name


def stack_key(kind, app, obj):
    # Несохранённые Book1 в разных процессах Excel не должны делить историю. У Word.Application нет Hwnd
    # (обращение к нему ломало бы каждую правку в Word), а Word работает одним процессом на пользователя.
    return kind, identity(kind, obj).lower(), int(app.Hwnd) if kind == "workbook" else 0


def _hash(value):
    return hashlib.sha1(json.dumps(value, ensure_ascii=False, default=str).encode("utf-8")).hexdigest()


_PART = re.compile(r'<pkg:part pkg:name="(/word/[^"]+)"(.*?)</pkg:part>', re.S)
# Части, которые меняет пользователь: тело, колонтитулы, примечания, сноски, стили, нумерация. settings/theme/шрифты — нет.
_USER_PARTS = re.compile(r"/word/(document|header\d*|footer\d*|comments\w*|footnotes|endnotes|styles|numbering)\.xml")
# Что Word генерирует сам при каждом чтении WordOpenXML (проверено на живом Word): идентификаторы абзацев и текста
# (w14:paraId/textId), сеансов правки (rsid), пометки фоновой проверки орфографии (proofErr).
_XML_NOISE = re.compile(r'<w:proofErr\b[^>]*/>|\s(?:w14:paraId|w14:textId|w:rsid\w*)="[^"]*"')


def _word_body_xml(doc):
    """Пользовательские части документа из WordOpenXML без шума — без единого обращения к колонтитулам через COM:
    даже Headers(i).Exists создаёт в документе пустые колонтитулы (живая проверка), то есть меняет документ пользователя
    и сам отпечаток. WordOpenXML всей области Content уже содержит колонтитулы, примечания и сноски."""
    xml = doc.Content.WordOpenXML
    parts = [(name, _XML_NOISE.sub("", body)) for name, body in _PART.findall(xml) if _USER_PARTS.fullmatch(name)]
    return parts or _XML_NOISE.sub("", xml)


def word_fingerprint(doc):
    state = [doc.Content.Text, int(doc.Paragraphs.Count), int(doc.Tables.Count), int(doc.InlineShapes.Count)]
    try:
        state.append(_word_body_xml(doc))  # замечает и ручное форматирование: без этого Undo(1) отменил бы правку пользователя
    except (AttributeError, pywintypes.com_error):
        pass
    state.extend([int(doc.Comments.Count), int(doc.Footnotes.Count), int(doc.Endnotes.Count), int(doc.Shapes.Count)])
    return _hash(state)


@contextlib.contextmanager
def _shown(backup):
    """Worksheet.Copy в книгу со скрытым окном Excel отклоняет (живая проверка): на время копирования окно служебной книги
    показываем при выключенной перерисовке экрана и сразу прячем обратно."""
    app, win = backup.Application, backup.Windows(1)
    updating = app.ScreenUpdating
    app.ScreenUpdating = False
    try:
        win.Visible = True
        yield
    finally:
        try:
            win.Visible = False
        finally:  # даже если окно не спряталось, перерисовка экрана у пользователя не должна остаться выключенной
            app.ScreenUpdating = updating


@contextlib.contextmanager
def _quiet(app):
    old = app.DisplayAlerts
    app.DisplayAlerts = False
    try:
        yield
    finally:
        app.DisplayAlerts = old


@contextlib.contextmanager
def _internal_events(app):
    old = app.EnableEvents
    xl.suspend_events(app, force=True)
    try:
        yield
    finally:
        app.EnableEvents = old


@contextlib.contextmanager
def _keep_sheet(wb):
    app = wb.Application
    previous_book, previous = app.ActiveWorkbook, app.ActiveSheet
    try:
        yield
    finally:
        if previous is not None:
            previous.Activate()
        elif previous_book is not None:
            previous_book.Activate()


def _process_alive(pid):
    """При отказе доступа считаем процесс живым. На Windows os.kill(pid, 0) небезопасен."""
    from ctypes import wintypes

    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
    kernel.OpenProcess.restype = wintypes.HANDLE
    kernel.GetExitCodeProcess.argtypes = (wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD))
    kernel.GetExitCodeProcess.restype = wintypes.BOOL
    kernel.CloseHandle.argtypes = (wintypes.HANDLE,)
    kernel.CloseHandle.restype = wintypes.BOOL
    handle = kernel.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    if not handle:
        return ctypes.get_last_error() != 87  # ERROR_INVALID_PARAMETER: такого PID нет
    try:
        code = wintypes.DWORD()
        return not kernel.GetExitCodeProcess(handle, ctypes.byref(code)) or code.value == 259  # STILL_ACTIVE
    finally:
        kernel.CloseHandle(handle)


def _close_orphan_backups(app):
    """Если в экземпляре Excel не осталось книг пользователя, закрываем наши служебные книги (и книги умерших серверов):
    отменять там больше нечего, а скрытая книга держала бы Excel в памяти без единого окна (живая проверка)."""
    books = [app.Workbooks(i) for i in range(1, int(app.Workbooks.Count) + 1)]
    if not books or not all(xl.is_undo_workbook(wb) for wb in books):
        return
    for wb in books:
        owner = re.fullmatch(r'="([0-9]+):([0-9a-f]{32}):[0-9a-f]{32}"', str(wb.Names(MARKER).RefersTo))
        if owner is None:
            continue
        pid, session = int(owner[1]), owner[2]
        if (pid == os.getpid() and session == SESSION) or not _process_alive(pid):
            with _internal_events(app), _quiet(app):
                wb.Close(False)


def _backup(app, token=None):
    """Находим по метке, а не по имени: закрытую книгу нельзя подменить новой одноимённой."""
    if token is None:
        xl.suspend_events(app, force=True)
    found = None
    for i in range(int(app.Workbooks.Count), 0, -1):
        wb = app.Workbooks(i)
        if xl.is_undo_workbook(wb):
            value = str(wb.Names(MARKER).RefersTo)
            owner = re.fullmatch(r'="([0-9]+):([0-9a-f]{32}):[0-9a-f]{32}"', value)
            if owner is None:
                continue  # старую/повреждённую метку нельзя безопасно приписать умершему процессу
            pid, session = int(owner[1]), owner[2]
            if pid == os.getpid() and session == SESSION:
                if token is None or value == token:
                    found = wb
            elif not _process_alive(pid):
                with _internal_events(app), _quiet(app):
                    wb.Close(False)
    if found is not None:
        return found
    if token is not None:
        raise ToolError("The hidden undo backup workbook was closed or is unavailable.")
    previous = app.ActiveWorkbook
    wb = None
    try:
        wb = app.Workbooks.Add(-4167)  # одна пустая таблица; не пользовательский шаблон
        wb.Names.Add(MARKER, f'="{os.getpid()}:{SESSION}:{uuid.uuid4().hex}"', False)
        wb.Windows(1).Visible = False
        wb.Saved = True
        return wb
    except Exception:
        if wb is not None:
            with _quiet(app):
                wb.Close(False)
        raise
    finally:
        if previous is not None:
            previous.Activate()


def _area(rng):
    return {"sheet": rng.Worksheet.Name, "address": xl.addr_of(rng)}


def _range(wb, area):
    return wb.Worksheets(area["sheet"]).Range(area["address"])


def _size(rng):
    return int(rng.Rows.Count) * int(rng.Columns.Count)


def _dimensions(rng, columns=True, rows=True, *, explicit=False):
    r1, c1, r2, c2 = xl.bounds(rng)
    ws = rng.Worksheet
    # Снимок целых строк не меняет ширины столбцов, и наоборот. Явная правка размеров — исключение.
    if not explicit:
        columns = columns and c2 - c1 + 1 < int(ws.Columns.Count)
        rows = rows and r2 - r1 + 1 < int(ws.Rows.Count)
    return {
        "widths": [(i, ws.Columns(i).ColumnWidth) for i in range(c1, c2 + 1)] if columns else [],
        "heights": [(i, ws.Rows(i).RowHeight) for i in range(r1, r2 + 1)] if rows else [],
    }


def _restore_dimensions(ws, data):
    for index, value in data.get("widths", ()):
        ws.Columns(index).ColumnWidth = value
    for index, value in data.get("heights", ()):
        ws.Rows(index).RowHeight = value


def _formula_state(rng):
    # Formula теряет различие между массивом и неявным пересечением (@).
    try:
        return "Formula2", to_grid(rng.Formula2)
    except (AttributeError, pywintypes.com_error):
        return "Formula", to_grid(rng.Formula)


def _format_state(rng):
    # По одному чтению каждого свойства всего диапазона; None для смешанного оформления допустим.
    font, interior = rng.Font, rng.Interior
    return [rng.NumberFormat, font.Name, font.Size, font.Bold, font.Italic, font.Underline, font.Color,
            interior.Color, interior.Pattern, rng.HorizontalAlignment, rng.VerticalAlignment, rng.WrapText, rng.MergeCells,
            font.Strikethrough, font.Subscript, font.Superscript, font.OutlineFont, font.Shadow, font.TintAndShade,
            interior.PatternColor, interior.TintAndShade, interior.PatternTintAndShade,
            rng.ShrinkToFit, rng.IndentLevel, rng.Orientation, rng.ReadingOrder, rng.AddIndent, rng.Locked, rng.FormulaHidden,
            [[rng.Borders(i).LineStyle, rng.Borders(i).Weight, rng.Borders(i).Color,
              *_optional_state(rng.Borders(i), ("ThemeColor", "TintAndShade"))] for i in range(5, 13)],
            _optional_state(font, ("ThemeColor",)), _optional_state(interior, ("ThemeColor", "PatternThemeColor")),
            _property_state(rng, "Style")]


def _optional_state(obj, names):
    out = []
    for name in names:
        try:
            out.append(getattr(obj, name))
        except (pywintypes.com_error, AttributeError):
            out.append(None)  # например Formula2 отсутствует у правила с одним операндом
    return out


def _metadata_state(rng, mode="all"):
    """Метаданные, которые переносит Copy/PasteSpecial; читаем только выбранный диапазон."""
    state = []
    if mode in {"all", "formats"}:
        rules = rng.FormatConditions
        for i in range(1, int(rules.Count) + 1):
            rule = rules(i)
            state.append([int(rule.Type), xl.addr_of(rule.AppliesTo),
                          _optional_state(rule, ("Priority", "StopIfTrue", "Formula1", "Formula2", "Operator", "Text", "TextOperator", "DateOperator",
                                                 "Rank", "Percent", "TopBottom", "AboveBelow", "NumStdDev", "DupeUnique"))])
            for member, props in (("Font", ("Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Color")),
                                  ("Interior", ("Color", "Pattern", "PatternColor"))):
                try:
                    state.append(_optional_state(getattr(rule, member), props))
                except (pywintypes.com_error, AttributeError):
                    state.append(None)
    if mode in {"all", "comments", "validation"}:
        for cell_type, member in ((-4144, "Comment"), (-4174, "Validation")):
            if mode != "all" and member.lower() != mode.removesuffix("s"):
                continue
            try:
                selected = rng.SpecialCells(cell_type)
            except pywintypes.com_error:
                continue  # Excel: no cells found
            selected = rng.Worksheet.Parent.Application.Intersect(rng, selected)
            if selected is None:
                continue
            r1, c1, r2, c2 = xl.bounds(rng)
            for cell in selected.Cells:
                # SpecialCells у одиночной ячейки может искать по всему UsedRange.
                if not (r1 <= int(cell.Row) <= r2 and c1 <= int(cell.Column) <= c2):
                    continue
                obj = getattr(cell, member)
                if member == "Comment":
                    value = [obj.Text(), obj.Author, obj.Visible]
                else:
                    value = _optional_state(obj, ("Type", "AlertStyle", "Operator", "Formula1", "Formula2", "IgnoreBlank", "InCellDropdown",
                                                   "ShowError", "ErrorTitle", "ErrorMessage", "ShowInput", "InputTitle", "InputMessage"))
                state.append([xl.addr_of(cell), member, value])
    if mode == "all":
        links = rng.Hyperlinks
        state.append([[xl.addr_of(links(i).Range), links(i).Address, links(i).SubAddress, links(i).ScreenTip, links(i).TextToDisplay]
                      for i in range(1, int(links.Count) + 1)])
    return state


def _property_object(rng, path):
    parts = path.split(".")
    for name in parts[:-1]:
        rng = rng.Borders(int(name[7:])) if name.startswith("Borders") else getattr(rng, name)
    return rng, parts[-1]


def _property_state(rng, path):
    obj, name = _property_object(rng, path)
    try:
        value = getattr(obj, name)
    except pywintypes.com_error:
        if name in {"ThemeColor", "PatternThemeColor"}:
            return {"undefined": True}  # явный RGB не имеет номера цвета темы
        raise
    if path == "Style" and value is not None and not isinstance(value, str):
        value = value.Name
    if value is not None:
        return {"value": value}
    # Смешанное свойство восстанавливается по ячейкам, а не присваиванием COM Null всему диапазону.
    r1, c1, r2, c2 = xl.bounds(rng)
    cells = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            obj, name = _property_object(rng.Worksheet.Cells(r, c), path)
            value = getattr(obj, name)
            if path == "Style" and value is not None and not isinstance(value, str):
                value = value.Name
            cells.append((r, c, value))
    return {"cells": cells}


def _restore_properties(rng, properties):
    for path, data in properties.items():
        if "undefined" in data:
            continue  # RGB/ColorIndex восстанавливается отдельным свойством
        if "value" in data:
            obj, name = _property_object(rng, path)
            setattr(obj, name, data["value"])
        else:
            for r, c, value in data["cells"]:
                obj, name = _property_object(rng.Worksheet.Cells(r, c), path)
                setattr(obj, name, value)


def _repair_formulas(rng, formulas, formula_property="Formula"):
    """Copy между книгами меняет ссылки; константы не присваиваем через Formula (код '007')."""
    current = to_grid(getattr(rng, formula_property))
    r1, c1, _, _ = xl.bounds(rng)
    for i, row in enumerate(formulas):
        for j, value in enumerate(row):
            if isinstance(value, str) and value.startswith("=") and value != current[i][j]:
                setattr(rng.Worksheet.Cells(r1 + i, c1 + j), formula_property, value)


@dataclass
class Entry:
    tool: str
    where: str
    kind: str
    stamp: str = field(default_factory=lambda: time.strftime("%Y-%m-%d %H:%M:%S"))
    ops: list = field(default_factory=list)
    areas: list = field(default_factory=list)
    used: list = field(default_factory=list)
    structural: bool = False
    reason: str = ""
    force_reason: str = ""
    fingerprint: str = ""
    before_fingerprint: str = ""
    refresh_previous: bool = True
    backup_token: str = ""
    backup_sheets: list = field(default_factory=list)
    warning: str = ""
    transaction: str = ""
    peers: list = field(default_factory=list)

    @property
    def undoable(self):
        return not self.reason

    def barrier(self, reason):
        self.reason = str(reason)[:300]


def _snapshots(app, wb, entry, ranges, *, columns=False, rows=False, restore="all", properties=()):
    ranges = list(ranges)
    if any(journal.owns_sheet(rng.Worksheet) or (rng.Worksheet.Name == journal.LOG_SHEET and journal.sheet_enabled(wb)) for rng in ranges):
        entry.barrier("The service log sheet is not recorded by undo.")
        return
    size = sum(_size(rng) for rng in ranges)
    if size > config.SETTINGS.undo_max_cells:
        entry.barrier(f"{size} cells exceeds OFFICE_LIVE_UNDO_MAX_CELLS={config.SETTINGS.undo_max_cells}")
        return
    with _internal_events(app), _keep_sheet(wb):
        backup = _backup(app)
        entry.backup_token = str(backup.Names(MARKER).RefersTo)
        try:
            for rng in ranges:
                prop, formulas = _formula_state(rng)
                copy_columns = entry.tool == "excel_copy_range" and restore == "all" and int(rng.Rows.Count) == int(rng.Worksheet.Rows.Count)
                copy_rows = entry.tool == "excel_copy_range" and restore == "all" and int(rng.Columns.Count) == int(rng.Worksheet.Columns.Count)
                data = {**_area(rng), **_dimensions(rng, columns or copy_columns, rows or copy_rows), "formulas": formulas, "formula_property": prop,
                        "restore": restore, "properties": {p: _property_state(rng, p) for p in properties}}
                saved = backup.Worksheets.Add(None, backup.Sheets(backup.Sheets.Count))
                entry.backup_sheets.append(saved.Name)
                data["backup_sheet"] = saved.Name
                rng.Copy(saved.Range(data["address"]))  # Destination, без буфера обмена
                entry.ops.append({"op": "range", **data})
                entry.areas.append(_area(rng))
        finally:
            backup.Saved = True


def _drop(app, entry):
    if not entry.backup_sheets:
        return
    try:
        backup = _backup(app, entry.backup_token)
        xl.suspend_events(app, force=True)
        with _quiet(app):
            try:
                for name in entry.backup_sheets:
                    backup.Sheets(name).Delete()
            finally:
                backup.Saved = True
    except Exception:  # noqa: BLE001 — книга могла быть закрыта пользователем
        pass
    entry.backup_sheets.clear()


def _view(app, wb, sheet, value=None):
    from .excel_format import _remember_view, _restore_view

    previous = _remember_view(app)
    try:
        wb.Activate()
        wb.Worksheets(sheet).Activate()
        win = app.ActiveWindow
        props = ("SplitRow", "SplitColumn", "FreezePanes", "Zoom", "DisplayGridlines", "DisplayHeadings", "ScrollRow", "ScrollColumn")
        if value is None:
            return {p: getattr(win, p) for p in props}
        win.FreezePanes = False
        for p in props:
            setattr(win, p, value[p])
    finally:
        _restore_view(app, previous)


def _forget_marker(wb):
    """Удаляет перенесённую копированием метку служебной книги (уровня книги и уровня листа) из книги пользователя."""
    names = wb.Names
    for i in range(int(names.Count), 0, -1):
        item = names(i)
        if item.Name == MARKER or str(item.Name).endswith("!" + MARKER):
            item.Delete()


def _name_state(wb, name):
    try:
        item = wb.Names(name)
        return [item.RefersTo, bool(item.Visible)]
    except (pywintypes.com_error, KeyError):
        return None


def _sheet_content_state(ws):
    """Счётчики по всему листу: выбранные свойства объектов не покрывают все ручные правки."""
    cells = ws.Cells
    state = {"shapes": int(ws.Shapes.Count), "tables (ListObjects)": int(ws.ListObjects.Count),
             "pivot tables": int(ws.PivotTables().Count), "conditional formatting": int(cells.FormatConditions.Count),
             "comments/notes": int(ws.Comments.Count)}
    try:
        cells.SpecialCells(-4174)  # xlCellTypeAllValidation; наличие, без перебора ячеек всего листа
        state["data validation"] = True
    except pywintypes.com_error:
        state["data validation"] = False  # Excel: no cells found
    try:
        state["threaded comments"] = int(ws.CommentsThreaded.Count)
    except (AttributeError, pywintypes.com_error):
        state["threaded comments"] = None  # старые версии Excel не поддерживают обсуждения
    return state


def _sheet_state(ws, rng):
    # Перед удалением созданного/скопированного листа проверяем и правки без изменения числа объектов.
    cells = ws.Cells
    dimensions = [cells.ColumnWidth, cells.RowHeight, ws.StandardWidth, _dimensions(rng, explicit=True)]
    shapes = []
    collection = ws.Shapes
    for i in range(1, int(collection.Count) + 1):
        shape = collection(i)
        kind = int(shape.Type)
        shapes.append([shape.Name, kind, shape.Left, shape.Top, shape.Width, shape.Height, shape.Rotation,
                       shape.Chart.ChartType if kind == 3 else None])  # msoChart
    tables = []
    collection = ws.ListObjects
    for i in range(1, int(collection.Count) + 1):
        table = collection(i)
        style = table.TableStyle
        if style is not None and not isinstance(style, str):
            style = style.Name
        tables.append([table.Name, xl.addr_of(table.Range), style, bool(table.ShowTotals), bool(table.ShowHeaders)])
    # Formula1/Formula2 и другие свойства существуют не у всех видов правил: чтение защищено в _metadata_state.
    return [dimensions, shapes, tables, _metadata_state(rng, "formats"), _sheet_content_state(ws)]


def excel_fingerprint(app, wb, entry):
    state = []
    if entry.structural:
        # Создание/удаление служебного журнала между вызовами не является чужой правкой.
        state.append([name for name in xl.sheet_names(wb, any_type=True)
                      if name.casefold() != journal.LOG_SHEET.casefold() or not journal.owns_sheet(wb.Sheets(name))])
    for area in entry.areas:
        rng = _range(wb, area)
        ops = [op for op in entry.ops if op["op"] == "range" and _area(rng) == {k: op[k] for k in ("sheet", "address")}]
        modes = {op.get("restore", "all") for op in ops} or {"all"}
        state.append([area, _formula_state(rng) if modes & {"all", "contents"} else None,
                      _format_state(rng) if modes & {"all", "formats"} else None])
        if entry.tool in {"excel_clean_text", "excel_split_column"}:
            state.append([_property_state(rng, p) for p in ("NumberFormat", "PrefixCharacter")])
        for mode in modes & {"all", "formats", "comments", "validation"}:
            state.append(_metadata_state(rng, mode))
    for name in entry.used:
        ws = wb.Worksheets(name)
        rng = ws.UsedRange
        if _size(rng) > config.SETTINGS.undo_max_cells:
            raise ToolError("The affected sheet exceeds the undo fingerprint size limit.")
        state.append([name, xl.addr_of(rng), _formula_state(rng), _format_state(rng), _sheet_state(ws, rng)])
    for op in entry.ops:
        what = op["op"]
        if what in {"dimensions", "range"}:
            rng = _range(wb, op)
            state.append(_dimensions(rng, bool(op["widths"]), bool(op["heights"]), explicit=True))
            if what == "range":
                state.append({p: _property_state(rng, p) for p in op.get("properties", {})})
        elif what == "hidden":
            ws = wb.Worksheets(op["sheet"])
            coll = ws.Rows if op["axis"] == "rows" else ws.Columns
            state.append([bool(coll(i).Hidden) for i, _ in op["values"]])
        elif what == "name":
            state.append(_name_state(wb, op["name"]))
        elif what == "sheet_property":
            ws = wb.Sheets(op["current"])
            state.append([ws.Name, int(ws.Index), int(ws.Visible), ws.Tab.ColorIndex, ws.Tab.Color])
        elif what == "calculation":
            state.append(app.Calculation)
        elif what == "view":
            state.append(_view(app, wb, op["sheet"]))
        elif what == "protection":
            from .excel_format import _protection_state

            target = wb.Worksheets(op["sheet"]) if op["scope"] == "sheet" else wb
            state.append(_protection_state(target, op["scope"]))
            state.extend(_property_state(target.Range(a), "Locked") for a in op["locked"])
    return _hash(state)


def _simple_range(app, wb, entry, args):
    _, rng = xl.get_range(wb, args.get("sheet", ""), args.get("cells", args.get("cell", "")))
    if entry.tool == "excel_clear_range" and args.get("what", "contents").lower() == "contents":
        _snapshots(app, wb, entry, [rng], restore="contents")
    elif entry.tool == "excel_manage_comments" or (entry.tool == "excel_clear_range" and args.get("what", "").lower() == "comments"):
        _snapshots(app, wb, entry, [rng], restore="comments")
    elif entry.tool == "excel_data_validation":
        _snapshots(app, wb, entry, [rng], restore="validation")
    elif entry.tool == "excel_conditional_format" or (entry.tool == "excel_clear_range" and args.get("what", "").lower() == "formats"):
        _snapshots(app, wb, entry, [rng], restore="formats")
    elif entry.tool == "excel_format_range" and args.get("style") is None:
        mapping = {"bold": "Font.Bold", "italic": "Font.Italic", "underline": "Font.Underline", "strikethrough": "Font.Strikethrough",
                   "font_name": "Font.Name", "font_size": "Font.Size", "number_format": "NumberFormat",
                   "horizontal_alignment": "HorizontalAlignment", "vertical_alignment": "VerticalAlignment", "wrap_text": "WrapText",
                   "shrink_to_fit": "ShrinkToFit", "indent": "IndentLevel", "text_rotation": "Orientation"}
        properties = [p for key, p in mapping.items() if args.get(key) is not None]
        if args.get("font_color") is not None:
            properties += ["Font.Color", "Font.ThemeColor", "Font.TintAndShade"]
        if args.get("fill_color") is not None:
            properties += ["Interior.Color", "Interior.ThemeColor", "Interior.TintAndShade", "Interior.Pattern"]
        if args.get("borders") is not None:
            from .excel_format import EDGES, _border_plan

            edges = _border_plan(rng, args["borders"].lower(), args.get("border_style", "thin").lower(), args.get("border_color", "#000000"))[0]
            properties += [f"Borders{EDGES[edge]}.{p}" for edge in edges for p in ("Color", "ThemeColor", "TintAndShade", "Weight", "LineStyle")]
        _snapshots(app, wb, entry, [rng], restore="properties", properties=properties)
    elif entry.tool == "excel_format_range":
        _snapshots(app, wb, entry, [rng], restore="formats")
    else:
        _snapshots(app, wb, entry, [rng])


def _write_range(app, wb, entry, args):
    from .excel_core import _prepare_target

    ws, rng = xl.get_range(wb, args["sheet"], args["cells"], empty_means_used=False)
    value = args["values"] if entry.tool == "excel_write_range" else args["formula"]
    if entry.tool == "excel_write_range" or not isinstance(value, str):
        grid = to_com_grid(value)
        rng, _ = _prepare_target(ws, rng, len(grid), len(grid[0]))
    _snapshots(app, wb, entry, [rng], restore="contents", properties=("NumberFormat",))


def _clean_text(app, wb, entry, args):
    _, rng = xl.bounded_range(app, wb, args["sheet"], args["cells"], 100000)
    if rng is not None:
        _snapshots(app, wb, entry, [rng], restore="contents", properties=("NumberFormat",))


def _split_column(app, wb, entry, args):
    rng = wb.Worksheets(args["_output_sheet"]).Range(args["_output_range"])
    _snapshots(app, wb, entry, [rng], restore="contents", properties=("NumberFormat",))


def _protection(app, wb, entry, args):
    from .excel_format import PASSWORD_BARRIER, _protection_state

    if args.get("_has_secret"):
        entry.barrier(PASSWORD_BARRIER)
        return
    scope = args.get("scope", "sheet")
    target = wb.Worksheets(args["_protection_sheet"]) if scope == "sheet" else wb
    ranges = args.get("_locked_ranges", [])
    if sum(_size(target.Range(a)) for a in ranges) > config.SETTINGS.undo_max_cells:
        entry.barrier("Locked snapshot exceeds OFFICE_LIVE_UNDO_MAX_CELLS")
        return
    entry.ops.append({"op": "protection", "scope": scope, "sheet": args["_protection_sheet"],
                      "state": _protection_state(target, scope), "locked": {a: _property_state(target.Range(a), "Locked") for a in ranges}})


def _autofill(app, wb, entry, args):
    ws, src = xl.get_range(wb, args["sheet"], args["source"], empty_means_used=False)
    dest = ws.Range(args["dest"])
    r1, c1, r2, c2 = xl.bounds(src)
    d1, e1, d2, e2 = xl.bounds(dest)
    _snapshots(app, wb, entry, [xl.sub_range(ws, min(r1, d1), min(c1, e1), max(r2, d2), max(c2, e2))])


def _replace(app, wb, entry, args):
    # Сам инструмент обрабатывает только выбранный (или активный) лист.
    ws, rng = xl.get_range(wb, args.get("sheet", ""), args.get("cells", ""))
    rng = xl.clip_to_used(app, ws, rng)
    if rng is not None:
        _snapshots(app, wb, entry, [rng], restore="contents", properties=("NumberFormat",))


def _dimensions_plan(app, wb, entry, args):
    ws, rng = xl.get_range(wb, args["sheet"], args.get("cells", ""))
    if not args.get("cells"):
        rng = ws.Cells
    columns = args.get("column_width") is not None or args.get("autofit", "").lower() in {"columns", "both"}
    rows = args.get("row_height") is not None or args.get("autofit", "").lower() in {"rows", "both"}
    count = (int(rng.Columns.Count) if columns else 0) + (int(rng.Rows.Count) if rows else 0)
    if count > config.SETTINGS.undo_max_cells:
        entry.barrier(f"{count} dimensions exceeds OFFICE_LIVE_UNDO_MAX_CELLS")
        return
    entry.ops.append({"op": "dimensions", **_area(rng), **_dimensions(rng, columns, rows, explicit=True)})


def _hidden(app, wb, entry, args):
    from .excel_core import _line_spec

    axis = args.get("axis", "rows").lower()
    if args.get("where"):
        ws, rng = xl.get_range(wb, args.get("sheet", ""), args.get("cells", ""))
        rng = xl.clip_to_used(app, ws, rng)
        if rng is None:
            return
        r1, _, r2, _ = xl.bounds(rng)
        spans = [(r1, r2)]
    else:
        ws = xl.pick_sheet(wb, args.get("sheet", ""))
        spans = _line_spec(args["lines"], axis)
    count = sum(b - a + 1 for a, b in spans)
    if count > config.SETTINGS.undo_max_cells:
        entry.barrier(f"{count} rows/columns exceeds OFFICE_LIVE_UNDO_MAX_CELLS")
        return
    coll = ws.Rows if axis == "rows" else ws.Columns
    values = [(i, bool(coll(i).Hidden)) for a, b in spans for i in range(a, b + 1)]
    entry.ops.append({"op": "hidden", "sheet": ws.Name, "axis": axis, "values": values})


def _rows_columns(app, wb, entry, args):
    from .excel_core import _rows_cols_range

    ws = xl.pick_sheet(wb, args["sheet"])
    rng = _rows_cols_range(ws, args["axis"], args["start"], args.get("count", 1))
    entry.structural = True
    entry.used = [ws.Name]
    deleting = entry.tool == "excel_delete_rows_columns"
    entry.ops.append({"op": "insert_lines" if deleting else "delete_lines", **_area(rng)})
    if deleting:
        _snapshots(app, wb, entry, [rng], columns=args["axis"].lower() == "columns", rows=args["axis"].lower() == "rows")
        entry.warning = f"Undo restores the deleted {args['axis'].lower()}, but formulas elsewhere that referred to them stay #REF!."


def _add_sheet(app, wb, entry, args):
    entry.structural = True
    entry.ops.append({"op": "delete_created_sheet", "before": xl.sheet_names(wb, any_type=True)})


def _delete_sheet(app, wb, entry, args):
    ws = xl.pick_sheet(wb, args["sheet"], any_type=True)
    if journal.owns_sheet(ws) or (ws.Name == journal.LOG_SHEET and journal.sheet_enabled(wb)):
        entry.barrier("The service log sheet is not recorded by undo.")
        return
    rng = ws.UsedRange
    if _size(rng) > config.SETTINGS.undo_max_cells:
        entry.barrier(f"{_size(rng)} cells exceeds OFFICE_LIVE_UNDO_MAX_CELLS")
        return
    prop, formulas = _formula_state(rng)
    data = {"op": "sheet_copy", "name": ws.Name, "index": int(ws.Index), "visible": int(ws.Visible),
            "address": xl.addr_of(rng), "formulas": formulas, "formula_property": prop}
    with _internal_events(app), _keep_sheet(wb):
        backup = _backup(app)
        entry.backup_token = str(backup.Names(MARKER).RefersTo)
        try:
            with _shown(backup):
                ws.Copy(None, backup.Sheets(backup.Sheets.Count))
            saved = backup.Sheets(backup.Sheets.Count)
            entry.backup_sheets.append(saved.Name)
            unique = "Undo_" + uuid.uuid4().hex[:20]
            saved.Name = unique
            entry.backup_sheets[-1] = unique
            data["backup_sheet"] = unique
            entry.ops.append(data)
            entry.structural = True
            entry.warning = "Formulas in other sheets that referred to the deleted sheet remain #REF!."
        finally:
            backup.Saved = True


def _manage_sheet(app, wb, entry, args):
    action = args["action"].lower()
    if action == "copy":
        _add_sheet(app, wb, entry, args)
        return
    if action not in {"rename", "move", "tab_color", "hide", "very_hide", "unhide"}:
        entry.barrier("This sheet operation has no content undo.")
        return
    ws = xl.pick_sheet(wb, args["sheet"], any_type=True)
    entry.structural = True
    entry.ops.append({"op": "sheet_property", "name": ws.Name, "current": ws.Name, "index": int(ws.Index),
                      "visible": int(ws.Visible), "color_index": ws.Tab.ColorIndex, "color": ws.Tab.Color,
                      "action": action})


def _names(app, wb, entry, args):
    entry.ops.append({"op": "name", "name": args["name"], "value": _name_state(wb, args["name"])})


def _shape(app, wb, entry, args):
    ws = xl.pick_sheet(wb, args.get("sheet", ""))
    entry.ops.append({"op": "shape", "sheet": ws.Name})
    entry.structural = True


def _slicer(app, wb, entry, args):
    if args["action"].lower() != "add":
        entry.barrier("This slicer operation has no inverse.")
    else:
        entry.ops.append({"op": "slicer"})


def _table(app, wb, entry, args):
    if args["action"].lower() not in {"add", "create"}:
        entry.barrier("Only table creation can be undone.")
        return
    ws, rng = xl.get_range(wb, args.get("sheet", ""), args["cells"], empty_means_used=False)
    entry.ops.append({"op": "table", "sheet": ws.Name})
    # Без заголовка Excel вставляет строку; итоговая строка расширяет таблицу.
    extra = int(not args.get("has_headers", True)) + int(bool(args.get("show_totals")))
    if extra:
        r1, c1, r2, c2 = xl.bounds(rng)
        # Создание заголовка/итога может сдвинуть данные ниже таблицы в тех же столбцах.
        _, _, used_bottom, _ = xl.bounds(ws.UsedRange)
        rng = xl.sub_range(ws, r1, c1, max(r2, used_bottom) + extra, c2)
    _snapshots(app, wb, entry, [rng])


def _sheet_view(app, wb, entry, args):
    ws = xl.pick_sheet(wb, args.get("sheet", ""))
    entry.ops.append({"op": "view", "sheet": ws.Name, "value": _view(app, wb, ws.Name)})


def _calculate(app, wb, entry, args):
    if args.get("mode"):
        entry.ops.append({"op": "calculation", "value": app.Calculation})


def _bridge(app, wb, entry, args):
    from .excel_core import _prepare_target
    from .word_tables import get_table, table_grid

    sheet = args.get("sheet", "")
    if args.get("create_sheet") and sheet and sheet.lower() not in [n.lower() for n in xl.sheet_names(wb, any_type=True)]:
        _add_sheet(app, wb, entry, args)
        return
    _, doc = wd.pick_document(args["document"])
    if entry.tool == "bridge_word_table_to_excel":
        grid = table_grid(get_table(doc, args["table_index"]))
        if grid["uniform"]:
            rows, cols = len(grid["values"]), max(map(len, grid["values"]))
        else:
            rows, cols = grid["rows"], max(c["col"] for c in grid["cells"])
    else:
        first = max(1, int(args.get("start_paragraph", 1)))
        last = min(int(doc.Paragraphs.Count), int(args.get("end_paragraph", 0)) or int(doc.Paragraphs.Count), first + 4999)
        rows, cols = 1, 3
        for i in range(first, last + 1):
            p = doc.Paragraphs(i)
            if args.get("headings_only") and int(p.OutlineLevel) >= 10:
                continue
            if clean_word_text(p.Range.Text).strip():
                rows += 1
    ws = xl.pick_sheet(wb, args.get("sheet", ""))
    rng, _ = _prepare_target(ws, ws.Range(args.get("top_left", "A1")), rows, cols)
    _snapshots(app, wb, entry, [rng], restore="contents", properties=("NumberFormat",))
    if entry.tool == "bridge_word_text_to_excel" or args.get("header_bold", True):
        r1, c1, _, c2 = xl.bounds(rng)
        _snapshots(app, wb, entry, [xl.sub_range(ws, r1, c1, r1, c2)], restore="properties", properties=("Font.Bold",))


RESOLVERS = dict.fromkeys((
    "excel_clear_range", "excel_format_range", "excel_merge_cells", "excel_conditional_format", "excel_data_validation",
    "excel_manage_comments", "excel_add_hyperlink", "excel_sort_range", "excel_remove_duplicates",
), _simple_range)
RESOLVERS.update({
    "excel_clean_text": _clean_text,
    "excel_split_column": _split_column, "excel_protection": _protection,
    "excel_write_range": _write_range, "excel_set_formula": _write_range, "excel_autofill": _autofill,
    "excel_replace": _replace, "excel_set_dimensions": _dimensions_plan, "excel_hide_rows_columns": _hidden,
    "excel_insert_rows_columns": _rows_columns, "excel_delete_rows_columns": _rows_columns,
    "excel_add_worksheet": _add_sheet, "excel_delete_sheet": _delete_sheet, "excel_manage_sheet": _manage_sheet,
    "excel_manage_names": _names, "excel_create_chart": _shape, "excel_insert_image": _shape,
    "excel_manage_slicers": _slicer, "excel_manage_tables": _table, "excel_sheet_view": _sheet_view,
    "excel_calculate": _calculate, "bridge_word_table_to_excel": _bridge, "bridge_word_text_to_excel": _bridge,
})


def _finish_excel(wb, entry, result):
    for op in entry.ops:
        what = op["op"]
        if what == "range":
            rng = _range(wb, op)
            op["properties"] = {p: before for p, before in op.get("properties", {}).items() if before != _property_state(rng, p)}
        elif what == "delete_created_sheet":
            before = op.pop("before")
            created = [n for n in xl.sheet_names(wb, any_type=True) if n not in before]
            if len(created) != 1:
                raise ToolError("Could not identify the created sheet.")
            op["name"] = created[0]
            entry.used.append(created[0])
            content = [name for name, count in _sheet_content_state(wb.Worksheets(created[0])).items() if count]
            if content:
                entry.force_reason = (f"the created sheet '{created[0]}' contains {', '.join(content)} whose manual edits "
                                      "cannot be detected reliably; pass force=true to delete it anyway.")
        elif what == "sheet_property" and op["action"] == "rename":
            op["current"] = next(n for n in xl.sheet_names(wb, any_type=True) if n == result.get("new_name", "")) if result.get("new_name") else wb.Sheets(op["index"]).Name
        elif what == "shape":
            op["name"] = result.get("shape") or result["chart"]
        elif what == "slicer":
            op["name"] = result["cache"]
        elif what == "table":
            op["name"] = result["table"]


def restore_excel(app, wb, entry):
    xl.suspend_events(app, force=True)
    backup = _backup(app, entry.backup_token) if entry.backup_sheets else None
    for op in entry.ops:
        what = op["op"]
        if what == "range":
            target = _range(wb, op)
            mode = op.get("restore", "all")
            if mode == "all":
                target.UnMerge()  # снимок сам вернёт старые объединения; текущие мешают Copy
                backup.Worksheets(op["backup_sheet"]).Range(op["address"]).Copy(target)
            elif mode in {"contents", "formats", "comments", "validation"}:
                backup.Worksheets(op["backup_sheet"]).Range(op["address"]).Copy()
                try:
                    paste = {"contents": -4123, "formats": -4122, "comments": -4144, "validation": 6}[mode]
                    if mode == "comments":
                        target.ClearComments()
                    elif mode == "validation":
                        target.Validation.Delete()
                    target.PasteSpecial(paste, -4142, False, False)
                finally:
                    app.CutCopyMode = False
            _restore_properties(target, op.get("properties", {}))
            _restore_dimensions(target.Worksheet, op)
            if mode in {"all", "contents"}:
                _repair_formulas(target, op["formulas"], op["formula_property"])
        elif what == "dimensions":
            _restore_dimensions(wb.Worksheets(op["sheet"]), op)
        elif what == "hidden":
            ws = wb.Worksheets(op["sheet"])
            coll = ws.Rows if op["axis"] == "rows" else ws.Columns
            for index, value in op["values"]:
                coll(index).Hidden = value
        elif what in {"insert_lines", "delete_lines"}:
            rng = _range(wb, op)
            rng.Insert() if what == "insert_lines" else rng.Delete()
        elif what == "delete_created_sheet":
            with _quiet(app):
                wb.Sheets(op["name"]).Delete()
        elif what == "sheet_copy":
            index = min(op["index"], int(wb.Sheets.Count) + 1)
            before = wb.Sheets(index) if index <= int(wb.Sheets.Count) else None
            after = None if before is not None else wb.Sheets(wb.Sheets.Count)
            had_marker = _name_state(wb, MARKER) is not None
            backup.Sheets(op["backup_sheet"]).Copy(before, after)
            if not had_marker:
                # Copy листа между книгами переносит и имена уровня книги — в том числе метку служебной книги
                # (живая проверка). С меткой книга пользователя стала бы «служебной» и пропала бы из всех инструментов.
                _forget_marker(wb)
            ws = wb.Sheets(index)
            ws.Name, ws.Visible = op["name"], op["visible"]
            _repair_formulas(ws.Range(op["address"]), op["formulas"], op["formula_property"])
        elif what == "sheet_property":
            ws = wb.Sheets(op["current"])
            action = op["action"]
            if action == "rename":
                ws.Name = op["name"]
            elif action == "move":
                index = op["index"]
                if int(ws.Index) > index:
                    ws.Move(wb.Sheets(index), None)
                elif int(ws.Index) < index:
                    ws.Move(None, wb.Sheets(index))
            elif action == "tab_color":
                if op["color_index"] == -4142:
                    ws.Tab.ColorIndex = -4142
                else:
                    ws.Tab.Color = op["color"]
            elif action in {"hide", "very_hide", "unhide"}:
                ws.Visible = op["visible"]
        elif what == "name":
            if _name_state(wb, op["name"]) is not None:
                wb.Names(op["name"]).Delete()
            if op["value"] is not None:
                wb.Names.Add(op["name"], *op["value"])
        elif what == "shape":
            wb.Worksheets(op["sheet"]).Shapes(op["name"]).Delete()
        elif what == "slicer":
            wb.SlicerCaches(op["name"]).Delete()
        elif what == "table":
            wb.Worksheets(op["sheet"]).ListObjects(op["name"]).Unlist()
        elif what == "view":
            _view(app, wb, op["sheet"], op["value"])
        elif what == "calculation":
            app.Calculation = op["value"]
        elif what == "protection":
            from .excel_format import _protect, _protected

            target = wb.Worksheets(op["sheet"]) if op["scope"] == "sheet" else wb
            try:
                target.Unprotect("")
                for address, value in op["locked"].items():
                    _restore_properties(target.Range(address), {"Locked": value})
                if _protected(op["state"], op["scope"]):
                    _protect(target, op["scope"], "", op["state"])
                if op["scope"] == "sheet":
                    target.EnableSelection = op["state"]["enable_selection"]
            except pywintypes.com_error:
                raise ToolError("Excel rejected protection undo (the password or protection may have changed).") from None


EXCLUDED = {"office_journal", "office_undo", "excel_new_workbook", "word_new_document", "excel_create_from_template",
            "excel_close_workbook", "word_close_document", "word_compare_documents"}

DEFERRED = {"excel_clean_text", "excel_split_column", "excel_protection", "word_sort_table"}


class Recording:
    def __init__(self, tool, kind, arguments, supplied):
        self.tool, self.kind, self.args, self.supplied = tool, kind, redact(arguments), redact(supplied)
        self.args["_has_secret"] = any("password" in k.casefold() and bool(v) for k, v in arguments.items())
        self.targets = {}
        self.pending = {}
        self.records = []
        self.preparing = False
        self.resolved = False
        self.enabled = config.SETTINGS.undo and kind in {"write", "destructive"} and tool not in EXCLUDED

    def target(self, kind, app, obj, ready=False):
        if self.preparing:
            return
        if self.tool.startswith("bridge_"):
            into_excel = self.tool in {"bridge_word_table_to_excel", "bridge_word_text_to_excel"}
            if (kind == "workbook") != into_excel:
                return  # источник моста только читается
        key = stack_key(kind, app, obj)
        self.targets[key] = (app, obj)
        if self.tool in DEFERRED and not ready:
            return
        if not self.enabled or key in self.pending:
            return
        if self.tool == "excel_calculate" and not self.args.get("mode"):
            return
        if self.resolved and self.tool in {"excel_copy_range", "excel_manage_sheet"}:
            return
        entry = Entry(self.tool, journal.describe(self.supplied)[0], kind)
        self.pending[key] = entry
        self.preparing = True
        try:
            if kind == "document":
                record = app.UndoRecord
                if bool(record.IsRecordingCustomRecord):
                    entry.barrier("Word is already recording another custom undo record.")
                else:
                    entry.before_fingerprint = word_fingerprint(obj)
                    record.StartCustomRecord(f"Office Live: {self.tool}")
                    self.records.append(record)
            elif self.tool == "excel_copy_range":
                self._copy(app, obj, key, entry)
                self.resolved = True
            elif self.tool == "excel_manage_sheet" and self.args.get("action", "").lower() == "copy" and self.args.get("dest_workbook"):
                dest_app, dest = xl.pick_workbook(self.args["dest_workbook"])
                dest_key = stack_key(kind, dest_app, dest)
                self.targets.pop(key, None)
                self.pending.pop(key, None)
                self.targets[dest_key] = (dest_app, dest)
                self.pending[dest_key] = entry
                _add_sheet(dest_app, dest, entry, self.args)
                self.resolved = True
            else:
                resolver = RESOLVERS.get(self.tool)
                if resolver is None:
                    entry.barrier("This operation has no safe undo strategy.")
                else:
                    resolver(app, obj, entry, self.args)
            for pending_key, pending in self.pending.items():
                if pending.kind == "workbook" and pending.undoable and not pending.before_fingerprint:
                    pa, po = self.targets[pending_key]
                    pending.before_fingerprint = excel_fingerprint(pa, po, pending)
                prior = STACKS.get(pending_key, [])
                if prior and prior[-1].undoable:
                    pa, po = self.targets[pending_key]
                    try:
                        current = word_fingerprint(po) if pending.kind == "document" else excel_fingerprint(pa, po, prior[-1])
                        pending.refresh_previous = current == prior[-1].fingerprint
                    except Exception:  # noqa: BLE001
                        pending.refresh_previous = False
        except Exception as exc:  # noqa: BLE001 — невозможность снимка создаёт барьер, но не ломает инструмент
            for pending in self.pending.values():
                pending.barrier(f"Could not prepare undo: {exc}")
        finally:
            self.preparing = False

    def _copy(self, app, wb, key, entry):
        args = self.args
        ws, source = xl.get_range(wb, args["sheet"], args["source"], empty_means_used=False)
        da, dw = xl.pick_workbook(args["dest_workbook"]) if args.get("dest_workbook") else (app, wb)
        dest_key = stack_key("workbook", da, dw)
        dest_ws = xl.pick_sheet(dw, args["dest_sheet"]) if args.get("dest_sheet") else (ws if not args.get("dest_workbook") else xl.pick_sheet(dw, ""))
        r1, c1, _, _ = parse_a1(args["dest_cell"])
        rows, cols = int(source.Rows.Count), int(source.Columns.Count)
        if args.get("transpose"):
            rows, cols = cols, rows
        dest = xl.sub_range(dest_ws, r1, c1, r1 + rows - 1, c1 + cols - 1)
        if ":" in args["dest_cell"] and args.get("what", "all").lower() in {"all", "formats"}:
            explicit = dest_ws.Range(args["dest_cell"])
            if _size(explicit) > 1:  # D1:D1 — тоже одиночная ячейка, Copy расширяет её до размера источника
                dest = explicit
        if args.get("move"):
            entry.warning = "Undo restores the source and destination cells, but formulas elsewhere that pointed to the moved cells keep pointing to the destination."
        if key == dest_key:
            mode = "contents" if args.get("what", "all").lower() in {"values", "formulas"} else "formats" if args.get("what", "all").lower() == "formats" else "all"
            _snapshots(app, wb, entry, [dest, source] if args.get("move") else [dest], restore=mode,
                       properties=("NumberFormat",) if mode == "contents" else ())
        else:
            de = Entry(self.tool, f"{dest_ws.Name}!{xl.addr_of(dest)}", "workbook")
            de.warning = entry.warning
            self.targets[dest_key] = da, dw
            self.pending[dest_key] = de
            if args.get("move") and _size(source) + _size(dest) > config.SETTINGS.undo_max_cells:
                for item in (entry, de):
                    item.barrier("Combined source and destination exceeds OFFICE_LIVE_UNDO_MAX_CELLS")
                return
            mode = "contents" if args.get("what", "all").lower() in {"values", "formulas"} else "formats" if args.get("what", "all").lower() == "formats" else "all"
            _snapshots(da, dw, de, [dest], restore=mode, properties=("NumberFormat",) if mode == "contents" else ())
            if args.get("move"):
                _snapshots(app, wb, entry, [source])
            else:
                self.targets.pop(key, None)
                self.pending.pop(key, None)

    def end_records(self):
        while self.records:
            record = self.records.pop()
            try:
                record.EndCustomRecord()
            except Exception as exc:  # noqa: BLE001
                for entry in self.pending.values():
                    if entry.kind == "document":
                        entry.barrier(f"Could not end Word undo record: {exc}")

    def finish(self, result, partial=False):
        self.end_records()
        if self.tool in {"excel_close_workbook", "word_close_document"}:
            for key, (app, _) in self.targets.items():
                for entry in STACKS.pop(key, []):
                    _drop(app, entry)
                if key[0] == "workbook":
                    try:
                        _close_orphan_backups(app)
                    except Exception:  # noqa: BLE001 — уборка не должна портить успешное закрытие книги
                        pass
            return  # закрытый документ уже нельзя читать или дополнять листом «Лог»
        if self.kind == "save":
            for key, (app, obj) in self.targets.items():
                new_key = stack_key(key[0], app, obj)
                if new_key != key:
                    com.note_target(f"{key[0]}:{identity(key[0], obj)}")
                    for entry in STACKS.pop(new_key, []):
                        _drop(app, entry)
                    if key in STACKS:
                        STACKS[new_key] = STACKS.pop(key)
                    for stack in STACKS.values():
                        for entry in stack:
                            entry.peers = [new_key if peer == key else peer for peer in entry.peers]
        if len(self.pending) > 1:
            transaction = uuid.uuid4().hex
            for entry in self.pending.values():
                entry.transaction = transaction
                entry.peers = list(self.pending)
        # Идентификация созданного листа должна предшествовать созданию служебного листа «Лог».
        for key, entry in self.pending.items():
            app, obj = self.targets[key]
            if entry.kind == "workbook" and entry.undoable:
                try:
                    _finish_excel(obj, entry, result if isinstance(result, dict) else {})
                except Exception as exc:  # noqa: BLE001
                    if partial:
                        entry.force_reason = f"Could not read the partial result: {exc}; pass force=true to restore the pre-call snapshot."
                    else:
                        entry.barrier(f"Could not finish undo snapshot: {exc}")
        from .navigation import result_links

        result_links(self, result)
        if self.kind in {"write", "destructive"} and self.tool != "office_undo":
            for key, (app, obj) in self.targets.items():
                if key[0] == "workbook":
                    warning = journal.append_sheet(app, obj, self.tool, self.supplied, links=result.get("links", []) if isinstance(result, dict) else [])
                    if warning and isinstance(result, dict):
                        result["log_sheet_warning"] = warning
        statuses = []
        for key, entry in self.pending.items():
            app, obj = self.targets[key]
            if entry.undoable:
                try:
                    entry.fingerprint = word_fingerprint(obj) if entry.kind == "document" else excel_fingerprint(app, obj, entry)
                except Exception as exc:  # noqa: BLE001
                    if partial:
                        entry.force_reason = f"Could not fingerprint the partial result: {exc}; pass force=true to restore the pre-call state."
                    else:
                        entry.barrier(f"Could not fingerprint the result: {exc}")
            if not entry.undoable:
                _drop(app, entry)
            if entry.kind == "document" and entry.undoable and entry.fingerprint == entry.before_fingerprint:
                statuses.append("not available: no document change detected")
                continue  # Word не создаёт отдельного undo-действия для пустой правки
            stack = STACKS.setdefault(key, [])
            stack.append(entry)
            while len(stack) > config.SETTINGS.undo_depth:
                _drop(app, stack.pop(0))
            statuses.append("available" if entry.undoable else "not available: " + entry.reason)
            if entry.warning and isinstance(result, dict):
                result["undo_warning"] = entry.warning
            if entry.force_reason and isinstance(result, dict):
                result["undo_warning"] = " ".join(filter(None, (entry.warning, entry.force_reason)))
        if statuses and isinstance(result, dict):
            result["undo"] = next((s for s in statuses if s != "available"), "available")

    def close(self, success):
        self.end_records()
        if not success:
            for key, entry in self.pending.items():
                app, obj = self.targets[key]
                stack = STACKS.get(key, [])
                if stack:
                    try:
                        current = word_fingerprint(obj) if entry.kind == "document" else excel_fingerprint(app, obj, entry)
                    except Exception:  # noqa: BLE001
                        current = None
                    if not entry.before_fingerprint or current != entry.before_fingerprint:
                        # Нового шага нет. Но перескочить через частично выполненную правку тоже нельзя.
                        stack[-1].barrier(f"A failed {self.tool} may have changed the document or native undo history.")
                _drop(app, entry)
        self.targets.clear()  # все COM-ссылки отпускаем до CoUninitialize
        self.pending.clear()


def selected(kind, app, obj):
    """Вызывается выборщиками после AutoSave/событий, но до первой правки инструмента."""
    if _active is not None:
        _active.target(kind, app, obj)


def require_undo(kind, app, obj, *, plan=None, barrier_reason=""):
    """Новые операции начинаются только после проверок и успешной подготовки отмены."""
    if _active is None:
        raise ToolError("This operation requires the office_tool undo context.")
    if not config.SETTINGS.undo:
        return  # отмена выключена владельцем (OFFICE_LIVE_UNDO=0): работаем, как остальные инструменты
    if not _active.enabled:
        raise ToolError("This operation requires the office_tool undo context.")
    if plan:
        _active.args.update(plan)
    _active.target(kind, app, obj, ready=True)
    entry = _active.pending.get(stack_key(kind, app, obj))
    if entry is not None and barrier_reason and entry.reason == barrier_reason:
        return  # explicitly chosen password barrier; snapshot failures are still refused
    if entry is None or not entry.undoable:
        if entry is not None:
            _drop(app, entry)
            _active.pending.pop(stack_key(kind, app, obj), None)
        raise ToolError("No changes made: could not prepare office_undo. " + (entry.reason if entry else "No snapshot."))


def cancel_prepared(kind, app, obj):
    """Only for a tool that has verified complete local rollback after a failed mutation."""
    if _active is not None:
        entry = _active.pending.pop(stack_key(kind, app, obj), None)
        if entry is not None:
            _drop(app, entry)


def record_call(fn, args, kwargs, tool, kind, arguments, supplied):
    global _active
    recording = Recording(tool, kind, arguments, supplied)
    previous, _active = _active, recording
    success = False
    try:
        result = fn(*args, **kwargs)
        if not isinstance(result, dict) or result.get("ok", True):
            recording.finish(result)
            success = True
        return result
    except PartialChangeError as exc:
        result = {}
        recording.finish(result, partial=True)
        success = True
        status = result.get("undo", "not available: no change detected")
        hint = "office_undo will restore the state before this call." if status == "available" else f"office_undo: {status}."
        if result.get("undo_warning"):
            hint += " " + result["undo_warning"]
        raise ToolError(f"{exc} {hint}") from None
    finally:
        recording.close(success)
        _active = previous


def _pick_file(file):
    # Ошибка неоднозначности/защиты Excel не должна маскироваться выбором Word.
    try:
        app, obj = xl.pick_workbook(file)
        return "workbook", app, obj
    except ToolError as exc:
        text = str(exc)
        if not any(s in text for s in ("is not running", "has no open workbooks", "not found. Open workbooks")):
            raise
    app, obj = wd.pick_document(file)
    return "document", app, obj


def _available(app, entry):
    if entry.undoable and entry.backup_sheets:
        try:
            backup = _backup(app, entry.backup_token)
            for name in entry.backup_sheets:
                backup.Sheets(name)
        except Exception as exc:  # noqa: BLE001
            entry.barrier(str(exc))
    return entry.undoable


def _participants(key, app, obj, entry):
    """Один move между книгами отменяется целиком; нельзя обходить новые правки второй книги."""
    out = []
    for peer in entry.peers or [key]:
        stack = STACKS.get(peer, [])
        if peer != key:
            if not any(e.transaction == entry.transaction for e in stack):
                raise ToolError(f"{entry.tool}: related undo history for {peer[1]} is no longer available.")
            if stack[-1].transaction != entry.transaction:
                raise ToolError(f"{entry.tool}: undo the newer changes in {peer[1]} first.")
            pa, po = xl.pick_workbook(peer[1]) if peer[0] == "workbook" else wd.pick_document(peer[1])
            if stack_key(peer[0], pa, po) != peer:
                raise ToolError(f"{entry.tool}: the original Office instance for {peer[1]} is no longer available.")
        else:
            pa, po = app, obj
        out.append((peer, pa, po, stack[-1]))
    return out


@office_tool("history", "destructive", title="Undo agent changes", read_actions=("history",))
def office_undo(file: str = "", action: str = "undo", steps: int = 1, force: bool = False) -> dict:
    """Undo recent agent changes to an open workbook/document, or read the session's undo history. Stops at unsupported operations (barriers). Later user edits block undo unless force=true. In Word, force also undoes the user's later edits to reach the state before the agent change, up to 20 native steps per entry; reports native_steps and stops at a barrier if that state cannot be reached. Excel is tried first, then Word.

    Args:
        file: workbook/document name/path (empty means active; exact targets are safer).
        action: 'undo' or 'history' (history is also available in read-only mode).
        steps: number of agent changes to undo, newest first.
        force: accept that later user edits may be overwritten; in Word also undo the user's later edits (up to 20 native steps per entry).
    """
    act = action.lower()
    if act not in {"undo", "history"}:
        raise ToolError("action must be undo or history.")
    if steps < 1:
        raise ToolError("steps must be positive.")
    kind, app, obj = _pick_file(file)
    key = stack_key(kind, app, obj)
    stack = STACKS.get(key, [])
    if act == "history":
        return {"file": identity(kind, obj), "history": [
            {"tool": e.tool, "where": e.where, "time": e.stamp, "undoable": _available(app, e),
             "needs_force": bool(e.force_reason), "reason": e.reason or e.force_reason}
            for e in reversed(stack)
        ]}
    undone = []
    out = {"file": identity(kind, obj), "undone": undone}
    if kind == "document":
        out["native_steps"] = 0
    for _ in range(steps):
        if not stack:
            if not undone:
                raise ToolError("No recorded changes to undo in this server session.")
            break
        entry = stack[-1]
        try:
            participants = _participants(key, app, obj, entry)
            for pk, pa, po, pe in participants:
                if not _available(pa, pe):
                    raise ToolError(f"Undo stopped at {pe.tool}: not available: {pe.reason}")
                if pe.force_reason and not force:
                    raise ToolError(f"{pe.tool}: {pe.force_reason}")
                try:
                    current = word_fingerprint(po) if pk[0] == "document" else excel_fingerprint(pa, po, pe)
                except Exception as exc:  # noqa: BLE001
                    current = f"unavailable: {exc}"
                if current != pe.fingerprint and not force:
                    raise ToolError(f"{pe.tool}: the document changed after the agent. Undo would also overwrite/undo the user's later edits; pass force=true to proceed.")
        except ToolError as exc:
            if not undone:
                raise
            out["stopped"] = str(exc)
            break
        native_steps = 0
        try:
            for pk, pa, po, pe in participants:
                if pk[0] == "document":
                    for _ in range(20 if force else 1):
                        if force and word_fingerprint(po) == pe.before_fingerprint:
                            break
                        if not po.Undo(1):
                            raise ToolError("Word refused Undo(1); its undo history may have been cleared.")
                        native_steps += 1
                        out["native_steps"] += 1
                    if force and word_fingerprint(po) != pe.before_fingerprint:
                        raise ToolError("The limit of 20 native undo steps was reached.")
                else:
                    restore_excel(pa, po, pe)
        except Exception as exc:
            for _, _, _, pe in participants:
                if pe.kind == "document" and force:
                    pe.barrier(f"Could not reach the state before {pe.tool}; {native_steps} native steps were undone, including the user's edits. {exc}")
                else:
                    pe.barrier(f"Undo failed (possibly partially applied): {exc}")
            if not undone:
                raise ToolError(f"{entry.tool}: {entry.reason}") from None
            out["stopped"] = entry.reason
            break
        undone.append({"tool": entry.tool, "where": entry.where, **({"native_steps": native_steps} if kind == "document" else {})})
        for pk, pa, po, pe in participants:
            peer_stack = STACKS[pk]
            peer_stack.pop()
            _drop(pa, pe)
            summary = f"отменено: {pe.tool} ({pe.where})"
            from .navigation import undo_links

            linked = undo_links(pa, po, pe)
            places = linked.get("links", [])
            if places:
                existing = out.get("links", [])
                combined = list({link["uri"]: link for link in [*existing, *places]}.values())
                out.update(links=combined[:20], links_truncated=out.get("links_truncated", False) or linked.get("links_truncated", False) or len(combined) > 20)
            journal.append([f"{pk[0]}:{identity(pk[0], po)}"], "office_undo", {}, where=pe.where, summary=summary, links=places)
            if pk[0] == "workbook":
                warning = journal.append_sheet(pa, po, "office_undo", {}, where=pe.where, summary=summary, links=places)
                if warning:
                    out["log_sheet_warning"] = warning
            if pe.warning:
                out["warning"] = pe.warning
            # Ручные правки МЕЖДУ вызовами не становятся «своими» после отмены нового шага.
            if peer_stack and peer_stack[-1].undoable and pe.refresh_previous:
                try:
                    peer_stack[-1].fingerprint = word_fingerprint(po) if pk[0] == "document" else excel_fingerprint(pa, po, peer_stack[-1])
                except Exception as exc:  # noqa: BLE001
                    peer_stack[-1].barrier(f"Could not fingerprint after undo: {exc}")
    return out
