"""Excel undo snapshots, resolvers and restoration."""

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
from . import config, journal, wd_common as wd, xl_common as xl
from .errors import ToolError
from .util import clean_word_text, to_com_grid, to_grid


MARKER = "__OfficeLiveUndo"
SESSION = uuid.uuid4().hex

def _hash(value):
    return hashlib.sha1(json.dumps(value, ensure_ascii=False, default=str).encode("utf-8")).hexdigest()

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
        for rule in _rules(rng):
            state.append([_optional_state(rule, ("Priority",)), *_rule_state(rule)])
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


def _rules(rng):
    rules = rng.FormatConditions
    return [rules(i) for i in range(1, int(rules.Count) + 1)]


def _rule_state(rule):
    """Признаки правила условного форматирования без приоритета: номера приоритетов Excel перенумеровывает сам."""
    state = [int(rule.Type), xl.addr_of(rule.AppliesTo),
             _optional_state(rule, ("StopIfTrue", "Formula1", "Formula2", "Operator", "Text", "TextOperator", "DateOperator",
                                    "Rank", "Percent", "TopBottom", "AboveBelow", "NumStdDev", "DupeUnique"))]
    for member, props in (("Font", ("Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Color")),
                          ("Interior", ("Color", "Pattern", "PatternColor"))):
        try:
            state.append(_optional_state(getattr(rule, member), props))
        except (pywintypes.com_error, AttributeError):
            state.append(None)
    return state


def _inside(rng, box):
    r1, c1, r2, c2 = box
    return all(r1 <= a[0] and c1 <= a[1] and a[2] <= r2 and a[3] <= c2 for a in map(xl.bounds, rng.Areas))


def _rule_closure(ws, box, *, whole_sheet=False):
    """Расширяет прямоугольник, пока каждое пересекающее его правило не попадёт в него целиком (whole_sheet — все правила).

    Вставка форматов из снимка, обрезающего правило, разрезает его на части (живая проверка): снимок должен содержать
    правила целиком, тогда вставка вернёт их без изменений.
    """
    rules = [[xl.bounds(area) for area in rule.AppliesTo.Areas] for rule in _rules(ws.Cells)]
    if whole_sheet:
        for areas in rules:
            box = (min(box[0], *(a[0] for a in areas)), min(box[1], *(a[1] for a in areas)),
                   max(box[2], *(a[2] for a in areas)), max(box[3], *(a[3] for a in areas)))
    changed = True
    while changed:
        changed = False
        for areas in rules:
            r1, c1, r2, c2 = box
            if any(a[0] <= r2 and r1 <= a[2] and a[1] <= c2 and c1 <= a[3] for a in areas) and not all(
                    r1 <= a[0] and c1 <= a[1] and a[2] <= r2 and a[3] <= c2 for a in areas):
                box = (min(r1, *(a[0] for a in areas)), min(c1, *(a[1] for a in areas)),
                       max(r2, *(a[2] for a in areas)), max(c2, *(a[3] for a in areas)))
                changed = True
    return xl.sub_range(ws, *box)


def _restore_rules(target, before):
    """Правила диапазона как до вызова: лишние (добавленные агентом) удаляем, итог сверяем.

    Вставка форматов из другой книги не удаляет правила, добавленные поверх пустого снимка (живая проверка), поэтому
    отмена excel_conditional_format не может полагаться на PasteSpecial. Не удалось вернуть точно — ошибка, а не «отменено».
    """
    box = xl.bounds(target)
    remaining, extra = list(before), []
    for rule in _rules(target):
        state = _rule_state(rule)
        if state in remaining:
            remaining.remove(state)
        else:
            extra.append(rule)
    if not remaining and all(_inside(rule.AppliesTo, box) for rule in extra):
        for rule in reversed(extra):
            rule.Delete()
    after = [_rule_state(rule) for rule in _rules(target)]
    if sorted(map(repr, after)) != sorted(map(repr, before)):
        raise ToolError(f"Conditional formatting in {target.Worksheet.Name}!{xl.addr_of(target)} could not be restored exactly "
                        f"({len(after)} rule(s) instead of {len(before)}); check them with excel_conditional_format action='list'.")


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
    before_light: str = ""  # Word: без styles.xml — для досбора шагов, на которые Word разбил одну правку агента
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


def _snapshots(app, wb, entry, ranges, *, columns=False, rows=False, restore="all", properties=(), rules=False):
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
                if rules or restore == "rules":
                    data["rules"] = [_rule_state(rule) for rule in _rules(rng)]
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
    from .excel_cf import _remember_view, _restore_view

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
        if entry.tool in {"excel_import_table", "excel_clean_text", "excel_split_column", "excel_goal_seek", "excel_what_if"}:
            state.append([_property_state(rng, p) for p in ("NumberFormat", "PrefixCharacter")])
        for mode in modes & {"all", "formats", "comments", "validation", "rules"}:
            state.append(_metadata_state(rng, "formats" if mode == "rules" else mode))
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
        elif what == "scenario":
            from .excel_analysis import scenario_items, scenario_state

            state.append([scenario_state(s) for s in scenario_items(wb.Worksheets(op["sheet"]))
                          if s.Name.casefold() == op["name"].casefold()])
        elif what == "sheet_property":
            ws = wb.Sheets(op["current"])
            state.append([ws.Name, int(ws.Index), int(ws.Visible), ws.Tab.ColorIndex, ws.Tab.Color])
        elif what == "calculation":
            state.append(app.Calculation)
        elif what == "import_view":
            state.append(_import_view(app, wb, op["sheet"]))
        elif what == "view":
            state.append(_view(app, wb, op["sheet"]))
        elif what == "protection":
            from .excel_format import _protection_state

            target = wb.Worksheets(op["sheet"]) if op["scope"] == "sheet" else wb
            state.append(_protection_state(target, op["scope"]))
            state.extend(_property_state(target.Range(a), "Locked") for a in op["locked"])
        elif what == "subtotal":
            from .excel_sheets import outline_levels

            state.append(outline_levels(wb.Worksheets(op["sheet"]), wb.Worksheets(op["sheet"]).Range(op["current_address"])))
        elif what == "sparklines":
            from .excel_format import spark_groups, spark_state

            ws = wb.Worksheets(op["sheet"])
            location = ws.Range(op["address"])
            state.append([spark_state(g, full=op["full"]) for g in spark_groups(ws) if app.Intersect(location, g.Location) is not None])
    return _hash(state)


def _simple_range(app, wb, entry, args):
    _, rng = xl.get_range(wb, args.get("sheet", ""), args.get("cells", args.get("cell", "")))
    if entry.tool == "excel_clear_range" and args.get("what", "contents").lower() == "contents":
        _snapshots(app, wb, entry, [rng], restore="contents")
    elif entry.tool == "excel_manage_comments" or (entry.tool == "excel_clear_range" and args.get("what", "").lower() == "comments"):
        _snapshots(app, wb, entry, [rng], restore="comments")
    elif entry.tool == "excel_data_validation":
        _snapshots(app, wb, entry, [rng], restore="validation")
    elif entry.tool == "excel_conditional_format":
        if args.get("action", "add").lower() == "add":
            _snapshots(app, wb, entry, [rng], restore="rules")  # добавление меняет только правила: снимаем лишние
        else:
            # clear удаляет и обрезает правила (без cells — на всём листе): снимок охватывает их целиком
            ws = rng.Worksheet
            whole = not args.get("cells")
            _snapshots(app, wb, entry, [_rule_closure(ws, xl.bounds(ws.UsedRange if whole else rng), whole_sheet=whole)], restore="formats", rules=True)
    elif entry.tool == "excel_clear_range" and args.get("what", "").lower() == "formats":
        _snapshots(app, wb, entry, [rng], restore="formats")
    elif entry.tool in {"excel_format_range", "excel_format_ranges"} and args.get("style") is None:
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
    elif entry.tool in {"excel_format_range", "excel_format_ranges"}:
        _snapshots(app, wb, entry, [rng], restore="formats")
    else:
        _snapshots(app, wb, entry, [rng])


IMPORT_BOOK_BARRIER = "Import created a new workbook; office_undo cannot close it. Close it explicitly with discard=true if unwanted."


def _union_range(app, ws, address):
    """Range() refuses addresses over 255 characters (a union of up to 32 far cells): build it area by area."""
    if len(str(address)) <= 255:
        return ws.Range(address)
    result = None
    for part in str(address).split(","):
        area = ws.Range(part)
        result = area if result is None else app.Union(result, area)
    return result


def _import_view(app, wb, sheet, value=None):
    # Never activate a sheet just to inspect/fingerprint an import.
    if app.ActiveWorkbook.FullName != wb.FullName or app.ActiveSheet.Name != sheet:
        raise ToolError("Activate the imported sheet before undoing its frozen view.")
    win = app.ActiveWindow
    props = ("SplitRow", "SplitColumn", "FreezePanes", "ScrollRow", "ScrollColumn")
    if value is None:
        return {p: getattr(win, p) for p in props}
    win.FreezePanes = False
    for prop in props:
        setattr(win, prop, value[prop])


def _import_table(app, wb, entry, args):
    if not args.get("workbook"):
        entry.barrier(IMPORT_BOOK_BARRIER)
    elif args.get("_import_created"):
        _add_sheet(app, wb, entry, args)
    else:
        ws, rng = xl.get_range(wb, args["sheet"], args["cells"])
        if args.get("layout", True) and args.get("header", True) and app.ActiveWorkbook.FullName == wb.FullName and app.ActiveSheet.Name == ws.Name:
            entry.ops.append({"op": "import_view", "sheet": ws.Name, "before": _import_view(app, wb, ws.Name)})
        _snapshots(app, wb, entry, [rng], columns=args.get("layout", True), properties=("NumberFormat",))
        if args.get("layout", True) and args.get("header", True):
            # Wrapped header text grows its row; restored after the range so pasted formats cannot re-grow it.
            r1, c1, _, c2 = xl.bounds(rng)
            header = xl.sub_range(ws, r1, c1, r1, c2)
            entry.ops.append({"op": "dimensions", **_area(header), **_dimensions(header, False, True, explicit=True)})


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


def _what_if(app, wb, entry, args):
    ws = wb.Worksheets(args["_what_sheet"])
    action = args["action"]
    if action == "data_table":
        if _size(ws.Range(args["_what_results"])) > config.SETTINGS.undo_max_cells or journal.owns_sheet(ws):
            entry.barrier("Data table exceeds the undo cell limit or targets the service log sheet.")
            return
        area = {"sheet": ws.Name, "address": args["_what_results"]}
        entry.areas.append(area)
        entry.ops.append({"op": "clear_data_table", **area})  # preflight proves every result is empty
    elif action == "show_scenario":
        _snapshots(app, wb, entry, [ws.Range(a) for a in args["_what_cells"]], restore="contents", properties=("NumberFormat",))
    elif action == "scenario_summary":
        _add_sheet(app, wb, entry, args)
    else:
        entry.ops.append({"op": "scenario", "sheet": ws.Name, "name": args["_what_name"], "before": args["_what_before"]})


def _goal_seek(app, wb, entry, args):
    rng = wb.Worksheets(args["_goal_sheet"]).Range(args["_goal_cell"])
    _snapshots(app, wb, entry, [rng], restore="contents", properties=("NumberFormat",))


def _subtotals(app, wb, entry, args):
    if args.get("_operation_barrier"):
        entry.barrier(args["_operation_barrier"])
        return
    rng = wb.Worksheets(args["_subtotal_sheet"]).Range(args["_subtotal_range"])
    entry.ops.append({"op": "subtotal", **_area(rng), "current_address": xl.addr_of(rng), "rows": int(rng.Rows.Count)})
    _snapshots(app, wb, entry, [rng], restore="contents", properties=("NumberFormat",))


def _sparklines(app, wb, entry, args):
    if args.get("_operation_barrier"):
        entry.barrier(args["_operation_barrier"])
        return
    entry.ops.append({"op": "sparklines", "sheet": args["_spark_sheet"], "address": args["_spark_location"],
                      "before": args["_spark_before"], "full": args.get("action") == "clear"})


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
    targets = args.get("_dimension_targets")
    if targets is None:
        from .excel_sheets import _dimension_targets

        sizes, fit, mode = _dimension_targets(wb, args["sheet"], args)
        targets = [{"cells": xl.addr_of(rng), "property": prop} for rng, prop, _ in sizes]
        if mode in {"columns", "both"}:
            targets.append({"cells": xl.addr_of(fit), "property": "ColumnWidth"})
        if mode in {"rows", "both"}:
            targets.append({"cells": xl.addr_of(fit), "property": "RowHeight"})
    ranges = [(xl.get_range(wb, args["sheet"], t["cells"])[1], t["property"]) for t in targets]
    count = sum(int(rng.Columns.Count if prop == "ColumnWidth" else rng.Rows.Count) for rng, prop in ranges)
    if count > config.SETTINGS.undo_max_cells:
        entry.barrier(f"{count} dimensions exceeds OFFICE_LIVE_UNDO_MAX_CELLS")
        return
    for rng, prop in ranges:
        entry.ops.append({"op": "dimensions", **_area(rng),
                          **_dimensions(rng, prop == "ColumnWidth", prop == "RowHeight", explicit=True)})


def _format_ranges(app, wb, entry, args):
    specs = args["_format_ranges"]
    ranges = [xl.get_range(wb, args["sheet"], spec["cells"], empty_means_used=False)[1] for spec in specs]
    if sum(_size(rng) for rng in ranges) > config.SETTINGS.undo_max_cells:
        entry.barrier("Combined formatting ranges exceed OFFICE_LIVE_UNDO_MAX_CELLS")
        return
    # Every snapshot precedes the first edit, including overlapping rectangles.
    for spec in specs:
        _simple_range(app, wb, entry, {"sheet": args["sheet"], "cells": spec["cells"], **spec["format"]})


def _hidden(app, wb, entry, args):
    from .excel_sheets import _line_spec

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
    from .excel_sheets import _rows_cols_range

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
    "excel_what_if": _what_if,
    "excel_import_table": _import_table,
    "excel_goal_seek": _goal_seek, "excel_subtotals": _subtotals, "excel_sparklines": _sparklines,
    "excel_format_ranges": _format_ranges,
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
            if entry.tool not in {"excel_goal_seek", "excel_subtotals", "excel_what_if"}:
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
            op["sheet"] = result.get("sheet") or op["sheet"]  # сводная диаграмма создаётся на листе сводной, а не на `sheet`
        elif what == "slicer":
            op["name"] = result["cache"]
        elif what == "table":
            op["name"] = result["table"]


def restore_excel(app, wb, entry):
    xl.suspend_events(app, force=True)
    backup = _backup(app, entry.backup_token) if entry.backup_sheets else None
    for op in entry.ops:
        what = op["op"]
        if what == "clear_data_table":
            _range(wb, op).ClearContents()  # clear the entire TABLE array in one call
        elif what == "scenario":
            from .excel_analysis import scenario_items

            ws = wb.Worksheets(op["sheet"])
            for item in scenario_items(ws):
                if item.Name.casefold() == op["name"].casefold():
                    item.Delete()
            before = op["before"]
            if before is not None:
                ws.Scenarios().Add(before["name"], _union_range(app, ws, before["changing_cells"]), before["values"],
                                   before["comment"], before["locked"], before["hidden"])
        elif what == "import_view":
            _import_view(app, wb, op["sheet"], op["before"])
        elif what == "subtotal":
            ws = wb.Worksheets(op["sheet"])
            current = ws.Range(op["current_address"])
            current_rows = int(current.Rows.Count)
            start = int(current.Row)
            below = ws.Cells(start + current_rows, int(current.Column))
            current.RemoveSubtotal()
            if int(below.Row) - start != op["rows"]:
                raise ToolError("Subtotal undo did not restore the original row count; snapshot was not written.")
        elif what == "sparklines":
            from .excel_format import spark_groups, spark_state

            ws = wb.Worksheets(op["sheet"])
            location = ws.Range(op["address"])
            for group in spark_groups(ws):
                if app.Intersect(location, group.Location) is not None:
                    if int(app.Intersect(location, group.Location).Count) != int(group.Location.Count):
                        raise ToolError("Sparkline undo refuses a group extending outside the recorded location.")
                    group.Delete()
            for before in op["before"]:
                group = ws.Range(before["location"]).SparklineGroups.Add(before["type"], before["source"])
                for path, value in before["properties"].items():
                    if path == "DateRange" and not value:
                        continue  # a newly added group already has no date axis; Excel may reject an empty DateRange setter
                    if path == "PlotBy" and not value:
                        continue  # 0 = non-square data: Excel chooses the orientation and rejects the setter
                    obj, member = _property_object(group, path)
                    setattr(obj, member, value)
                if spark_state(group, full=True) != before:
                    raise ToolError("Sparkline restore could not be verified.")
        elif what == "range":
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
            if "rules" in op:
                _restore_rules(target, op["rules"])
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
