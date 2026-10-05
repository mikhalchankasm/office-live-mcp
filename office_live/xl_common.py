"""Общие помощники Excel: выбор книги/листа/диапазона, чтение и запись сеток значений.

Правила COM (см. com.py): только позиционные аргументы, Resize/Offset не использовать —
границы считаем через parse_a1/a1_range.
"""

import contextlib
import os
import re

import pythoncom
import pywintypes

from . import com, config, safety
from .errors import ToolError
from .util import _split_literals, delocalize_number_format, localize_number_format, EXCEL_ERRORS, MAX_COLS, MAX_ROWS, a1_cell, a1_range, quote_sheet, split_sheet_ref, to_grid

WRITE_KINDS = {"write", "destructive", "save"}
XL_WORKSHEET = -4167


# ------------------------------------------------------------------ книги


def _strict_write() -> bool:
    """Строгий режим включён и вызывается пишущий инструмент: цель обязана быть названа точно."""
    return bool(config.SETTINGS.strict_target and com.current_kind() in WRITE_KINDS)


def _need_explicit(name: str, what: str):
    if _strict_write() and not (name or "").strip():  # пробельное имя — тоже «пусто»
        raise ToolError(
            f"{what} name is required for write operations (OFFICE_LIVE_STRICT_TARGET is on). "
            "Call the list tool first and pass the exact name."
        )


def _same_path(a: str, b: str) -> bool:
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def workbook_path(wb) -> str:
    try:
        return wb.FullName if wb.Path else ""
    except pywintypes.com_error:
        return ""


def workbook_allowed(wb) -> bool:
    """Книга в зоне доступа (OFFICE_LIVE_ALLOWED_DIRS)? Несохранённая книга без пути допустима."""
    return not is_undo_workbook(wb) and safety.doc_allowed(workbook_path(wb))


def is_undo_workbook(wb) -> bool:
    """Служебная книга отмены никогда не является пользовательской целью."""
    try:
        wb.Names("__OfficeLiveUndo")
        return True
    except (pywintypes.com_error, AttributeError, KeyError):
        return False


def suspend_events(app, force: bool = False) -> None:
    """Отключает события Excel (Worksheet_Change, BeforeSave, SheetActivate ...) до конца вызова; возврат — в run_com.

    Если отключить не удалось, вызов отклоняется: чужие обработчики могли бы изменить документ или сработать на наши правки.
    OFFICE_LIVE_ENABLE_EVENTS=1 оставляет события включёнными осознанно.
    """
    if config.SETTINGS.enable_events and not force:
        return
    try:
        if bool(app.EnableEvents):
            app.EnableEvents = False
            com.add_cleanup(lambda a=app: setattr(a, "EnableEvents", True))
    except pywintypes.com_error as exc:
        raise ToolError(
            "Could not switch off Excel events for this call (" + com.com_error_text(exc) + "), so it was not performed: macros of the open workbooks could react to it. "
            "Retry, or start the server with OFFICE_LIVE_ENABLE_EVENTS=1 to allow events."
        ) from None


def _guard_workbook(app, wb, allow_autosave: bool = False):
    """Проверки перед пишущим вызовом: AutoSave (правки сохраняются сами и не проходят проверку пользователем) и события Excel."""
    com.note_target(f"workbook:{workbook_path(wb) or wb.Name}")
    if com.current_kind() not in WRITE_KINDS:
        from .undo import selected

        selected("workbook", app, wb)
        return app, wb
    if not allow_autosave and config.SETTINGS.autosave == "block":
        try:
            auto = bool(wb.AutoSaveOn)
        except (pywintypes.com_error, AttributeError):
            auto = False
        if auto:
            raise ToolError(
                f"AutoSave is ON for {wb.Name}: every change would be saved to the cloud file immediately, before the user can review it. "
                "Ask the user to turn AutoSave off for this workbook (or start the server with OFFICE_LIVE_AUTOSAVE=allow)."
            )
    suspend_events(app)
    from .undo import selected

    selected("workbook", app, wb)
    return app, wb


def all_workbooks(launch: bool = False) -> list:
    """[(app, workbook)] по всем запущенным экземплярам Excel (только книги из зоны доступа)."""
    pairs = []
    for app in com.apps("excel", launch):
        for i in range(1, app.Workbooks.Count + 1):
            wb = app.Workbooks(i)
            if workbook_allowed(wb):
                pairs.append((app, wb))
    return pairs


def pick_workbook(name: str = "", launch: bool = False, allow_autosave: bool = False, *, exact_only: bool = False):
    """(app, workbook): точное имя / полный путь / уникальная подстрока / '' = активная.

    В строгом режиме для пишущих инструментов — только точное имя или полный путь.
    """
    _need_explicit(name, "Workbook")
    strict = _strict_write() or exact_only
    name = (name or "").strip()
    apps = com.apps("excel", launch)
    pairs, outside = [], 0
    for app in apps:
        for i in range(1, app.Workbooks.Count + 1):
            wb = app.Workbooks(i)
            if workbook_allowed(wb):
                pairs.append((app, wb))
            else:
                outside += 1
    if not pairs:
        note = f" ({outside} open workbook(s) are outside OFFICE_LIVE_ALLOWED_DIRS)" if outside else ""
        raise ToolError("Excel has no open workbooks" + note + ". Use excel_open_workbook or excel_new_workbook.")
    if not name:
        for app in apps:
            wb = app.ActiveWorkbook
            if wb is not None and not is_undo_workbook(wb):
                if not workbook_allowed(wb):
                    raise ToolError("The active workbook is outside OFFICE_LIVE_ALLOWED_DIRS; pass the name of an allowed workbook.")
                return _guard_workbook(app, wb, allow_autosave)
        return _guard_workbook(*pairs[0], allow_autosave)

    low = name.lower()
    has_sep = "\\" in name or "/" in name
    infos = [(app, wb, wb.Name) for app, wb in pairs]
    exact = [(a, w) for a, w, n in infos if n.lower() == low]
    if not exact and has_sep:
        exact = [(a, w) for a, w, n in infos if _same_path(w.FullName, name)]
    if len(exact) == 1:
        return _guard_workbook(*exact[0], allow_autosave)
    if len(exact) > 1:
        paths = [w.FullName for _, w in exact]
        raise ToolError(f"Workbook name '{name}' exists in several Excel instances; pass the full path. Candidates: {paths}")
    names = [n for _, _, n in infos]
    if not strict:
        subs = [(a, w) for a, w, n in infos if low in n.lower()]
        if len(subs) == 1:
            return _guard_workbook(*subs[0], allow_autosave)
        if subs:
            raise ToolError(f"Workbook name '{name}' is ambiguous. Open workbooks: {names}")
    raise ToolError(f"Workbook '{name}' not found. Open workbooks: {names}")


# ------------------------------------------------------------------ листы


def sheet_names(wb, any_type: bool = False) -> list[str]:
    coll = wb.Sheets if any_type else wb.Worksheets
    return [coll(i).Name for i in range(1, coll.Count + 1)]


def pick_sheet(wb, name: str = "", any_type: bool = False, *, exact_only: bool = False):
    """Лист по имени ('' = активный). any_type=True — включая листы диаграмм."""
    coll = wb.Sheets if any_type else wb.Worksheets
    if _strict_write() and not (name or "").strip():
        raise ToolError("Sheet name is required for write operations (OFFICE_LIVE_STRICT_TARGET is on). Pass the exact sheet name.")
    if not (name or "").strip():
        sh = wb.ActiveSheet
        if sh is None:
            raise ToolError(f"Workbook {wb.Name} has no active sheet.")
        return sh
    n = coll.Count
    names = [coll(i).Name for i in range(1, n + 1)]
    low = name.strip().lower()
    for i, nm in enumerate(names, start=1):
        if nm.lower() == low:
            return coll(i)
    if not _strict_write() and not exact_only:
        subs = [i for i, nm in enumerate(names, start=1) if low in nm.lower()]
        if len(subs) == 1:
            return coll(subs[0])
    raise ToolError(f"Sheet '{name}' not found in {wb.Name}. Sheets: {names}")


def validate_sheet_name(name: str):
    if not name or not name.strip():
        raise ToolError("Sheet name is empty.")
    if len(name) > 31:
        raise ToolError(f"Sheet name is longer than 31 characters: {name!r}")
    bad = [c for c in '[]:*?/\\' if c in name]
    if bad:
        raise ToolError(f"Sheet name contains forbidden characters {bad}: {name!r}")
    if name.startswith("'") or name.endswith("'"):
        raise ToolError("Sheet name cannot start or end with an apostrophe.")


# ------------------------------------------------------------------ диапазоны


def addr_of(rng) -> str:
    return rng.Address.replace("$", "")


def bounds(rng) -> tuple[int, int, int, int]:
    """(r1, c1, r2, c2) первой области диапазона."""
    r1, c1 = int(rng.Row), int(rng.Column)
    return r1, c1, r1 + int(rng.Rows.Count) - 1, c1 + int(rng.Columns.Count) - 1


def sub_range(ws, r1: int, c1: int, r2: int, c2: int):
    return ws.Range(a1_range(r1, c1, r2, c2))


def get_range(wb, sheet: str, cells: str, empty_means_used: bool = True, allow_other_sheet: bool = False):
    """(worksheet, range). `cells` может содержать 'Лист!A1:B2'; пусто -> UsedRange.

    Если лист указан дважды и по-разному ('sheet' и префикс в `cells`) — это ошибка, а не молчаливый выбор одного из них
    (allow_other_sheet=True — для источников данных, которые законно лежат на другом листе).
    """
    sheet_override, addr = split_sheet_ref(cells or "")
    if sheet_override and (sheet or "").strip() and not allow_other_sheet and sheet.strip().lower() != sheet_override.strip().lower():
        raise ToolError(
            f"Conflicting sheets: sheet='{sheet}' but cells refers to '{sheet_override}'. Pass one of them (the sheet prefix in `cells` or the `sheet` argument)."
        )
    ws = pick_sheet(wb, sheet_override or sheet)
    if not addr:
        if not empty_means_used:
            raise ToolError("'cells' is required.")
        return ws, ws.UsedRange
    try:
        return ws, ws.Range(addr)
    except pywintypes.com_error:
        raise ToolError(
            f"Invalid range '{cells}'. Use A1 notation like 'B2', 'B2:D10', 'A:C', '3:7', 'Sheet2!A1:B3' or a defined name."
        ) from None


def clip_to_used(app, ws, rng):
    """Для ссылок на целые строки/столбцы — пересечение с UsedRange (иначе читали бы миллион строк)."""
    if int(rng.Rows.Count) >= MAX_ROWS or int(rng.Columns.Count) >= MAX_COLS:
        return app.Intersect(rng, ws.UsedRange)  # None, если пересечения нет
    return rng


def bounded_range(app, wb, sheet, cells, limit):
    """Обязательный прямоугольник с лимитом до чтения данных."""
    ws, rng = get_range(wb, sheet, cells, empty_means_used=False)
    if int(rng.Areas.Count) != 1:
        raise ToolError("Multi-area ranges are not supported; select one rectangle.")
    ws = rng.Worksheet
    if ws.Parent.FullName != wb.FullName:
        raise ToolError("The range must belong to the selected workbook.")
    rng = clip_to_used(app, ws, rng)
    if rng is not None and int(rng.Rows.Count) * int(rng.Columns.Count) > limit:
        raise ToolError(f"Range exceeds {limit} cells; narrow the range.")
    return ws, rng


def count_nonempty(app, rng) -> int:
    try:
        return int(app.WorksheetFunction.CountA(rng))
    except pywintypes.com_error:
        return -1


def sheet_is_empty(app, ws) -> bool:
    return count_nonempty(app, ws.UsedRange) == 0


def ref_label(ws, rng) -> str:
    return f"{quote_sheet(ws.Name)}!{addr_of(rng)}"


# ------------------------------------------------------------------ чтение/запись сеток


def read_grid(rng, what: str = "values") -> list[list]:
    """what: values | formulas | text (отображаемый текст, медленно — по ячейкам)."""
    if what == "formulas":
        f = rng.Formula
        return to_grid(f)
    if what == "text":
        r1, c1, r2, c2 = bounds(rng)
        ws = rng.Worksheet
        return [[ws.Cells(r, c).Text for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
    return to_grid(rng.Value)


def formula_cells(rng, limit: int = 500) -> list[dict]:
    """Разреженный список формул диапазона: [{'cell': 'B3', 'formula': '=SUM(A1:A2)'}]."""
    r1, c1, r2, c2 = bounds(rng)
    grid = to_grid(rng.Formula)
    out = []
    for i, row in enumerate(grid):
        for j, f in enumerate(row):
            if isinstance(f, str) and f.startswith("="):
                out.append({"cell": a1_cell(r1 + i, c1 + j), "formula": f})
                if len(out) >= limit:
                    return out
    return out


_ERROR_TEXTS = frozenset(EXCEL_ERRORS.values())


def error_cells(grid: list[list], r1: int, c1: int, limit: int = 20) -> list[dict]:
    out = []
    for i, row in enumerate(grid):
        for j, v in enumerate(row):
            if isinstance(v, str) and v in _ERROR_TEXTS:
                out.append({"cell": a1_cell(r1 + i, c1 + j), "error": v})
                if len(out) >= limit:
                    return out
    return out


def preview(rng, max_cells: int = 200) -> list[list]:
    """Контрольное чтение для ответов пишущих инструментов (ограничено по размеру)."""
    r1, c1, r2, c2 = bounds(rng)
    rows, cols = r2 - r1 + 1, c2 - c1 + 1
    if rows * cols <= max_cells:
        return to_grid(rng.Value)
    keep_rows = max(1, max_cells // max(cols, 1))
    ws = rng.Worksheet
    return to_grid(ws.Range(a1_range(r1, c1, min(r2, r1 + keep_rows - 1), c2)).Value)


@contextlib.contextmanager
def scratch_workbook(app):
    """Скрытая временная книга для служебных операций (перевод формул, снимок диапазона).

    Книги пользователя не трогаем: ни временных листов и диаграмм, ни подмены флага Saved, ни событий. Активной
    остаётся прежняя книга; временная закрывается без сохранения даже при ошибке.
    """
    prev = app.ActiveWorkbook
    alerts = app.DisplayAlerts
    events = app.EnableEvents
    app.EnableEvents = False
    try:
        tmp = app.Workbooks.Add()
    except BaseException:
        app.EnableEvents = events
        raise
    try:
        try:
            tmp.Windows(1).Visible = False
        except pywintypes.com_error:
            pass
        if prev is not None:
            try:
                prev.Activate()
            except pywintypes.com_error:
                pass
        yield tmp
    finally:
        try:
            app.DisplayAlerts = False
            tmp.Close(False)
        except pywintypes.com_error:
            pass
        finally:
            app.DisplayAlerts = alerts
            app.EnableEvents = events
            if prev is not None:
                try:
                    prev.Activate()
                except pywintypes.com_error:
                    pass


_STRING_LITERAL = re.compile(r'"(?:[^"]|"")*"')
_STRUCTURED_REF = re.compile(r"[\w.\u0400-\u04FF]+\[(?:[^\[\]]|\[[^\]]*\])*\]")
_SHEET_PREFIX = re.compile(r"(?:'((?:[^']|'')+)'|([\w.\u0400-\u04FF]+(?::[\w.\u0400-\u04FF]+)?))!")


def _mask_formula(formula: str):
    """Формула -> (безопасная для временной книги, функция обратной подстановки, имена листов-ссылок).

    Строковые литералы и структурные ссылки Таблица[Столбец] заменяются метками (перевод затрагивает только имена
    функций и разделители); ссылки на другие книги не поддерживаются — Excel открыл бы диалог выбора файла.
    """
    saved: dict[str, str] = {}

    def keep(match, kind):
        token = f"__OL{kind}{len(saved)}__"
        saved[token] = match.group(0)
        return token

    text = _STRING_LITERAL.sub(lambda m: '"' + keep(m, "S") + '"', formula)
    text = _STRUCTURED_REF.sub(lambda m: keep(m, "T"), text)
    if "[" in text:
        raise ToolError("Formulas that refer to other workbooks ('[Book.xlsx]Sheet'!A1) are not supported here; copy the data into this workbook first.")
    sheets: list[str] = []
    for m in _SHEET_PREFIX.finditer(text):
        raw = m.group(1).replace("''", "'") if m.group(1) is not None else m.group(2)
        sheets.extend(part for part in raw.split(":") if part)

    def restore(result: str) -> str:
        for token, original in saved.items():
            result = result.replace('"' + token + '"', original) if token.startswith("__OLS") else result.replace(token, original)
        return result

    return text, restore, sheets


def _translate_formulas(app, formulas: list[str], to_local: bool, strict: bool) -> list[str]:
    """Пакетный перевод формул английский <-> язык интерфейса через одну временную книгу."""
    out = []
    with scratch_workbook(app) as tmp:
        ws = tmp.Worksheets(1)
        have = {tmp.Sheets(i).Name.lower() for i in range(1, tmp.Sheets.Count + 1)}
        for i, formula in enumerate(formulas, start=1):
            try:
                masked, restore, sheets = _mask_formula(formula)
                for name in sheets:  # формула со ссылкой на несуществующий лист открыла бы диалог «Обновление значений»
                    if name.lower() not in have and len(name) <= 31 and not any(ch in name for ch in "[]:*?/\\"):
                        sh = tmp.Worksheets.Add(None, tmp.Sheets(tmp.Sheets.Count))
                        sh.Name = name
                        have.add(name.lower())
                cell = ws.Cells(i, 16000)  # далеко от A1: формулы вида =A1>5 не должны ссылаться на свою же ячейку
                if to_local:
                    cell.Formula = masked
                    out.append(restore(cell.FormulaLocal))
                else:
                    cell.FormulaLocal = masked if masked.startswith("=") else "=" + masked
                    out.append(restore(cell.Formula))
            except (pywintypes.com_error, ToolError):
                if strict:
                    raise
                out.append(formula)  # при чтении лучше показать как есть, чем потерять правило
    return out


def _is_english_locale(app) -> bool:
    return app.International[4] == "," and app.International[2] == "."


def localize_formula(app, wb, formula: str) -> str:
    """Английская формула -> формула на языке интерфейса Excel.

    FormatConditions.Add и Validation.Add (в отличие от Range.Formula) принимают формулы на ЛОКАЛЬНОМ языке
    (в русском Excel: =НЕ(ЕОШИБКА(ПОИСК("a";A1))), разделитель ';'). Переводим во временной скрытой книге.
    """
    if _is_english_locale(app):
        return formula
    return _translate_formulas(app, [formula], True, strict=True)[0]


def number_format_for_write(app, fmt: str, shortcuts: dict | None = None) -> str:
    """Инвариантный код формата ('#,##0.00', '0.0%') или ярлык ('percent') -> запись региональных настроек.

    Через COM-вызовы pywin32 Excel трактует NumberFormat в ЛОКАЛЬНОЙ записи (в русском: десятичная ',', тысячи ' ';
    '0.00' отвергается), поэтому агенту показываем и принимаем от него инвариантную запись, а переводим здесь.
    """
    inv = (shortcuts or {}).get(fmt.strip().lower(), fmt)
    intl = app.International
    return localize_number_format(inv, intl[2], intl[3])


def number_format_for_read(app, fmt):
    if not isinstance(fmt, str):
        return fmt
    intl = app.International
    return delocalize_number_format(fmt, intl[2], intl[3])


LCID_EN_US = 1033


def _local_codes(fmt: str) -> bool:
    """Код уже в локальной записи ('ДД.ММ.ГГГГ', 'Основной'): буквы не из ASCII вне кавычек, [скобок] и экранирования."""
    return any(ch.isalpha() and not ch.isascii() for text, literal in _split_literals(fmt) if not literal for ch in text)


def set_number_format(app, obj, fmt: str, shortcuts: dict | None = None) -> None:
    """Записать NumberFormat (Range, поле сводной, подписи осей) в инвариантной записи.

    pywin32 вызывает свойства с LCID пользователя, и русский Excel понимает только локальный код: 'dd.mm.yyyy' становится
    текстом, а 'General', '[Red]', '[h]' отвергаются. Запись с LCID 1033 (так пишет VBA) принимает английские коды целиком
    (живая проверка). Код, уже записанный по-местному, и объекты без COM-интерфейса (заглушки тестов) идут прежним путём.
    """
    inv = (shortcuts or {}).get(fmt.strip().lower(), fmt)
    ole = getattr(obj, "_oleobj_", None)
    if ole is not None and not _local_codes(inv):
        try:
            ole.Invoke(ole.GetIDsOfNames("NumberFormat"), LCID_EN_US, pythoncom.DISPATCH_PROPERTYPUT, 0, inv)
            return
        except pywintypes.com_error:
            pass  # пробуем локальную запись ниже; её ошибка уйдёт вызывающему
    obj.NumberFormat = number_format_for_write(app, inv)


def read_number_format(app, obj):
    """NumberFormat в инвариантной записи (None для смешанного формата): обратная пара к set_number_format."""
    ole = getattr(obj, "_oleobj_", None)
    if ole is not None:
        try:
            return ole.Invoke(ole.GetIDsOfNames("NumberFormat"), LCID_EN_US, pythoncom.DISPATCH_PROPERTYGET, 1)
        except pywintypes.com_error:
            pass
    return number_format_for_read(app, obj.NumberFormat)


def delocalize_formulas(app, wb, formulas: list[str]) -> list[str]:
    """Формулы на языке интерфейса -> английские (пакетом, через одну временную книгу). Нераспознанная формула остаётся как есть."""
    if not formulas or _is_english_locale(app):
        return list(formulas)
    return _translate_formulas(app, formulas, False, strict=False)
