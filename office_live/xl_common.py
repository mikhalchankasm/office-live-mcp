"""Общие помощники Excel: выбор книги/листа/диапазона, чтение и запись сеток значений.

Правила COM (см. com.py): только позиционные аргументы, Resize/Offset не использовать —
границы считаем через parse_a1/a1_range.
"""

import os

import pywintypes

from . import com, config
from .errors import ToolError
from .util import delocalize_number_format, localize_number_format, EXCEL_ERRORS, MAX_COLS, MAX_ROWS, a1_cell, a1_range, quote_sheet, split_sheet_ref, to_grid

WRITE_KINDS = {"write", "destructive", "save"}
XL_WORKSHEET = -4167


# ------------------------------------------------------------------ книги


def _need_explicit(name: str, what: str):
    if not name and config.SETTINGS.strict_target and com.current_kind() in WRITE_KINDS:
        raise ToolError(
            f"{what} name is required for write operations (OFFICE_LIVE_STRICT_TARGET is on). "
            "Call the list tool first and pass the exact name."
        )


def _same_path(a: str, b: str) -> bool:
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def all_workbooks(launch: bool = False) -> list:
    """[(app, workbook)] по всем запущенным экземплярам Excel."""
    pairs = []
    for app in com.apps("excel", launch):
        for i in range(1, app.Workbooks.Count + 1):
            pairs.append((app, app.Workbooks(i)))
    return pairs


def pick_workbook(name: str = "", launch: bool = False):
    """(app, workbook): точное имя / полный путь / уникальная подстрока / '' = активная."""
    _need_explicit(name, "Workbook")
    apps = com.apps("excel", launch)
    pairs = []
    for app in apps:
        for i in range(1, app.Workbooks.Count + 1):
            pairs.append((app, app.Workbooks(i)))
    if not pairs:
        raise ToolError("Excel has no open workbooks. Use excel_open_workbook or excel_new_workbook.")
    if not name:
        for app in apps:
            wb = app.ActiveWorkbook
            if wb is not None:
                return app, wb
        return pairs[0]

    low = name.strip().lower()
    has_sep = "\\" in name or "/" in name
    infos = [(app, wb, wb.Name) for app, wb in pairs]
    exact = [(a, w) for a, w, n in infos if n.lower() == low]
    if not exact and has_sep:
        exact = [(a, w) for a, w, n in infos if _same_path(w.FullName, name)]
    if len(exact) == 1:
        return exact[0]
    if len(exact) > 1:
        paths = [w.FullName for _, w in exact]
        raise ToolError(f"Workbook name '{name}' exists in several Excel instances; pass the full path. Candidates: {paths}")
    subs = [(a, w) for a, w, n in infos if low in n.lower()]
    if len(subs) == 1:
        return subs[0]
    names = [n for _, _, n in infos]
    if subs:
        raise ToolError(f"Workbook name '{name}' is ambiguous. Open workbooks: {names}")
    raise ToolError(f"Workbook '{name}' not found. Open workbooks: {names}")


# ------------------------------------------------------------------ листы


def sheet_names(wb, any_type: bool = False) -> list[str]:
    coll = wb.Sheets if any_type else wb.Worksheets
    return [coll(i).Name for i in range(1, coll.Count + 1)]


def pick_sheet(wb, name: str = "", any_type: bool = False):
    """Лист по имени ('' = активный). any_type=True — включая листы диаграмм."""
    coll = wb.Sheets if any_type else wb.Worksheets
    if not name:
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


def get_range(wb, sheet: str, cells: str, empty_means_used: bool = True):
    """(worksheet, range). `cells` может содержать 'Лист!A1:B2'; пусто -> UsedRange."""
    sheet_override, addr = split_sheet_ref(cells or "")
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


def localize_formula(app, wb, formula: str) -> str:
    """Английская формула -> формула на языке интерфейса Excel.

    FormatConditions.Add и Validation.Add (в отличие от Range.Formula) принимают формулы на ЛОКАЛЬНОМ языке
    (в русском Excel: =НЕ(ЕОШИБКА(ПОИСК("a";A1))), разделитель ';'). Переводим через временный лист.
    """
    if app.International[4] == "," and app.International[2] == ".":
        return formula  # англоязычные региональные настройки — перевод не нужен
    prev = wb.ActiveSheet
    was_saved = bool(wb.Saved)
    alerts = app.DisplayAlerts
    try:
        tmp = wb.Worksheets.Add(None, wb.Sheets(wb.Sheets.Count))
    except pywintypes.com_error:
        return formula  # структура книги защищена
    try:
        tmp.Range("A1").Formula = formula
        return tmp.Range("A1").FormulaLocal
    finally:
        try:
            app.DisplayAlerts = False
            tmp.Delete()
        finally:
            app.DisplayAlerts = alerts
            if prev is not None:
                try:
                    prev.Activate()
                except pywintypes.com_error:
                    pass
            wb.Saved = was_saved


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


def delocalize_formulas(app, wb, formulas: list[str]) -> list[str]:
    """Формулы на языке интерфейса -> английские (пакетом, через один временный лист). При ошибке формула остаётся как есть."""
    if not formulas or (app.International[4] == "," and app.International[2] == "."):
        return list(formulas)
    prev = wb.ActiveSheet
    was_saved = bool(wb.Saved)
    alerts = app.DisplayAlerts
    try:
        tmp = wb.Worksheets.Add(None, wb.Sheets(wb.Sheets.Count))
    except pywintypes.com_error:
        return list(formulas)
    out = []
    try:
        for i, f in enumerate(formulas, start=1):
            try:
                tmp.Cells(i, 1).FormulaLocal = f if f.startswith("=") else "=" + f
                out.append(tmp.Cells(i, 1).Formula)
            except pywintypes.com_error:
                out.append(f)
    finally:
        try:
            app.DisplayAlerts = False
            tmp.Delete()
        finally:
            app.DisplayAlerts = alerts
            if prev is not None:
                try:
                    prev.Activate()
                except pywintypes.com_error:
                    pass
            wb.Saved = was_saved
    return out
