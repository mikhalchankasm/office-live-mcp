"""Excel: управление сводными таблицами — фильтры, срезы и временные шкалы, поля, сортировка, группировка, параметры.

Создание/список/обновление/удаление сводных — в excel_analysis.py. Здесь всё, что делает с готовой сводной
то же, что пользователь делает мышью: скрывает строки (элементы), фильтрует, ставит срезы, группирует.
"""

import datetime

import pythoncom
import pywintypes

from . import com
from .errors import ToolError
from .registry import office_tool
from .util import parse_a1, to_grid
from .xl_common import addr_of, get_range, pick_sheet, pick_workbook, preview

MISSING = pythoncom.Missing

ORIENTATIONS = {"hidden": 0, "row": 1, "column": 2, "filter": 3, "page": 3, "data": 4}
ORIENT_NAMES = {0: "hidden", 1: "row", 2: "column", 3: "filter", 4: "data"}
PIVOT_FUNCS = {
    "sum": -4157, "count": -4112, "average": -4106, "avg": -4106, "max": -4136, "min": -4139,
    "product": -4149, "count_numbers": -4113, "stdev": -4155, "var": -4164,
}
SHOW_AS = {"normal": -4143, "percent_of_total": 8, "percent_of_row": 6, "percent_of_column": 7, "percent_of_parent": 12, "difference": 2, "running_total": 5}
# XlPivotFilterType: подписи (xlCaption*) и значения (xlValue*) — разные диапазоны кодов
LABEL_FILTERS = {
    "equals": 15, "not_equals": 16, "begins_with": 17, "not_begins_with": 18, "ends_with": 19, "not_ends_with": 20,
    "contains": 21, "not_contains": 22, "greater": 23, "greater_equal": 24, "less": 25, "less_equal": 26,
    "between": 27, "not_between": 28,
}
VALUE_FILTERS = {"equals": 7, "not_equals": 8, "greater": 9, "greater_equal": 10, "less": 11, "less_equal": 12, "between": 13, "not_between": 14}
TOP_FILTERS = {"top_items": 1, "bottom_items": 2, "top_percent": 3, "bottom_percent": 4, "top_sum": 5, "bottom_sum": 6}
DATE_PERIODS = ["seconds", "minutes", "hours", "days", "months", "quarters", "years"]


# ------------------------------------------------------------------ поиск объектов


def _all_pivots(wb) -> list:
    out = []
    for i in range(1, int(wb.Worksheets.Count) + 1):
        ws = wb.Worksheets(i)
        pts = ws.PivotTables()
        for j in range(1, int(pts.Count) + 1):
            out.append((ws, pts(j)))
    return out


def find_pivot(wb, name: str = ""):
    pivots = _all_pivots(wb)
    if not pivots:
        raise ToolError(f"{wb.Name} has no pivot tables. Create one with excel_create_pivot_table.")
    if not name:
        if len(pivots) == 1:
            return pivots[0]
        raise ToolError(f"Several pivot tables exist; pass 'pivot'. Available: {[pt.Name for _, pt in pivots]}")
    for ws, pt in pivots:
        if pt.Name.lower() == name.lower():
            return ws, pt
    raise ToolError(f"Pivot table '{name}' not found. Available: {[pt.Name for _, pt in pivots]}")


def _field_names(pt) -> list[str]:
    return [f.Name for f in pt.PivotFields()]


def _field(pt, name: str):
    try:
        return pt.PivotFields(name)
    except pywintypes.com_error:
        raise ToolError(f"Field '{name}' not found in pivot '{pt.Name}'. Fields: {_field_names(pt)}") from None


def _data_field(pt, caption: str):
    try:
        return pt.DataFields(caption)
    except pywintypes.com_error:
        raise ToolError(f"Data field '{caption}' not found. Data fields (use the caption): {[f.Name for f in pt.DataFields]}") from None


class _Manual:
    """Откладывает пересчёт сводной на время массовых изменений (иначе каждый шаг перерисовывает её)."""

    def __init__(self, pt):
        self.pt = pt

    def __enter__(self):
        try:
            self.pt.ManualUpdate = True
        except pywintypes.com_error:
            pass
        return self

    def __exit__(self, *exc):
        try:
            self.pt.ManualUpdate = False
        except pywintypes.com_error:
            pass
        return False


def _items(pf, limit: int = 300) -> list:
    out = []
    try:
        coll = pf.PivotItems()
        for k in range(1, min(int(coll.Count), limit) + 1):
            out.append(coll(k))
    except pywintypes.com_error:
        pass
    return out


def _item_visible(pi) -> bool:
    try:
        return bool(pi.Visible)
    except pywintypes.com_error:
        return True


def _summary(ws, pt) -> dict:
    area = pt.TableRange1
    return {"sheet": ws.Name, "pivot": pt.Name, "range": addr_of(area), "preview": preview(area, 120)}


# ================================================================== информация


@office_tool("excel_analysis", "read", title="Pivot table details")
def excel_pivot_info(workbook: str, pivot: str = "", max_items: int = 100) -> dict:
    """Describe a pivot table in detail: every field with its role (row/column/filter/data/unused), the items of each field and whether they are visible (hidden = filtered out), data fields with their function, and applied filters. Use it before filtering or restructuring.

    Args:
        workbook: exact workbook name.
        pivot: pivot table name ('' if the workbook has only one).
        max_items: cap on items listed per field.
    """
    app, wb = pick_workbook(workbook)
    ws, pt = find_pivot(wb, pivot)
    fields = []
    for pf in pt.PivotFields():
        try:
            orient = int(pf.Orientation)
        except pywintypes.com_error:
            continue
        info = {"name": pf.Name, "role": ORIENT_NAMES.get(orient, str(orient))}
        if orient in (1, 2, 3):
            items = _items(pf, int(max_items))
            info["items"] = [{"name": pi.Name, "visible": _item_visible(pi)} for pi in items]
            info["hidden_items"] = sum(1 for pi in items if not _item_visible(pi))
            try:
                info["filters"] = int(pf.PivotFilters.Count)
            except pywintypes.com_error:
                pass
        fields.append(info)
    data = []
    for df in pt.DataFields:
        entry = {"caption": df.Name, "source_field": df.SourceName, "function": int(df.Function)}
        try:
            entry["number_format"] = df.NumberFormat
        except pywintypes.com_error:
            pass
        data.append(entry)
    return {
        "workbook": wb.Name, "sheet": ws.Name, "pivot": pt.Name, "range": addr_of(pt.TableRange1),
        "fields": fields, "data_fields": data,
        "grand_totals": {"rows": bool(pt.RowGrand), "columns": bool(pt.ColumnGrand)},
    }


# ================================================================== фильтры


@office_tool("excel_analysis", "write", title="Filter pivot table", read_actions=("items",))
def excel_pivot_filter(
    workbook: str,
    field: str = "",
    action: str = "items",
    pivot: str = "",
    items: list[str] | None = None,
    mode: str = "only",
    operator: str = "equals",
    value1: str | float | int | None = None,
    value2: str | float | int | None = None,
    data_field: str = "",
) -> dict:
    """Filter a pivot table - hide/show rows (items) of a field, apply label/value/top-N filters, or clear filters. This is what the field's filter dropdown does in Excel.

    Args:
        workbook: exact workbook name.
        field: the pivot field to filter (a row, column or filter field). May be empty only for action='clear' (clears all fields).
        action: 'items' (choose which items are visible), 'label' (text filter on the item names), 'value' (filter by an aggregated number), 'top' (top/bottom N), 'clear'.
        pivot: pivot table name ('' if only one).
        items: for action='items' - the item names.
        mode: for action='items': 'only' = show only these items, 'hide' = hide these items, 'show' = make these visible too, 'all' = show every item.
        operator: label: equals|not_equals|begins_with|not_begins_with|ends_with|not_ends_with|contains|not_contains|greater|greater_equal|less|less_equal|between|not_between. value: equals|not_equals|greater|greater_equal|less|less_equal|between|not_between. top: top_items|bottom_items|top_percent|bottom_percent|top_sum|bottom_sum.
        value1, value2: operands (value2 for between; for top, value1 is N).
        data_field: caption of the data field (as shown in the pivot, e.g. 'Sum of Revenue') for action='value' and 'top'.
    """
    app, wb = pick_workbook(workbook)
    ws, pt = find_pivot(wb, pivot)
    act = action.lower()
    if act == "clear":
        targets = [_field(pt, field)] if field else [f for f in pt.PivotFields() if _safe_orient(f) in (1, 2, 3)]
        with _Manual(pt):
            for pf in targets:
                try:
                    pf.ClearAllFilters()
                except pywintypes.com_error:
                    pass
                for pi in _items(pf):
                    if not _item_visible(pi):
                        pi.Visible = True
        return {"ok": True, "workbook": wb.Name, "cleared": field or "all fields", **_summary(ws, pt)}
    if not field:
        raise ToolError("'field' is required.")
    pf = _field(pt, field)
    if _safe_orient(pf) not in (1, 2, 3):
        raise ToolError(f"Field '{field}' is not on the row, column or filter axis; add it first with excel_pivot_fields(action='add').")

    if act == "items":
        m = mode.lower()
        if m not in ("only", "hide", "show", "all"):
            raise ToolError("mode must be 'only', 'hide', 'show' or 'all'.")
        pitems = _items(pf, 100000)
        names = [pi.Name for pi in pitems]
        index = {n.lower(): pi for n, pi in zip(names, pitems)}
        wanted = []
        if m != "all":
            if not items:
                raise ToolError("'items' is required for this mode.")
            for it in items:
                pi = index.get(str(it).lower())
                if pi is None:
                    raise ToolError(f"Item '{it}' not found in field '{field}'. Items: {names[:60]}{'...' if len(names) > 60 else ''}")
                wanted.append(pi)
        orient = _safe_orient(pf)
        if orient == 3 and m == "only" and len(wanted) == 1:
            pf.EnableMultiplePageItems = False
            pf.CurrentPage = wanted[0].Name
        else:
            if orient == 3:
                pf.EnableMultiplePageItems = True
            with _Manual(pt):
                if m == "all":
                    for pi in pitems:
                        if not _item_visible(pi):
                            pi.Visible = True
                elif m == "show":
                    for pi in wanted:
                        pi.Visible = True
                elif m == "hide":
                    hide_ids = {id(raw_of(pi)) for pi in wanted}
                    if all((not _item_visible(pi)) or id(raw_of(pi)) in hide_ids for pi in pitems):
                        raise ToolError("At least one item must stay visible.")
                    for pi in wanted:
                        pi.Visible = False
                else:  # only: сначала показать нужные, потом спрятать остальные
                    keep = {n.lower() for n in (str(i) for i in items)}
                    for pi in wanted:
                        pi.Visible = True
                    for n, pi in zip(names, pitems):
                        if n.lower() not in keep:
                            pi.Visible = False
        visible = _displayed_items(pf)
        return {"ok": True, "workbook": wb.Name, "field": field, "visible_items": visible[:200], "visible_count": len(visible), "total_items": len(names), **_summary(ws, pt)}

    if act in ("label", "value", "top"):
        op = operator.lower()
        table = {"label": LABEL_FILTERS, "value": VALUE_FILTERS, "top": TOP_FILTERS}[act]
        if op not in table:
            raise ToolError(f"For action='{act}' operator must be one of {sorted(table)}")
        if value1 is None:
            raise ToolError("value1 is required.")
        df = None
        if act in ("value", "top"):
            if not data_field:
                raise ToolError(f"'data_field' (the caption of an aggregated field, e.g. 'Sum of Revenue') is required for action='{act}'.")
            df = _data_field(pt, data_field)
        if act != "top" and op in ("between", "not_between") and value2 is None:
            raise ToolError("value2 is required for between/not_between.")
        pf.ClearAllFilters()
        if act == "label":
            operands = [str(value1)] + ([str(value2)] if value2 is not None else [])
        else:  # значения и N — числа
            operands = [_number(value1, "value1")] + ([_number(value2, "value2")] if value2 is not None else [])
        try:
            # PivotFilters.Add2(Type, DataField, Value1, Value2): для текстовых фильтров DataField = None (Missing здесь не принимается)
            pf.PivotFilters.Add2(table[op], df, *operands)
        except pywintypes.com_error as exc:
            raise ToolError("Excel rejected the filter (check the operator against the field type and the operands): " + com.com_error_text(exc)) from None
        visible = _displayed_items(pf)
        return {"ok": True, "workbook": wb.Name, "field": field, "filter": f"{act}:{op}", "visible_items": visible[:200], "visible_count": len(visible), **_summary(ws, pt)}
    raise ToolError("action must be 'items', 'label', 'value', 'top' or 'clear'.")


def _number(v, what: str) -> float:
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        raise ToolError(f"{what} must be a number for this filter, got {v!r}.") from None


def _displayed_items(pf) -> list[str]:
    """Элементы, реально показанные в сводной (учитывает и ручной выбор, и фильтры по подписи/значению/топ-N)."""
    try:
        seen: list[str] = []
        for row in to_grid(pf.DataRange.Value):
            for v in row:
                if v not in (None, "") and str(v) not in seen:
                    seen.append(str(v))
        return seen
    except pywintypes.com_error:
        return [pi.Name for pi in _items(pf, 100000) if _item_visible(pi)]


def _safe_orient(pf) -> int:
    try:
        return int(pf.Orientation)
    except pywintypes.com_error:
        return 0


def raw_of(proxy):
    return com.raw(proxy)


# ================================================================== поля, сортировка, группировка


def _refetch_data_field(pt, source: str, position: int):
    """Поле данных после смены функции: Excel переименовывает его, и старая ссылка недействительна.

    Ищем по позиции среди полей данных (источник у нескольких полей может совпадать — «Сумма» и «Среднее» по одному полю);
    если позиция не подошла, берём единственное поле с тем же источником, иначе честно отказываем.
    """
    try:
        cand = pt.DataFields(int(position))
        if cand.SourceName == source:
            return cand
    except pywintypes.com_error:
        pass
    try:
        same = [pt.DataFields(k) for k in range(1, int(pt.DataFields.Count) + 1) if pt.DataFields(k).SourceName == source]
    except pywintypes.com_error:
        same = []
    if len(same) == 1:
        return same[0]
    raise ToolError(
        f"Excel changed the aggregation, but the edited value field of '{source}' could not be identified afterwards "
        "(several value fields share this source). Read the pivot and adjust the caption/format by the new caption."
    )


def pivot_number_format(app, fmt: str) -> str:
    """Формат поля данных сводной: ярлыки и инвариантные коды -> локальная запись."""
    from .excel_format import NUMBER_FORMATS
    from .xl_common import number_format_for_write

    return number_format_for_write(app, fmt, NUMBER_FORMATS)


def _set_field_format(pt, df, fmt: str):
    app = pt.Application
    try:
        df.NumberFormat = pivot_number_format(app, fmt)
    except pywintypes.com_error:
        raise ToolError(f"Excel rejected number_format '{fmt}' for a pivot value. Use a plain code such as '0', '0.00', '#,##0', '#,##0.00', '0.0%', 'dd.mm.yyyy'.") from None


def _add_data_field(pt, spec: dict):
    fname = spec.get("field")
    func = str(spec.get("function", "sum")).lower()
    if func not in PIVOT_FUNCS:
        raise ToolError(f"Unknown function '{func}'. Use {sorted(PIVOT_FUNCS)}")
    pf = _field(pt, fname)
    caption = spec.get("caption") or f"{func.capitalize()} of {fname}"
    df = pt.AddDataField(pf, caption, PIVOT_FUNCS[func])
    if spec.get("number_format"):
        _set_field_format(pt, df, spec["number_format"])
    if spec.get("show_as"):
        sa = str(spec["show_as"]).lower()
        if sa not in SHOW_AS:
            raise ToolError(f"show_as must be one of {sorted(SHOW_AS)}")
        df.Calculation = SHOW_AS[sa]
    return df


@office_tool("excel_analysis", "write", title="Edit pivot fields", destructive=True)
def excel_pivot_fields(
    workbook: str,
    action: str,
    pivot: str = "",
    field: str = "",
    orientation: str = "row",
    position: int = 0,
    function: str = "",
    caption: str = "",
    number_format: str = "",
    show_as: str = "",
    new_caption: str = "",
    order: str = "asc",
    sort_by: str = "",
    group_by: list[str] | None = None,
    group_start: float | None = None,
    group_end: float | None = None,
    group_step: float | None = None,
    formula: str = "",
    items: list[str] | None = None,
) -> dict:
    """Restructure a pivot table: add/move/remove fields, change how a value is aggregated, sort, group dates or numbers, add calculated fields, expand/collapse.

    Args:
        workbook: exact workbook name.
        action: 'add' (put `field` on `orientation`; for data supply function/caption/number_format/show_as) | 'move' (change orientation/position of a field) | 'remove' (take a field off the pivot; for a data field pass its caption as `field`) | 'data_field' (change an existing data field `field`=its caption: function, new_caption, number_format, show_as) | 'sort' (field sorted by labels, or by a data field with sort_by) | 'group' (date field: group_by; number field: group_start/group_end/group_step) | 'ungroup' | 'calculated_field' (name in `field`, formula like '=Revenue/Qty') | 'expand' | 'collapse' | 'rename' (field -> new_caption).
        pivot: pivot table name ('' if only one).
        field: pivot field name (or data field caption where stated).
        orientation: row | column | filter | data | hidden. position: 1-based place among the fields of that axis (0 = last).
        function: sum|count|average|max|min|product|count_numbers|stdev|var (empty = sum when adding a value field; for 'data_field' empty leaves the aggregation unchanged). caption: caption of the new data field. number_format, show_as (normal|percent_of_total|percent_of_row|percent_of_column|percent_of_parent|difference|running_total).
        new_caption: new caption for 'data_field'/'rename'.
        order: 'asc' | 'desc' | 'manual' for sort. sort_by: data field caption to sort by values (e.g. 'Sum of Revenue').
        group_by: for dates a list of: days, months, quarters, years, hours, minutes, seconds. group_start/end/step: numeric grouping bounds and bin size.
        formula: calculated field formula using field names, e.g. '=Revenue/Qty'.
        items: for expand/collapse - only these items (default: the whole field).
    """
    app, wb = pick_workbook(workbook)
    ws, pt = find_pivot(wb, pivot)
    act = action.lower()

    if act == "add":
        if not field:
            raise ToolError("'field' is required.")
        o = orientation.lower()
        if o not in ORIENTATIONS or o == "hidden":
            raise ToolError("orientation must be row, column, filter or data.")
        with _Manual(pt):
            if o == "data":
                _add_data_field(pt, {"field": field, "function": function or "sum", "caption": caption, "number_format": number_format, "show_as": show_as})
            else:
                pf = _field(pt, field)
                pf.Orientation = ORIENTATIONS[o]
                if position:
                    pf.Position = int(position)
        return {"ok": True, "workbook": wb.Name, "action": act, "field": field, "orientation": o, **_summary(ws, pt)}

    if act == "move":
        o = orientation.lower()
        if o not in ORIENTATIONS or o in ("hidden", "data"):
            raise ToolError("orientation must be row, column or filter.")
        pf = _field(pt, field)
        with _Manual(pt):
            pf.Orientation = ORIENTATIONS[o]
            if position:
                pf.Position = int(position)
        return {"ok": True, "workbook": wb.Name, "action": act, "field": field, "orientation": o, **_summary(ws, pt)}

    if act == "remove":
        try:
            df = pt.DataFields(field)
            df.Orientation = 0
        except pywintypes.com_error:
            _field(pt, field).Orientation = 0
        return {"ok": True, "workbook": wb.Name, "action": act, "field": field, **_summary(ws, pt)}

    if act == "data_field":
        df = _data_field(pt, field)
        changed = []
        if function:
            f = function.lower()
            if f not in PIVOT_FUNCS:
                raise ToolError(f"Unknown function '{function}'. Use {sorted(PIVOT_FUNCS)}")
            source = df.SourceName
            try:
                position = int(df.Position)
            except pywintypes.com_error:
                position = 0
            df.Function = PIVOT_FUNCS[f]
            # Excel сам переименовывает поле («Среднее по полю X») и после этого старая ссылка недействительна — берём заново
            df = _refetch_data_field(pt, source, position)
            cap = new_caption or f"{f.capitalize()} of {source}"
            df.Caption = cap
            changed.append(f"function={f}")
            changed.append(f"caption={cap}")
            new_caption = ""
        if number_format:
            _set_field_format(pt, df, number_format)
            changed.append("number_format")
        if show_as:
            sa = show_as.lower()
            if sa not in SHOW_AS:
                raise ToolError(f"show_as must be one of {sorted(SHOW_AS)}")
            df.Calculation = SHOW_AS[sa]
            changed.append(f"show_as={sa}")
        if new_caption:
            df.Caption = new_caption
            changed.append(f"caption={new_caption}")
        return {"ok": True, "workbook": wb.Name, "action": act, "changed": changed, **_summary(ws, pt)}

    if act == "rename":
        if not new_caption:
            raise ToolError("'new_caption' is required.")
        _field(pt, field).Caption = new_caption
        return {"ok": True, "workbook": wb.Name, "action": act, "field": field, "new_caption": new_caption, **_summary(ws, pt)}

    if act == "sort":
        pf = _field(pt, field)
        o = order.lower()
        code = {"asc": 1, "desc": 2, "manual": -4135}.get(o)
        if code is None:
            raise ToolError("order must be 'asc', 'desc' or 'manual'.")
        if sort_by:
            _data_field(pt, sort_by)
        # AutoSort(Order, Field): Field — подпись поля данных; для сортировки по меткам — имя самого поля
        pf.AutoSort(code, sort_by or field)
        return {"ok": True, "workbook": wb.Name, "action": act, "field": field, "order": o, "by": sort_by or "labels", **_summary(ws, pt)}

    if act in ("group", "ungroup"):
        pf = _field(pt, field)
        if _safe_orient(pf) not in (1, 2):
            raise ToolError("Only a field on the row or column axis can be grouped.")
        first = pf.DataRange.Cells(1)
        if act == "ungroup":
            first.Ungroup()
            return {"ok": True, "workbook": wb.Name, "action": act, "field": field, **_summary(ws, pt)}
        if group_by:
            wanted = {g.lower() for g in group_by}
            unknown = wanted - set(DATE_PERIODS)
            if unknown:
                raise ToolError(f"group_by values must come from {DATE_PERIODS}; got {sorted(unknown)}")
            periods = tuple(p in wanted for p in DATE_PERIODS)
            first.Group(True, True, MISSING, periods)  # Group(Start, End, By, Periods)
        elif group_step is not None:
            first.Group(
                group_start if group_start is not None else True,
                group_end if group_end is not None else True,
                float(group_step),
            )
        else:
            raise ToolError("Pass group_by (dates) or group_step (numbers).")
        return {"ok": True, "workbook": wb.Name, "action": act, "field": field, **_summary(ws, pt)}

    if act == "calculated_field":
        if not field or not formula:
            raise ToolError("Pass the new field name in 'field' and the 'formula'.")
        f = formula if formula.startswith("=") else "=" + formula
        pt.CalculatedFields().Add(field, f, True)  # Add(Name, Formula, UseStandardFormula)
        if orientation.lower() == "data":
            _add_data_field(pt, {"field": field, "function": function or "sum", "caption": caption or field, "number_format": number_format})
        return {"ok": True, "workbook": wb.Name, "action": act, "field": field, "formula": f, **_summary(ws, pt)}

    if act in ("expand", "collapse"):
        pf = _field(pt, field)
        flag = act == "expand"
        if items:
            for it in items:
                try:
                    pf.PivotItems(str(it)).ShowDetail = flag
                except pywintypes.com_error:
                    raise ToolError(f"Item '{it}' not found in '{field}'.") from None
        else:
            pf.ShowDetail = flag
        return {"ok": True, "workbook": wb.Name, "action": act, "field": field, **_summary(ws, pt)}

    raise ToolError("action must be one of: add, move, remove, data_field, rename, sort, group, ungroup, calculated_field, expand, collapse")


# ================================================================== параметры сводной


@office_tool("excel_analysis", "write", title="Pivot table options")
def excel_pivot_options(
    workbook: str,
    pivot: str = "",
    layout: str | None = None,
    repeat_labels: bool | None = None,
    show_row_totals: bool | None = None,
    show_column_totals: bool | None = None,
    subtotals: bool | None = None,
    style: str | None = None,
    empty_cell_text: str | None = None,
    error_text: str | None = None,
    keep_column_widths: bool | None = None,
    refresh_on_open: bool | None = None,
    new_source: str = "",
    source_sheet: str = "",
    new_name: str = "",
) -> dict:
    """Layout and behaviour options of a pivot table.

    Args:
        workbook: exact workbook name.
        pivot: pivot table name ('' if only one).
        layout: compact | tabular | outline. repeat_labels: repeat item labels in every row.
        show_row_totals: totals column at the right. show_column_totals: totals row at the bottom.
        subtotals: show/hide subtotals of the row/column fields.
        style: pivot style name such as 'PivotStyleMedium9', 'PivotStyleLight16'.
        empty_cell_text: text for empty cells (e.g. '0' or '-'); '' restores the default. error_text: text shown instead of errors.
        keep_column_widths: false stops Excel from resizing columns on refresh.
        refresh_on_open: refresh the data when the file is opened.
        new_source: change the source range (e.g. 'A1:F900'; with source_sheet) or an Excel table name.
        new_name: rename the pivot table.
    """
    app, wb = pick_workbook(workbook)
    ws, pt = find_pivot(wb, pivot)
    applied = []
    if layout is not None:
        codes = {"compact": 0, "tabular": 1, "outline": 2}
        if layout.lower() not in codes:
            raise ToolError("layout must be compact, tabular or outline.")
        pt.RowAxisLayout(codes[layout.lower()])
        applied.append(f"layout={layout}")
    if repeat_labels is not None:
        pt.RepeatAllLabels(2 if repeat_labels else 1)  # xlRepeatLabels / xlDoNotRepeatLabels
        applied.append(f"repeat_labels={repeat_labels}")
    if show_row_totals is not None:
        pt.RowGrand = bool(show_row_totals)
        applied.append(f"show_row_totals={show_row_totals}")
    if show_column_totals is not None:
        pt.ColumnGrand = bool(show_column_totals)
        applied.append(f"show_column_totals={show_column_totals}")
    if subtotals is not None:
        for pf in pt.PivotFields():
            if _safe_orient(pf) in (1, 2):
                try:
                    pf.Subtotals = (bool(subtotals),) + (False,) * 11  # первый элемент — «Автоматически»
                except pywintypes.com_error:
                    pass
        applied.append(f"subtotals={subtotals}")
    if style is not None:
        pt.TableStyle2 = style
        applied.append(f"style={style}")
    if empty_cell_text is not None:
        pt.DisplayNullString = bool(empty_cell_text)
        pt.NullString = empty_cell_text
        applied.append("empty_cell_text")
    if error_text is not None:
        pt.DisplayErrorString = bool(error_text)
        pt.ErrorString = error_text
        applied.append("error_text")
    if keep_column_widths is not None:
        pt.HasAutoFormat = not keep_column_widths
        applied.append(f"keep_column_widths={keep_column_widths}")
    if refresh_on_open is not None:
        pt.PivotCache().RefreshOnFileOpen = bool(refresh_on_open)
        applied.append(f"refresh_on_open={refresh_on_open}")
    if new_source:
        src = None
        for sh_i in range(1, int(wb.Worksheets.Count) + 1):
            coll = wb.Worksheets(sh_i).ListObjects
            for j in range(1, int(coll.Count) + 1):
                if coll(j).Name.lower() == new_source.strip().lower():
                    src = coll(j).Range
        if src is None:
            _, src = get_range(wb, source_sheet, new_source, empty_means_used=False)
        pt.ChangePivotCache(wb.PivotCaches().Create(1, src))
        pt.RefreshTable()
        applied.append("new_source")
    if new_name:
        pt.Name = new_name
        applied.append(f"name={new_name}")
    if not applied:
        raise ToolError("Nothing to change: pass at least one option.")
    return {"ok": True, "workbook": wb.Name, "applied": applied, **_summary(ws, pt)}


# ================================================================== срезы и временные шкалы


def _cache_pivots(sc) -> list[str]:
    out = []
    try:
        coll = sc.PivotTables
        for k in range(1, int(coll.Count) + 1):
            out.append(coll(k).Name)
    except pywintypes.com_error:
        pass
    return out


def _slicer_info(sc, max_items: int = 100) -> dict:
    info = {"cache": sc.Name, "field": sc.SourceName, "pivot_tables": _cache_pivots(sc), "slicers": []}
    try:
        sls = sc.Slicers
        for k in range(1, int(sls.Count) + 1):
            sl = sls(k)
            shape = sl.Shape
            info["slicers"].append({
                "name": sl.Name, "caption": sl.Caption, "left": float(shape.Left), "top": float(shape.Top),
                "width": float(shape.Width), "height": float(shape.Height),
            })
    except pywintypes.com_error:
        pass
    try:
        items = sc.SlicerItems
        names = []
        for k in range(1, min(int(items.Count), max_items) + 1):
            it = items(k)
            names.append({"name": it.Name, "selected": bool(it.Selected)})
        info["items"] = names
    except pywintypes.com_error:
        info["timeline"] = True
    return info


def _find_cache(wb, key: str):
    """Кэш среза по имени кэша, имени среза или полю."""
    seen = []
    for k in range(1, int(wb.SlicerCaches.Count) + 1):
        sc = wb.SlicerCaches(k)
        seen.append(sc.Name)
        if sc.Name.lower() == key.lower() or sc.SourceName.lower() == key.lower():
            return sc
        try:
            for j in range(1, int(sc.Slicers.Count) + 1):
                if sc.Slicers(j).Name.lower() == key.lower() or sc.Slicers(j).Caption.lower() == key.lower():
                    return sc
        except pywintypes.com_error:
            pass
    raise ToolError(f"Slicer '{key}' not found. Slicer caches: {seen}")


@office_tool("excel_analysis", "write", title="Slicers and timelines", read_actions=("list",), destructive=True)
def excel_manage_slicers(
    workbook: str,
    action: str = "list",
    slicer: str = "",
    pivot: str = "",
    table: str = "",
    field: str = "",
    sheet: str = "",
    anchor_cell: str = "",
    caption: str = "",
    width: float = 150,
    height: float = 180,
    columns: int = 1,
    style: str = "",
    items: list[str] | None = None,
    connect_pivots: list[str] | None = None,
    kind: str = "slicer",
    date_from: str = "",
    date_to: str = "",
    timeline_level: str = "",
) -> dict:
    """Create and drive slicers (clickable filter buttons, 'срезы') and date timelines for pivot tables and Excel tables.

    Args:
        workbook: exact workbook name.
        action: 'list' | 'add' (needs field and pivot or table) | 'select' (choose items/date range) | 'clear' (reset the filter) | 'connect' (link more pivots to the slicer) | 'move' (reposition/resize/restyle) | 'delete'.
        slicer: slicer or cache name or the field it filters (for select/clear/connect/move/delete).
        pivot: pivot table the slicer filters (for add; '' if the workbook has only one pivot and no table is given).
        table: Excel table name to slice instead of a pivot.
        field: field/column to slice by (for add).
        sheet: sheet to place the slicer on (default: the pivot's sheet).
        anchor_cell: top-left cell for the slicer (default: right of the pivot).
        caption: slicer header text. width, height: size in points. columns: number of button columns.
        style: slicer style, e.g. 'SlicerStyleLight1', 'SlicerStyleDark3', 'SlicerStyleOther1'.
        items: for select - the item names to keep selected (all others get deselected).
        connect_pivots: pivot table names to connect (they must share the same source data); for add also connected immediately.
        kind: 'slicer' (default) or 'timeline' (date field).
        date_from, date_to: for select on a timeline, 'YYYY-MM-DD'.
        timeline_level: for a new timeline - years | quarters | months | days.
    """
    app, wb = pick_workbook(workbook)
    act = action.lower()
    if act == "list":
        return {"workbook": wb.Name, "slicers": [_slicer_info(wb.SlicerCaches(k)) for k in range(1, int(wb.SlicerCaches.Count) + 1)]}

    if act == "add":
        if not field:
            raise ToolError("'field' is required.")
        if table:
            src_ws, src = None, None
            for i in range(1, int(wb.Worksheets.Count) + 1):
                coll = wb.Worksheets(i).ListObjects
                for j in range(1, int(coll.Count) + 1):
                    if coll(j).Name.lower() == table.lower():
                        src_ws, src = wb.Worksheets(i), coll(j)
            if src is None:
                raise ToolError(f"Table '{table}' not found.")
            area = src.Range
        else:
            src_ws, src = find_pivot(wb, pivot)
            area = src.TableRange2
            if field.lower() not in [n.lower() for n in _field_names(src)]:
                raise ToolError(f"Field '{field}' not found. Fields: {_field_names(src)}")
        dest_ws = pick_sheet(wb, sheet) if sheet else src_ws
        if anchor_cell:
            r, c, _, _ = parse_a1(anchor_cell)
            anchor = dest_ws.Cells(r, c)
        else:
            anchor = dest_ws.Cells(int(area.Row), int(area.Column) + int(area.Columns.Count) + 1)
        left, top = float(anchor.Left), float(anchor.Top)
        if kind.lower() == "timeline":
            # Add2(Source, SourceField, Name, SlicerCacheType): xlTimeline = 2 (xlSlicer = 1); Name обязателен, иначе тип игнорируется
            sc = wb.SlicerCaches.Add2(src, field, f"Timeline_{field}", 2)
        else:
            sc = wb.SlicerCaches.Add2(src, field)
        # Slicers.Add(SlicerDestination, Level, Name, Caption, Top, Left, Width, Height); размеры при создании игнорируются —
        # поэтому позицию и размер задаём ниже через Shape
        sl = sc.Slicers.Add(dest_ws, MISSING, MISSING, caption or field, top, left, float(width), float(height))
        shape = sl.Shape
        shape.Left, shape.Top, shape.Width, shape.Height = left, top, float(width), float(height)
        if kind.lower() == "timeline" and timeline_level:
            levels = {"years": 0, "quarters": 1, "months": 2, "days": 3}
            if timeline_level.lower() not in levels:
                raise ToolError("timeline_level must be years, quarters, months or days.")
            sl.TimelineViewState.Level = levels[timeline_level.lower()]
        if kind.lower() != "timeline":
            if int(columns) > 1:
                sl.NumberOfColumns = int(columns)
            if style:
                sl.Style = style
        for pname in connect_pivots or []:
            _, pt2 = find_pivot(wb, pname)
            sc.PivotTables.AddPivotTable(pt2)
        out = _slicer_info(sc)
        return {"ok": True, "workbook": wb.Name, "slicer": sl.Name, **out}

    if not slicer:
        raise ToolError("'slicer' (name or field) is required.")
    sc = _find_cache(wb, slicer)

    if act == "select":
        if kind.lower() == "timeline" or date_from or date_to:
            if not (date_from and date_to):
                raise ToolError("Pass both date_from and date_to for a timeline.")
            d1 = datetime.datetime.strptime(date_from, "%Y-%m-%d")
            d2 = datetime.datetime.strptime(date_to, "%Y-%m-%d")
            sc.TimelineState.SetFilterDateRange(d1, d2)
        else:
            if not items:
                raise ToolError("'items' is required for select.")
            coll = sc.SlicerItems
            by_name = {}
            for k in range(1, int(coll.Count) + 1):
                it = coll(k)
                by_name[it.Name.lower()] = it
            for n in items:
                if str(n).lower() not in by_name:
                    raise ToolError(f"Item '{n}' not found. Items: {[i for i in by_name][:60]}")
            keep = {str(n).lower() for n in items}
            for n in keep:
                by_name[n].Selected = True  # сначала включаем нужные, иначе Excel не даст снять выбор со всех
            for n, it in by_name.items():
                if n not in keep:
                    it.Selected = False
        return {"ok": True, "workbook": wb.Name, **_slicer_info(sc)}

    if act == "clear":
        sc.ClearManualFilter()
        return {"ok": True, "workbook": wb.Name, **_slicer_info(sc)}

    if act == "connect":
        if not connect_pivots:
            raise ToolError("'connect_pivots' is required.")
        for pname in connect_pivots:
            _, pt2 = find_pivot(wb, pname)
            sc.PivotTables.AddPivotTable(pt2)
        return {"ok": True, "workbook": wb.Name, **_slicer_info(sc)}

    if act == "move":
        sl = sc.Slicers(1)
        shape = sl.Shape
        if anchor_cell:
            ws2 = pick_sheet(wb, sheet) if sheet else None
            r, c, _, _ = parse_a1(anchor_cell)
            anchor = (ws2 or shape.TopLeftCell.Worksheet).Cells(r, c)
            shape.Left, shape.Top = float(anchor.Left), float(anchor.Top)
        shape.Width, shape.Height = float(width), float(height)
        if caption:
            sl.Caption = caption
        if int(columns) > 1:
            sl.NumberOfColumns = int(columns)
        if style:
            sl.Style = style
        return {"ok": True, "workbook": wb.Name, **_slicer_info(sc)}

    if act == "delete":
        name = sc.Name
        sc.Delete()
        return {"ok": True, "workbook": wb.Name, "deleted": name}

    raise ToolError("action must be 'list', 'add', 'select', 'clear', 'connect', 'move' or 'delete'.")
