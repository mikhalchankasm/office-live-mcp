"""Excel conditional formatting and data validation."""

import re
import pywintypes
from .errors import ToolError
from .registry import office_tool
from .util import a1_cell, color_to_hex, parse_color
from .xl_common import (
    addr_of,
    bounds,
    delocalize_formulas,
    get_range,
    localize_formula,
    pick_sheet,
    pick_workbook,
)


CF_OPERATORS = {"between": 1, "not_between": 2, "equal": 3, "not_equal": 4, "greater": 5, "less": 6, "greater_equal": 7, "less_equal": 8}
ICON_SETS = {
    "3_arrows": 1, "3_arrows_gray": 2, "3_flags": 3, "3_traffic_lights": 4, "3_traffic_lights_rimmed": 5, "3_signs": 6,
    "3_symbols": 7, "3_symbols_circled": 8, "4_arrows": 9, "4_arrows_gray": 10, "4_red_to_black": 11, "4_ratings": 12,
    "4_traffic_lights": 13, "5_arrows": 14, "5_arrows_gray": 15, "5_ratings": 16, "5_quarters": 17,
}
TEXT_RULES = {"text_contains": 0, "text_not_contains": 1, "text_begins_with": 2, "text_ends_with": 3}
CF_TYPE_NAMES = {1: "cell_value", 2: "formula", 3: "color_scale", 4: "data_bar", 5: "top_bottom", 6: "icon_set", 8: "unique_duplicates", 9: "text", 10: "blanks", 12: "above_average", 13: "no_blanks", 16: "errors", 17: "no_errors"}
VALIDATION_TYPES = {"whole": 1, "decimal": 2, "list": 3, "date": 4, "time": 5, "text_length": 6, "custom": 7}
ALERT_STYLES = {"stop": 1, "warning": 2, "information": 3}

def _cf_operand(v) -> str | None:
    """Операнд правила: число -> как есть, '=ссылка/формула' -> как есть, прочее -> строка в кавычках."""
    if v is None:
        return None
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return str(v).replace(",", ".")
    s = str(v).strip()
    if s.startswith("="):
        return s
    if re.fullmatch(r"[+-]?\d+([.,]\d+)?", s):
        return s.replace(",", ".")
    return '="' + s.replace('"', '""') + '"'


def _style_condition(fc, fill_color, font_color, bold):
    if fill_color:
        fc.Interior.Color = parse_color(fill_color)
    if font_color:
        fc.Font.Color = parse_color(font_color)
    if bold is not None:
        fc.Font.Bold = bool(bold)


def cf_rules_detail(app, wb, ws, limit: int = 60) -> list[dict]:
    """Правила условного форматирования листа: диапазон, тип, формула (английская, относительно левой верхней ячейки),
    заливка/цвет/жирность, приоритет. Для чтения относительных формул временно выделяем левую верхнюю ячейку диапазона."""
    fcs = ws.Cells.FormatConditions
    n = int(fcs.Count)
    items, formulas = [], []
    prev = _remember_view(app)
    try:
        wb.Activate()
        ws.Activate()
        for i in range(1, min(n, limit) + 1):
            fc = fcs(i)
            t = int(fc.Type)
            applies = fc.AppliesTo
            r1, c1, _, _ = bounds(applies)
            ws.Range(a1_cell(r1, c1)).Select()
            item = {"index": i, "rule": CF_TYPE_NAMES.get(t, f"type_{t}"), "applies_to": addr_of(applies)}
            for attr in ("Priority", "StopIfTrue"):
                try:
                    item[attr[0].lower() + attr[1:]] = getattr(fc, attr)
                except (pywintypes.com_error, AttributeError):
                    pass
            if t in (1, 2, 9):
                try:
                    item["_f1"] = fc.Formula1
                    formulas.append(item["_f1"])
                except pywintypes.com_error:
                    pass
            if t == 1:
                try:
                    item["operator"] = {v: k for k, v in CF_OPERATORS.items()}.get(int(fc.Operator), int(fc.Operator))
                except pywintypes.com_error:
                    pass
            if t in (1, 2, 5, 8, 9, 10, 12, 13, 16, 17):
                try:
                    if int(fc.Interior.ColorIndex) != -4142:
                        item["fill"] = color_to_hex(fc.Interior.Color)
                except (pywintypes.com_error, TypeError):
                    pass
                try:
                    fcolor = color_to_hex(fc.Font.Color)
                    if fcolor and fcolor != "#000000":
                        item["font_color"] = fcolor
                    if fc.Font.Bold:
                        item["bold"] = True
                except (pywintypes.com_error, TypeError):
                    pass
            items.append(item)
    finally:
        _restore_view(app, prev)
    english = delocalize_formulas(app, wb, formulas)
    mapping = dict(zip(formulas, english, strict=False))
    for item in items:
        f = item.pop("_f1", None)
        if f is not None:
            item["formula"] = mapping.get(f, f)
    return items


@office_tool("excel_format", "write", title="Conditional formatting", read_actions=("list",), destructive=True)
def excel_conditional_format(
    workbook: str,
    sheet: str,
    action: str = "add",
    cells: str = "",
    rule: str = "cell_value",
    operator: str = "greater",
    value1: str | float | int | None = None,
    value2: str | float | int | None = None,
    formula: str = "",
    text: str = "",
    rank: int = 10,
    percent: bool = False,
    fill_color: str | None = None,
    font_color: str | None = None,
    bold: bool | None = None,
    min_color: str = "#F8696B",
    mid_color: str | None = "#FFEB84",
    max_color: str = "#63BE7B",
    bar_color: str = "#638EC6",
    icon_set: str = "3_traffic_lights",
) -> dict:
    """Add, list or clear conditional formatting ('highlight cells that ...').

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        action: 'add', 'list' (all rules on the sheet) or 'clear' (rules in `cells`, or the whole sheet when cells='').
        cells: range the rule applies to.
        rule: for 'add' - cell_value | formula | text_contains | text_not_contains | text_begins_with | text_ends_with | duplicates | unique | top | bottom | above_average | below_average | blanks | errors | color_scale | data_bar | icon_set.
        operator: for cell_value - between | not_between | equal | not_equal | greater | less | greater_equal | less_equal.
        value1, value2: operands for cell_value (numbers, text, or '=$A$1' references; value2 only for between).
        formula: for rule=formula - a boolean formula written for the TOP-LEFT cell of `cells`, e.g. '=$B2>$C2' (use $ to lock columns).
        text: for the text_* rules. rank/percent: for top/bottom (top 10 items, or top 10 percent).
        fill_color, font_color, bold: the formatting applied to matching cells (color_scale/data_bar/icon_set ignore them).
        min_color, mid_color (null for a 2-color scale), max_color: color_scale colors. bar_color: data_bar color. icon_set: 3_arrows, 3_traffic_lights, 3_symbols, 3_flags, 4_arrows, 5_arrows, 5_quarters, ...
    """
    app, wb = pick_workbook(workbook)
    ws = pick_sheet(wb, sheet)
    act = action.lower()
    if act == "list":
        return {"workbook": wb.Name, "sheet": ws.Name, "rules": cf_rules_detail(app, wb, ws)}
    if act == "clear":
        target = get_range(wb, sheet, cells)[1] if cells else ws.Cells
        target.FormatConditions.Delete()
        return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cleared": addr_of(target) if cells else "entire sheet"}
    if act != "add":
        raise ToolError("action must be 'add', 'list' or 'clear'.")
    if not cells:
        raise ToolError("'cells' is required for 'add'.")
    ws, rng = get_range(wb, sheet, cells, empty_means_used=False)
    fcs = rng.FormatConditions
    r = rule.lower()
    if r == "cell_value":
        if operator.lower() not in CF_OPERATORS:
            raise ToolError(f"operator must be one of {sorted(CF_OPERATORS)}")
        if value1 is None:
            raise ToolError("value1 is required for cell_value.")
        op = CF_OPERATORS[operator.lower()]
        if op in (1, 2) and value2 is None:
            raise ToolError("value2 is required for between/not_between.")
        fc = fcs.Add(1, op, _cf_operand(value1), _cf_operand(value2))
        _style_condition(fc, fill_color, font_color, bold)
    elif r == "formula":
        if not formula.startswith("="):
            raise ToolError("'formula' must start with '=' and be written for the top-left cell of the range.")
        r1, c1, _, _ = bounds(rng)
        # формула условного форматирования трактуется относительно АКТИВНОЙ ячейки: временно выделяем левую верхнюю
        prev = _remember_view(app)
        try:
            wb.Activate()
            ws.Activate()
            ws.Range(a1_cell(r1, c1)).Select()
            fc = fcs.Add(2, 1, localize_formula(app, wb, formula))  # Add(Type=xlExpression, Operator, Formula1)
            _style_condition(fc, fill_color, font_color, bold)
        finally:
            _restore_view(app, prev)
    elif r in TEXT_RULES:
        if not text:
            raise ToolError("'text' is required for text rules.")
        # Add(Type=xlTextString, Operator, Formula1, Formula2, String, TextOperator)
        fc = fcs.Add(9, 1, None, None, text, TEXT_RULES[r])
        _style_condition(fc, fill_color, font_color, bold)
    elif r in ("duplicates", "unique"):
        fc = fcs.AddUniqueValues()
        fc.DupeUnique = 1 if r == "duplicates" else 0
        _style_condition(fc, fill_color or "#FFC7CE", font_color, bold)
    elif r in ("top", "bottom"):
        fc = fcs.AddTop10()
        fc.TopBottom = 1 if r == "top" else 2
        fc.Rank = int(rank)
        fc.Percent = bool(percent)
        _style_condition(fc, fill_color or "#C6EFCE", font_color, bold)
    elif r in ("above_average", "below_average"):
        fc = fcs.AddAboveAverage()
        fc.AboveBelow = 0 if r == "above_average" else 1
        _style_condition(fc, fill_color or "#C6EFCE", font_color, bold)
    elif r in ("blanks", "errors"):
        fc = fcs.Add(10 if r == "blanks" else 16, 1)
        _style_condition(fc, fill_color or "#FFC7CE", font_color, bold)
    elif r == "color_scale":
        if mid_color:
            cs = fcs.AddColorScale(3)
            cs.ColorScaleCriteria(1).FormatColor.Color = parse_color(min_color)
            cs.ColorScaleCriteria(2).FormatColor.Color = parse_color(mid_color)
            cs.ColorScaleCriteria(3).FormatColor.Color = parse_color(max_color)
        else:
            cs = fcs.AddColorScale(2)
            cs.ColorScaleCriteria(1).FormatColor.Color = parse_color(min_color)
            cs.ColorScaleCriteria(2).FormatColor.Color = parse_color(max_color)
    elif r == "data_bar":
        db = fcs.AddDatabar()
        db.BarColor.Color = parse_color(bar_color)
    elif r == "icon_set":
        if icon_set not in ICON_SETS:
            raise ToolError(f"icon_set must be one of {sorted(ICON_SETS)}")
        ic = fcs.AddIconSetCondition()
        ic.IconSet = wb.IconSets(ICON_SETS[icon_set])
    else:
        raise ToolError("rule must be one of: cell_value, formula, text_contains, text_not_contains, text_begins_with, text_ends_with, duplicates, unique, top, bottom, above_average, below_average, blanks, errors, color_scale, data_bar, icon_set")
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "applies_to": addr_of(rng), "rule": r, "rules_on_range": int(rng.FormatConditions.Count)}


def _remember_view(app) -> dict:
    """Запоминает активную книгу/лист/выделение пользователя, чтобы вернуть после временной активации."""
    state = {"wb": app.ActiveWorkbook, "ws": app.ActiveSheet, "sel": None}
    try:
        sel = app.Selection
        state["sel"] = sel.Address
    except (pywintypes.com_error, AttributeError):
        pass
    return state


def _restore_view(app, state: dict):
    try:
        if state["wb"] is not None:
            state["wb"].Activate()
        if state["ws"] is not None:
            state["ws"].Activate()
            if state["sel"]:
                state["ws"].Range(state["sel"]).Select()
    except pywintypes.com_error:
        pass


# ================================================================== проверка данных


@office_tool("excel_format", "write", title="Data validation", read_actions=("get",), destructive=True)
def excel_data_validation(
    workbook: str,
    sheet: str,
    cells: str,
    action: str = "set",
    type: str = "list",
    items: list[str] | None = None,
    source_range: str = "",
    operator: str = "between",
    value1: str | float | int | None = None,
    value2: str | float | int | None = None,
    allow_blank: bool = True,
    show_dropdown: bool = True,
    alert_style: str = "stop",
    error_title: str = "",
    error_message: str = "",
    prompt_title: str = "",
    prompt_message: str = "",
) -> dict:
    """Restrict what can be typed into cells (dropdown lists, number/date limits, custom rules), read the rule, or remove it.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        cells: target range.
        action: 'set' (replaces any existing rule), 'get' (rule of the top-left cell) or 'clear'.
        type: list | whole | decimal | date | time | text_length | custom.
        items: for type=list - the dropdown choices, e.g. ["Yes","No","Maybe"].
        source_range: for type=list - alternatively a range holding the choices, e.g. 'Lists!$A$1:$A$10' or '=$H$2:$H$9'.
        operator: for whole/decimal/date/time/text_length - between | not_between | equal | not_equal | greater | less | greater_equal | less_equal.
        value1, value2: limits (value2 only for between/not_between). For type=custom value1 is a formula like '=COUNTIF($A:$A,A1)=1'.
        allow_blank, show_dropdown: behavior flags. alert_style: stop | warning | information.
        error_title, error_message: shown when input is rejected. prompt_title, prompt_message: hint shown when the cell is selected.
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells, empty_means_used=False)
    act = action.lower()
    if act == "clear":
        rng.Validation.Delete()
        return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cleared": addr_of(rng)}
    if act == "get":
        first = ws.Range(a1_cell(*bounds(rng)[:2]))
        try:
            v = first.Validation
            return {"workbook": wb.Name, "sheet": ws.Name, "cell": addr_of(first), "has_validation": True, "type": int(v.Type), "formula1": v.Formula1, "formula2": _safe(lambda: v.Formula2)}
        except pywintypes.com_error:
            return {"workbook": wb.Name, "sheet": ws.Name, "cell": addr_of(first), "has_validation": False}
    if act != "set":
        raise ToolError("action must be 'set', 'get' or 'clear'.")
    t = type.lower()
    if t not in VALIDATION_TYPES:
        raise ToolError(f"type must be one of {sorted(VALIDATION_TYPES)}")
    if alert_style.lower() not in ALERT_STYLES:
        raise ToolError("alert_style must be stop, warning or information.")
    if operator.lower() not in CF_OPERATORS:
        raise ToolError(f"operator must be one of {sorted(CF_OPERATORS)}")
    op = CF_OPERATORS[operator.lower()]
    f1 = f2 = None
    if t == "list":
        if items:
            sep = app.International[4]  # разделитель списков из региональных настроек Excel (в RU это ';')
            joined = sep.join(str(x) for x in items)
            if len(joined) > 255:
                raise ToolError("The choices are longer than 255 characters in total; put them in a range and use source_range.")
            if any(sep in str(x) for x in items):
                raise ToolError(f"A choice contains the list separator {sep!r}; use source_range instead.")
            f1 = joined
        elif source_range:
            f1 = source_range if source_range.startswith("=") else "=" + source_range
        else:
            raise ToolError("For type=list pass `items` or `source_range`.")
    elif t == "custom":
        if not value1 or not str(value1).startswith("="):
            raise ToolError("For type=custom value1 must be a formula starting with '='.")
        f1 = localize_formula(app, wb, str(value1))
    else:
        if value1 is None:
            raise ToolError("value1 is required.")
        f1 = _cf_operand_raw(value1)
        if op in (1, 2):
            if value2 is None:
                raise ToolError("value2 is required for between/not_between.")
            f2 = _cf_operand_raw(value2)
    v = rng.Validation
    v.Delete()
    # Add(Type, AlertStyle, Operator, Formula1, Formula2)
    v.Add(VALIDATION_TYPES[t], ALERT_STYLES[alert_style.lower()], op, f1, f2)
    v.IgnoreBlank = bool(allow_blank)
    if t == "list":
        v.InCellDropdown = bool(show_dropdown)
    if error_title or error_message:
        v.ShowError = True
        v.ErrorTitle = error_title[:32]
        v.ErrorMessage = error_message[:255]
    if prompt_title or prompt_message:
        v.ShowInput = True
        v.InputTitle = prompt_title[:32]
        v.InputMessage = prompt_message[:255]
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cells": addr_of(rng), "type": t, "formula1": f1, "formula2": f2}


def _cf_operand_raw(v) -> str:
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return str(v).replace(",", ".")
    return str(v)


def _safe(fn):
    try:
        return fn()
    except (pywintypes.com_error, AttributeError):
        return None
