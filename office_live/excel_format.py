"""Excel: форматирование, условное форматирование, валидация, ссылки, примечания, вид листа, картинки."""


import inspect
import io
import json
import os
import re
import shutil
import sys
import tempfile
import time

import pythoncom
import pywintypes
from mcp.server.mcpserver import Image

from . import com, config, undo
from . import xl_common as xl
from .errors import PartialChangeError, ToolError
from .registry import office_tool
from .safety import IMAGE_EXTS, check_path
from .util import cm_to_points, color_to_hex, parse_a1, parse_color, quote_sheet, to_grid
from .xl_common import (
    addr_of,
    bounds,
    count_nonempty,
    get_range,
    pick_sheet,
    pick_workbook,
    read_number_format,
    set_number_format,
    suspend_events,
)

# ------------------------------------------------------------------ словари значений

NUMBER_FORMATS = {
    "general": "General", "text": "@", "number": "0.00", "integer": "0", "thousands": "#,##0",
    "thousands_2": "#,##0.00", "currency": "#,##0.00 \"₽\"", "currency_usd": "$#,##0.00", "currency_eur": "#,##0.00 \"€\"",
    "percent": "0.0%", "percent_0": "0%", "date": "dd.mm.yyyy", "date_iso": "yyyy-mm-dd", "datetime": "dd.mm.yyyy hh:mm",
    "time": "hh:mm:ss", "scientific": "0.00E+00",
}
H_ALIGN = {"left": -4131, "center": -4108, "right": -4152, "justify": -4130, "fill": 5, "distributed": -4117, "general": 1, "center_across": 7}
V_ALIGN = {"top": -4160, "center": -4108, "bottom": -4107, "justify": -4130, "distributed": -4117}
H_NAMES = {v: k for k, v in H_ALIGN.items()}
V_NAMES = {v: k for k, v in V_ALIGN.items()}
BORDER_STYLES = {  # имя -> (LineStyle, Weight)
    "thin": (1, 2), "medium": (1, -4138), "thick": (1, 4), "dashed": (-4115, 2), "dotted": (-4118, 2),
    "double": (-4119, 4), "hair": (1, 1),
}
EDGES = {"left": 7, "top": 8, "bottom": 9, "right": 10, "inside_vertical": 11, "inside_horizontal": 12}


MISSING = pythoncom.Missing

# Order is Worksheet.Protect's positional Allow* tail, verified against the type library.
PROTECTION_ALLOW = dict(zip(
    ("format_cells", "format_columns", "format_rows", "insert_columns", "insert_rows", "insert_hyperlinks",
     "delete_columns", "delete_rows", "sort", "filter", "pivot_tables"),
    ("AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows", "AllowInsertingColumns", "AllowInsertingRows",
     "AllowInsertingHyperlinks", "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting", "AllowFiltering", "AllowUsingPivotTables"),
    strict=True,
))
PASSWORD_BARRIER = "Password operation cannot be undone: re-protection/unprotection would require retaining the password."


def _protection_state(target, scope):
    if scope == "workbook":
        return {"structure": bool(target.ProtectStructure), "windows": bool(target.ProtectWindows)}
    return {"contents": bool(target.ProtectContents), "objects": bool(target.ProtectDrawingObjects),
            "scenarios": bool(target.ProtectScenarios), "user_interface_only": bool(target.ProtectionMode),
            "allow": {k: bool(getattr(target.Protection, p)) for k, p in PROTECTION_ALLOW.items()},
            "enable_selection": int(target.EnableSelection)}


def _protected(state, scope):
    return any(state[p] for p in (("structure", "windows") if scope == "workbook" else ("contents", "objects", "scenarios")))


def _protect(target, scope, password, state):
    if scope == "workbook":
        target.Protect(password, state["structure"], state["windows"])
    else:
        target.Protect(password, state["objects"], state["contents"], state["scenarios"], state["user_interface_only"],
                       *(state["allow"][k] for k in PROTECTION_ALLOW))


@office_tool("excel_format", "write", title="Sheet/workbook protection", read_actions=("status",))
def excel_protection(workbook: str = "", sheet: str = "", scope: str = "sheet", action: str = "status",
                     password: str = "", allow: list[str] | None = None, unlocked_cells: list[str] | None = None) -> dict:
    """Inspect, protect or unprotect a sheet/workbook. Status is available in readonly mode. Passwords are redacted from journals, audit, errors and history; password operations create an undo barrier and never retain the password.

    Args:
        workbook, sheet: target names; sheet required for sheet changes, unused for workbook scope.
        scope: sheet or workbook (workbook protect enables Structure=True, Windows=False).
        action: status, protect or unprotect. Already protected targets must be unprotected first; unprotect on an unprotected target is a no-op.
        password: optional case-sensitive password; explicitly passed even when empty to avoid password prompts. Status never guesses whether a protection password exists.
        allow: sheet protect only: format_cells, format_columns, format_rows, insert_columns, insert_rows, insert_hyperlinks, delete_columns, delete_rows, sort, filter, pivot_tables, edit_objects, edit_scenarios.
        unlocked_cells: sheet protect only: list of same-sheet rectangles to set Locked=False before protecting (total <=100000 cells, no merges).

    Without a password, undo restores the original protection flags and Locked properties and detects later changes to them.
    Read-only workbooks are refused. Partial COM errors retain undo or the explicit password barrier; no raw protection COM errors are returned.
    """
    action = action.lower()
    if scope not in {"sheet", "workbook"} or action not in {"status", "protect", "unprotect"}:
        raise ToolError("scope must be sheet/workbook; action must be status/protect/unprotect.")
    if not isinstance(password, str):
        raise ToolError("password must be a string.")
    allowed = set(PROTECTION_ALLOW) | {"edit_objects", "edit_scenarios"}
    if allow is not None and (not isinstance(allow, list) or any(not isinstance(a, str) or a not in allowed for a in allow) or len(set(allow)) != len(allow)):
        raise ToolError("allow must be a unique list of supported permissions.")
    if unlocked_cells is not None and (not isinstance(unlocked_cells, list) or any(not isinstance(c, str) or not c.strip() for c in unlocked_cells)):
        raise ToolError("unlocked_cells must be a list of nonempty rectangles.")
    if (allow or unlocked_cells) and (scope != "sheet" or action != "protect"):
        raise ToolError("allow/unlocked_cells require sheet protect.")
    if action == "status" and password:
        raise ToolError("status does not accept a password.")
    app, wb = pick_workbook(workbook)
    target = pick_sheet(wb, sheet) if scope == "sheet" else wb
    before = _protection_state(target, scope)
    result = {"workbook": wb.Name, "scope": scope, "action": action, "state": before,
              "workbook_structure": bool(wb.ProtectStructure)}
    if scope == "sheet":
        result["sheet"] = target.Name
    if action == "status":
        return result
    if wb.ReadOnly:
        raise ToolError("Workbook is read-only; no protection was changed.")
    if action == "protect" and _protected(before, scope):
        raise ToolError("Target is already protected; unprotect it first.")
    if action == "unprotect" and not _protected(before, scope):
        return {**result, "changed": False}
    ranges, size = [], 0
    for address in unlocked_cells or []:
        _, rng = get_range(wb, target.Name, address, empty_means_used=False)
        if (int(rng.Areas.Count) != 1 or rng.Worksheet.Name != target.Name or rng.Worksheet.Parent.FullName != wb.FullName):
            raise ToolError("unlocked_cells must be rectangles on the selected sheet.")
        size += int(rng.Rows.Count) * int(rng.Columns.Count)
        if size > 100000:
            raise ToolError("unlocked_cells exceeds 100000 cells; use smaller rectangles.")
        if rng.MergeCells is not False:
            raise ToolError("unlocked_cells contains merged cells.")
        ranges.append(rng)
    from .undo_excel import _property_state, _restore_properties
    from .undo import cancel_prepared, require_undo

    # Snapshot Locked even for password barriers: best-effort local rollback on failed Protect.
    locked = [(rng, _property_state(rng, "Locked")) for rng in ranges]
    desired = {"structure": True, "windows": False} if scope == "workbook" else {
        "contents": True, "objects": "edit_objects" not in (allow or []), "scenarios": "edit_scenarios" not in (allow or []),
        "user_interface_only": False, "allow": {k: k in (allow or []) for k in PROTECTION_ALLOW}}
    require_undo("workbook", app, wb, plan={"_protection_sheet": target.Name if scope == "sheet" else "",
                                          "_locked_ranges": [addr_of(rng) for rng in ranges]},
                 barrier_reason=PASSWORD_BARRIER if password else "")
    applied = 0
    try:
        if action == "protect":
            for rng in ranges:
                rng.Locked = False
                applied += int(rng.Rows.Count) * int(rng.Columns.Count)
            _protect(target, scope, password, desired)
        else:
            target.Unprotect(password)  # Never omit: a missing password can open a modal prompt.
        result["state"] = _protection_state(target, scope)
        result["workbook_structure"] = bool(wb.ProtectStructure)
        return {**result, "changed": True, "unlocked_cells_applied": applied}
    except (pywintypes.com_error, ToolError):
        # COM descriptions may echo the password; do not even translate these exceptions.
        rolled_back = False
        try:
            current = _protection_state(target, scope)
            if current == before:
                for rng, value in locked:
                    _restore_properties(rng, {"Locked": value})
                rolled_back = True
        except (pywintypes.com_error, ToolError):
            pass
        if rolled_back:
            cancel_prepared("workbook", app, wb)
            raise ToolError("Неверный пароль или Excel отклонил операцию защиты; состояние восстановлено.") from None
        raise PartialChangeError(f"Protection interrupted: unlocked_cells_applied={applied}; protection state may be partially changed.") from None


def _num_format(fmt: str) -> str:
    return NUMBER_FORMATS.get(fmt.strip().lower(), fmt)


# ================================================================== форматирование


@office_tool("excel_format", "write", title="Format range")
def excel_format_range(
    workbook: str,
    sheet: str,
    cells: str,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    font_color: str | None = None,
    fill_color: str | None = None,
    number_format: str | None = None,
    horizontal_alignment: str | None = None,
    vertical_alignment: str | None = None,
    wrap_text: bool | None = None,
    shrink_to_fit: bool | None = None,
    indent: int | None = None,
    text_rotation: int | None = None,
    borders: str | None = None,
    border_style: str = "thin",
    border_color: str = "#000000",
    style: str | None = None,
) -> dict:
    """Format a range. Only the properties you pass are changed. Colors: '#RRGGBB' or a name (red, lightblue, lightgreen, lightyellow, ...); fill_color='none' removes the fill.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        cells: range, e.g. 'A1:D1'.
        bold, italic, underline, strikethrough: font styles.
        font_name: e.g. 'Calibri'. font_size: points. font_color / fill_color: colors.
        number_format: Excel format code ('0.00', '#,##0', 'dd.mm.yyyy', '0.0%', '@' for text) or a shortcut: general, text, number, integer, thousands, thousands_2, currency, currency_usd, currency_eur, percent, percent_0, date, date_iso, datetime, time, scientific.
        horizontal_alignment: left|center|right|justify|fill|distributed|general|center_across. vertical_alignment: top|center|bottom|justify|distributed.
        wrap_text, shrink_to_fit: text fitting. indent: indent level. text_rotation: degrees -90..90.
        borders: all|outline|inside|top|bottom|left|right|none. border_style: thin|medium|thick|dashed|dotted|double|hair. border_color: color of borders.
        style: apply a named cell style first (e.g. 'Good', 'Bad', 'Neutral', 'Heading 1', 'Total').
    """
    options = locals().copy()
    for key in ("workbook", "sheet", "cells"):
        options.pop(key)
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells, empty_means_used=False)
    return _prepare_format(app, wb, ws, rng, **options)()


def _prepare_format(app, wb, ws, rng, *,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    font_color: str | None = None,
    fill_color: str | None = None,
    number_format: str | None = None,
    horizontal_alignment: str | None = None,
    vertical_alignment: str | None = None,
    wrap_text: bool | None = None,
    shrink_to_fit: bool | None = None,
    indent: int | None = None,
    text_rotation: int | None = None,
    borders: str | None = None,
    border_style: str = "thin",
    border_color: str = "#000000",
    style: str | None = None,
):
    # Validate the complete format before the first assignment.
    if not any(v is not None for v in (style, bold, italic, underline, strikethrough, font_size, font_color, fill_color, number_format,
                                        horizontal_alignment, vertical_alignment, wrap_text, shrink_to_fit, indent, text_rotation, borders)) and not font_name:
        raise ToolError("Nothing to apply: pass at least one formatting parameter.")
    h_code = v_code = None
    if horizontal_alignment is not None:
        if horizontal_alignment.lower() not in H_ALIGN:
            raise ToolError(f"horizontal_alignment must be one of {sorted(H_ALIGN)}")
        h_code = H_ALIGN[horizontal_alignment.lower()]
    if vertical_alignment is not None:
        if vertical_alignment.lower() not in V_ALIGN:
            raise ToolError(f"vertical_alignment must be one of {sorted(V_ALIGN)}")
        v_code = V_ALIGN[vertical_alignment.lower()]
    font_rgb = parse_color(font_color) if font_color is not None else None
    fill_none = fill_color is not None and fill_color.strip().lower() == "none"
    fill_rgb = parse_color(fill_color) if fill_color is not None and not fill_none else None
    if font_size is not None and not 1 <= float(font_size) <= 409:
        raise ToolError("font_size must be between 1 and 409 points.")
    if indent is not None and not 0 <= int(indent) <= 250:
        raise ToolError("indent must be between 0 and 250.")
    if text_rotation is not None and not -90 <= int(text_rotation) <= 90:
        raise ToolError("text_rotation must be between -90 and 90 degrees.")
    border_plan = _border_plan(rng, borders.lower(), border_style.lower(), border_color) if borders is not None else None


    def apply():
        applied = []

        # Стиль и формат числа Excel может отвергнуть только при применении — делаем их первыми, чтобы отказ не оставил частичного оформления
        if style is not None:
            try:
                rng.Style = style
            except pywintypes.com_error:
                raise ToolError(f"Cell style '{style}' not found (style names depend on the Excel UI language). Nothing was changed.") from None
            applied.append(f"style={style}")
        if number_format is not None:
            try:
                set_number_format(app, rng, number_format, NUMBER_FORMATS)
            except pywintypes.com_error:
                raise ToolError(
                    f"Excel rejected number_format '{number_format}'. Use codes like '0', '0.00', '#,##0.00', '0.0%', 'dd.mm.yyyy', '@' or a shortcut ({sorted(NUMBER_FORMATS)})."
                    + (f" Already applied: {applied}." if applied else " Nothing was changed.")
                ) from None
            applied.append(f"number_format={number_format}")
        font = rng.Font
        if bold is not None:
            font.Bold = bool(bold)
            applied.append(f"bold={bold}")
        if italic is not None:
            font.Italic = bool(italic)
            applied.append(f"italic={italic}")
        if underline is not None:
            font.Underline = 2 if underline else -4142
            applied.append(f"underline={underline}")
        if strikethrough is not None:
            font.Strikethrough = bool(strikethrough)
            applied.append(f"strikethrough={strikethrough}")
        if font_name:
            font.Name = font_name
            applied.append(f"font_name={font_name}")
        if font_size is not None:
            font.Size = float(font_size)
            applied.append(f"font_size={font_size}")
        if font_rgb is not None:
            font.Color = font_rgb
            applied.append(f"font_color={font_color}")
        if fill_color is not None:
            if fill_none:
                rng.Interior.ColorIndex = -4142
            else:
                rng.Interior.Color = fill_rgb
            applied.append(f"fill_color={fill_color}")
        if h_code is not None:
            rng.HorizontalAlignment = h_code
            applied.append(f"horizontal_alignment={horizontal_alignment.lower()}")
        if v_code is not None:
            rng.VerticalAlignment = v_code
            applied.append(f"vertical_alignment={vertical_alignment.lower()}")
        if wrap_text is not None:
            rng.WrapText = bool(wrap_text)
            applied.append(f"wrap_text={wrap_text}")
        if shrink_to_fit is not None:
            rng.ShrinkToFit = bool(shrink_to_fit)
            applied.append(f"shrink_to_fit={shrink_to_fit}")
        if indent is not None:
            rng.IndentLevel = int(indent)
            applied.append(f"indent={indent}")
        if text_rotation is not None:
            rng.Orientation = int(text_rotation)
            applied.append(f"text_rotation={text_rotation}")
        if border_plan is not None:
            _apply_border_plan(rng, border_plan)
            applied.append(f"borders={borders}")
        return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cells": addr_of(rng), "applied": applied}

    return apply


@office_tool("excel_format", "write", title="Format multiple ranges")
def excel_format_ranges(workbook: str, sheet: str, ranges: list[dict]) -> dict:
    """Apply several {cells, format} objects in order with one office_undo entry. All ranges and formatting arguments are validated before editing.

    Args:
        workbook: exact workbook name.
        sheet: selected sheet name.
        ranges: nonempty list (at most 128) of {'cells': 'A1:D1', 'format': {'bold': True}}.
            format accepts the existing excel_format_range parameters; overlapping ranges apply in list order.

    Combined range sizes must fit OFFICE_LIVE_UNDO_MAX_CELLS when undo is enabled.
    On partial Excel errors, office_undo restores every range to its pre-call state.
    """
    if not isinstance(ranges, list) or not 1 <= len(ranges) <= 128:
        raise ToolError("ranges must contain 1 to 128 {cells, format} objects.")
    signature = inspect.signature(_prepare_format)
    bool_keys = {"bold", "italic", "underline", "strikethrough", "wrap_text", "shrink_to_fit"}
    number_keys = {"font_size", "indent", "text_rotation"}
    for spec in ranges:
        if not isinstance(spec, dict) or set(spec) != {"cells", "format"}:
            raise ToolError("Each range must contain exactly cells and format.")
        if not isinstance(spec["cells"], str) or not spec["cells"].strip() or not isinstance(spec["format"], dict):
            raise ToolError("Each range needs nonempty cells and a format object.")
        for key, value in spec["format"].items():
            if key not in signature.parameters or key in {"app", "wb", "ws", "rng"}:
                raise ToolError(f"Unknown formatting parameter: {key}.")
            if value is None:
                continue
            expected = bool if key in bool_keys else (int, float) if key in number_keys else str
            if not isinstance(value, expected) or (key in number_keys and isinstance(value, bool)):
                raise ToolError(f"Invalid value for formatting parameter: {key}.")
    app, wb = pick_workbook(workbook)
    plans, specs = [], []
    for spec in ranges:
        _, rng = get_range(wb, sheet, spec["cells"], empty_means_used=False)
        ws = rng.Worksheet
        xl.validate_rectangle(rng, ws, wb)
        plans.append(_prepare_format(app, wb, ws, rng, **spec["format"]))
        specs.append({"cells": xl.ref_label(ws, rng), "format": spec["format"]})
    undo.require_undo("workbook", app, wb, plan={"sheet": "", "_format_ranges": specs})
    results = []
    try:
        for apply in plans:
            results.append(apply())
    except (pywintypes.com_error, ToolError) as exc:
        raise PartialChangeError(f"Formatting interrupted after {len(results)} completed ranges. {com.translate(exc)}") from None
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "ranges": results}


def _border_plan(rng, which: str, style: str, color: str) -> tuple:
    """Проверяет параметры границ и возвращает (границы, режим, линия, толщина, цвет) — без обращений к документу, кроме размеров."""
    r1, c1, r2, c2 = bounds(rng)
    multi_rows, multi_cols = r2 > r1, c2 > c1
    inside = (["inside_vertical"] if multi_cols else []) + (["inside_horizontal"] if multi_rows else [])
    groups = {
        "all": ["left", "top", "bottom", "right"] + inside,
        "outline": ["left", "top", "bottom", "right"],
        "inside": inside,
        "top": ["top"], "bottom": ["bottom"], "left": ["left"], "right": ["right"],
        "none": ["left", "top", "bottom", "right"] + inside,
    }
    if which not in groups:
        raise ToolError(f"borders must be one of {sorted(groups)}")
    if style not in BORDER_STYLES:
        raise ToolError(f"border_style must be one of {sorted(BORDER_STYLES)}")
    line, weight = BORDER_STYLES[style]
    return groups[which], which, line, weight, parse_color(color)


def _apply_border_plan(rng, plan: tuple):
    edges, which, line, weight, col = plan
    for edge in edges:
        b = rng.Borders(EDGES[edge])
        if which == "none":
            b.LineStyle = -4142
        else:
            b.LineStyle = line
            b.Weight = weight
            b.Color = col


@office_tool("excel_format", "read", title="Inspect formatting")
def excel_get_format(workbook: str = "", sheet: str = "", cells: str = "A1") -> dict:
    """Read the formatting of a range. A property is reported as 'mixed' when the cells of the range differ - narrow the range to see each value.

    Args:
        workbook: exact name or '' for the active workbook.
        sheet: sheet name or '' for active.
        cells: range to inspect.
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells)

    def val(v, fn=lambda x: x):
        return "mixed" if v is None else fn(v)

    interior = rng.Interior
    fill = "mixed" if interior.ColorIndex is None else (None if int(interior.ColorIndex) == -4142 else color_to_hex(interior.Color))
    font = rng.Font
    out = {
        "workbook": wb.Name, "sheet": ws.Name, "cells": addr_of(rng),
        "font_name": val(font.Name), "font_size": val(font.Size),
        "bold": val(font.Bold, bool), "italic": val(font.Italic, bool),
        "underline": val(font.Underline, lambda u: int(u) != -4142),
        "strikethrough": val(font.Strikethrough, bool),
        "font_color": val(font.Color, color_to_hex), "fill_color": fill,
        "number_format": val(read_number_format(app, rng)),
        "horizontal_alignment": val(rng.HorizontalAlignment, lambda a: H_NAMES.get(int(a), int(a))),
        "vertical_alignment": val(rng.VerticalAlignment, lambda a: V_NAMES.get(int(a), int(a))),
        "wrap_text": val(rng.WrapText, bool),
        "merged": val(rng.MergeCells, bool),
        "column_width": val(rng.ColumnWidth), "row_height": val(rng.RowHeight),
    }
    try:
        out["style"] = rng.Style.Name
    except (pywintypes.com_error, AttributeError):
        pass
    try:
        out["conditional_formats"] = int(rng.FormatConditions.Count)
    except pywintypes.com_error:
        pass
    return out


# ================================================================== условное форматирование


from .excel_cf import (
    CF_OPERATORS as CF_OPERATORS,
    ICON_SETS as ICON_SETS,
    TEXT_RULES as TEXT_RULES,
    CF_TYPE_NAMES as CF_TYPE_NAMES,
    VALIDATION_TYPES as VALIDATION_TYPES,
    ALERT_STYLES as ALERT_STYLES,
    _cf_operand as _cf_operand,
    _style_condition as _style_condition,
    cf_rules_detail as cf_rules_detail,
    excel_conditional_format as excel_conditional_format,
    _remember_view as _remember_view,
    _restore_view as _restore_view,
    excel_data_validation as excel_data_validation,
    _cf_operand_raw as _cf_operand_raw,
    _safe as _safe,
)


# ================================================================== ссылки, примечания, картинки


@office_tool("excel_format", "write", title="Add hyperlink")
def excel_add_hyperlink(workbook: str, sheet: str, cell: str, url: str = "", target_cell: str = "", target_sheet: str = "", text: str = "", tooltip: str = "") -> dict:
    """Turn a cell into a hyperlink - to a web/mail address or to another place in the workbook.

    Args:
        workbook: exact workbook name.
        sheet: sheet holding the link cell ('' = active).
        cell: the cell that becomes the link.
        url: 'https://...' or 'mailto:...' (leave empty for an internal link).
        target_cell, target_sheet: for an internal link, e.g. target_sheet='Summary', target_cell='A1'.
        text: displayed text (default: keep the cell's current text or the URL).
        tooltip: hover text.
    """
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cell, empty_means_used=False)
    if url:
        if not re.match(r"^(https?://|mailto:|ftp://)", url, re.IGNORECASE):
            raise ToolError("Only http(s), ftp and mailto links are allowed. For other places use target_sheet/target_cell.")
        address, sub = url, ""
    elif target_cell:
        tws = pick_sheet(wb, target_sheet) if target_sheet else ws
        address, sub = "", f"{quote_sheet(tws.Name)}!{target_cell}"
    else:
        raise ToolError("Pass `url` or `target_cell`.")
    # Add(Anchor, Address, SubAddress, ScreenTip, TextToDisplay): None недопустим (E_INVALIDARG), пропуск — Missing
    ws.Hyperlinks.Add(rng, address, sub or MISSING, tooltip or MISSING, text or MISSING)
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cell": addr_of(rng), "link": url or sub}


@office_tool("excel_format", "write", title="Cell notes", read_actions=("list",), destructive=True)
def excel_manage_comments(workbook: str, sheet: str, action: str = "list", cells: str = "", text: str = "") -> dict:
    """List, add/replace or delete cell notes (the yellow-sticker comments).

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        action: 'list', 'add' (replaces an existing note on the cell) or 'delete'.
        cells: the cell (add) or range (delete) for the note.
        text: note text for 'add'.
    """
    app, wb = pick_workbook(workbook)
    ws = pick_sheet(wb, sheet)
    act = action.lower()
    if act == "list":
        items = []
        for i in range(1, int(ws.Comments.Count) + 1):
            c = ws.Comments(i)
            items.append({"cell": c.Parent.Address.replace("$", ""), "author": _safe(lambda c=c: c.Author), "text": c.Text()})
        return {"workbook": wb.Name, "sheet": ws.Name, "notes": items}
    _, rng = get_range(wb, sheet, cells, empty_means_used=False)
    if act == "add":
        if not text:
            raise ToolError("'text' is required.")
        rng.ClearComments()
        rng.AddComment(text)
        return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cell": addr_of(rng)}
    if act == "delete":
        rng.ClearComments()
        return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cleared": addr_of(rng)}
    raise ToolError("action must be 'list', 'add' or 'delete'.")


@office_tool("excel_format", "write", title="Insert picture")
def excel_insert_image(workbook: str, sheet: str, path: str, cell: str = "A1", width: float | None = None, name: str = "") -> dict:
    """Insert a picture file (png/jpg/gif/bmp/emf) onto a sheet, anchored at a cell.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        path: full path of the image file.
        cell: top-left anchor cell.
        width: width in points (aspect ratio is kept); default - native size.
        name: optional shape name.
    """
    app, wb = pick_workbook(workbook)
    ws, anchor = get_range(wb, sheet, cell, empty_means_used=False)
    full = check_path(path, "read")
    if not os.path.isfile(full):
        raise ToolError(f"Image file not found: {full}")
    if full.rsplit(".", 1)[-1].lower() not in IMAGE_EXTS:
        raise ToolError(f"Unsupported image type; allowed: {sorted(IMAGE_EXTS)}")
    # AddPicture(Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
    shape = ws.Shapes.AddPicture(full, False, True, float(anchor.Left), float(anchor.Top), -1.0, -1.0)
    if width:
        shape.LockAspectRatio = True
        shape.Width = float(width)
    if name:
        shape.Name = name
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "shape": shape.Name, "width": float(shape.Width), "height": float(shape.Height)}


# ================================================================== вид листа и печать


@office_tool("excel_format", "write", title="Sheet view settings")
def excel_sheet_view(
    workbook: str,
    sheet: str = "",
    freeze_at: str | None = None,
    zoom: int | None = None,
    gridlines: bool | None = None,
    headings: bool | None = None,
) -> dict:
    """Change how a sheet looks on screen: freeze panes, zoom, gridlines, row/column headings.

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        freeze_at: cell below/right of the frozen area, e.g. 'A2' freezes the first row, 'B2' the first row and column; '' or 'none' unfreezes.
        zoom: percent, 10..400.
        gridlines, headings: show/hide.
    """
    app, wb = pick_workbook(workbook)
    ws = pick_sheet(wb, sheet)
    prev = _remember_view(app)
    applied = []
    try:
        wb.Activate()
        ws.Activate()
        win = app.ActiveWindow
        if freeze_at is not None:
            win.FreezePanes = False
            if freeze_at.strip().lower() not in ("", "none"):
                r, c, _, _ = parse_a1(freeze_at)
                _freeze_panes(win, r, c)
            applied.append(f"freeze_at={freeze_at or 'none'}")
        if zoom is not None:
            win.Zoom = max(10, min(400, int(zoom)))
            applied.append(f"zoom={zoom}")
        if gridlines is not None:
            win.DisplayGridlines = bool(gridlines)
            applied.append(f"gridlines={gridlines}")
        if headings is not None:
            win.DisplayHeadings = bool(headings)
            applied.append(f"headings={headings}")
    finally:
        _restore_view(app, prev)
    if not applied:
        raise ToolError("Nothing to change: pass freeze_at, zoom, gridlines or headings.")
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "applied": applied}


def _freeze_panes(win, row, column):
    win.FreezePanes = False
    win.ScrollRow = win.ScrollColumn = 1
    win.SplitColumn, win.SplitRow = column - 1, row - 1
    win.FreezePanes = True


@office_tool("excel_format", "write", title="Page setup")
def excel_page_setup(
    workbook: str,
    sheet: str = "",
    orientation: str | None = None,
    paper: str | None = None,
    fit_to_width: int | None = None,
    fit_to_height: int | None = None,
    print_area: str | None = None,
    title_rows: str | None = None,
    margins_cm: float | None = None,
    center_footer: str | None = None,
) -> dict:
    """Print settings of a sheet (the needed printer driver must be available to Windows).

    Args:
        workbook: exact workbook name.
        sheet: sheet name ('' = active).
        orientation: 'portrait' or 'landscape'. paper: 'A4', 'A3', 'A5' or 'Letter'.
        fit_to_width, fit_to_height: scale to this many pages wide/tall (0 = automatic).
        print_area: e.g. 'A1:H50' ('' clears it). title_rows: rows repeated on each page, e.g. '1:1'.
        margins_cm: all four margins in centimetres. center_footer: footer text ('&P / &N' = page x of y).
    """
    app, wb = pick_workbook(workbook)
    ws = pick_sheet(wb, sheet)
    applied = []
    try:
        ps = ws.PageSetup
        if orientation:
            o = orientation.lower()
            if o not in ("portrait", "landscape"):
                raise ToolError("orientation must be 'portrait' or 'landscape'.")
            ps.Orientation = 1 if o == "portrait" else 2
            applied.append(f"orientation={o}")
        if paper:
            codes = {"a4": 9, "a3": 8, "a5": 11, "letter": 1}
            if paper.lower() not in codes:
                raise ToolError("paper must be A4, A3, A5 or Letter.")
            ps.PaperSize = codes[paper.lower()]
            applied.append(f"paper={paper}")
        if fit_to_width is not None or fit_to_height is not None:
            ps.Zoom = False
            ps.FitToPagesWide = int(fit_to_width) if fit_to_width else False
            ps.FitToPagesTall = int(fit_to_height) if fit_to_height else False
            applied.append(f"fit={fit_to_width}x{fit_to_height}")
        if print_area is not None:
            ps.PrintArea = print_area
            applied.append(f"print_area={print_area}")
        if title_rows is not None:
            ps.PrintTitleRows = ("$" + title_rows.replace(":", ":$")) if title_rows and "$" not in title_rows else title_rows
            applied.append(f"title_rows={title_rows}")
        if margins_cm is not None:
            pts = cm_to_points(margins_cm)
            ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = pts
            applied.append(f"margins_cm={margins_cm}")
        if center_footer is not None:
            ps.CenterFooter = center_footer
            applied.append("center_footer")
    except pywintypes.com_error as exc:
        raise ToolError("Excel could not change the page setup (is a printer driver installed?): " + com.com_error_text(exc)) from None
    if not applied:
        raise ToolError("Nothing to change.")
    return {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "applied": applied}


# ================================================================== снимок диапазона


def _is_blank_png(data: bytes) -> bool:
    """Картинка почти однотонная (экранный CopyPicture вне видимой области/в свёрнутом окне даёт белый прямоугольник с рамкой)."""
    try:
        from PIL import Image as PILImage

        im = PILImage.open(io.BytesIO(data)).convert("L")
        if im.width > 12 and im.height > 12:
            im = im.crop((4, 4, im.width - 4, im.height - 4))  # серая рамка по краю есть и у пустого снимка
        hist = im.histogram()
        return max(hist) / float(im.width * im.height) > 0.998
    except Exception:  # noqa: BLE001 — нет Pillow/битый файл: считаем картинку годной
        return len(data) < 600


def _copy_picture_png(app, wb, ws, rng, appearance: int) -> bytes:
    """CopyPicture -> временная диаграмма-контейнер на листе -> Export PNG. appearance: 1 = как на экране, 2 = как при печати.

    Единственный проверенный способ: вставка в диаграмму ДРУГОЙ книги даёт пустую картинку, а внешнее чтение буфера обмена
    (ImageGrab) роняло Excel (исключение 0xC015000F). Поэтому на долю секунды на листе появляется диаграмма, которая тут же удаляется.
    Флаг «сохранено» НЕ возвращаем: удаление диаграммы не доказывает, что рядом не появилось настоящее изменение (пользователь,
    надстройка), а скрыть его значило бы потерять правку при закрытии без вопроса. Цена — книга после снимка помечена как изменённая.
    """
    tmpdir = tempfile.mkdtemp(prefix="office_live_")
    png = os.path.join(tmpdir, "range.png")
    holder, holder_name, removed = None, "(unnamed)", True
    data = b""
    try:
        rng.CopyPicture(appearance, 2)  # Appearance, Format=xlBitmap
        holder = ws.ChartObjects().Add(float(rng.Left), float(rng.Top), float(rng.Width), float(rng.Height))
        holder_name = str(holder.Name)
        holder.Chart.Paste()
        holder.Chart.Export(png, "PNG")
        with open(png, "rb") as f:
            data = f.read()
    finally:
        if holder is not None:
            try:
                holder.Delete()
            except pywintypes.com_error:
                removed = False
        try:
            app.CutCopyMode = False
        except pywintypes.com_error:
            pass
        shutil.rmtree(tmpdir, ignore_errors=True)
    if not removed:
        raise ToolError(f"The snapshot was taken, but the temporary chart '{holder_name}' could not be removed from sheet '{ws.Name}'. Delete it manually.")
    return data


def _pause_autosave(wb) -> bool:
    """True, если автосохранение было включено и мы его приостановили (вернуть после снимка)."""
    try:
        on = bool(wb.AutoSaveOn)
    except (pywintypes.com_error, AttributeError):
        return False
    if not on:
        return False
    try:
        wb.AutoSaveOn = False
    except pywintypes.com_error:
        raise ToolError(
            "AutoSave is ON for this workbook and could not be paused: a snapshot places a temporary chart on the sheet, which would be synced to the cloud. "
            "Turn AutoSave off and retry."
        ) from None
    return True


def render_range_png(app, wb, ws, rng, max_cells: int = 1500) -> bytes:
    """Диапазон «как на экране» -> PNG. Кратко занимает буфер обмена и на долю секунды ставит на лист временную диаграмму (она удаляется; содержимое книги не меняется).

    Экранный CopyPicture рисует только видимое: лист должен быть активным и диапазон прокручен в видимую область —
    поэтому временно активируем книгу/лист и прокручиваем (состояние пользователя потом возвращается). Если снимок
    подозрительно пуст (окно свёрнуто и т. п.), повторяем в режиме «как при печати»; если и он пуст при непустых
    данных — честная ошибка вместо белого прямоугольника.
    """
    r1, c1, r2, c2 = bounds(rng)
    cells = (r2 - r1 + 1) * (c2 - c1 + 1)
    if cells > max_cells:
        raise ToolError(f"Range is larger than max_cells={max_cells}; render a smaller area.")
    if float(rng.Width) < 1 or float(rng.Height) < 1:
        raise ToolError("The range is hidden or has zero size.")
    has_content = count_nonempty(app, rng) != 0
    suspend_events(app)  # activation of sheets must not run the user's SheetActivate macros
    resume_autosave = _pause_autosave(wb)
    prev = None
    restore_states = []
    data, blank, last_error = b"", True, ""
    try:  # всё после паузы AutoSave — внутри try: любая ошибка (даже при запоминании вида) должна его вернуть
        prev = _remember_view(app)
        # свёрнутое или маленькое окно (приложения/книги) даёт пустой снимок — на время разворачиваем на весь экран,
        # после снимка возвращаем как было
        try:
            if int(app.WindowState) == -4140:
                restore_states.append((app, int(app.WindowState)))
                app.WindowState = -4143  # xlNormal
            win = wb.Windows(1)
            if int(win.WindowState) != -4137:
                restore_states.append((win, int(win.WindowState)))
                win.WindowState = -4137  # xlMaximized
        except pywintypes.com_error:
            pass
        try:
            wb.Activate()
            ws.Activate()
            app.Goto(rng.Cells(1, 1), True)  # прокрутить к началу диапазона
        except pywintypes.com_error:
            pass  # окно скрыто — сработает запасной режим ниже
        for attempt in range(3):
            time.sleep(0.4 * (attempt + 1))  # дать окну перерисоваться после смены состояния/прокрутки
            try:
                data = _copy_picture_png(app, wb, ws, rng, 1)
                blank = _is_blank_png(data)
            except pywintypes.com_error as exc:
                data, blank, last_error = b"", True, com.com_error_text(exc)
            if not blank:
                break
        if blank:
            try:
                fallback = _copy_picture_png(app, wb, ws, rng, 2)  # «как при печати» (нужен принтер по умолчанию)
                if fallback and not _is_blank_png(fallback):
                    data, blank = fallback, False
            except pywintypes.com_error as exc:
                last_error = last_error or com.com_error_text(exc)
    finally:
        try:
            for obj, state in restore_states:
                try:
                    obj.WindowState = state
                except pywintypes.com_error:
                    pass
            if prev is not None:
                _restore_view(app, prev)
        finally:  # окна могут не вернуться («Office занят»), но AutoSave возвращаем в любом случае
            if resume_autosave:
                try:
                    wb.AutoSaveOn = True
                except Exception:  # не заслонять исходную ошибку, но и не молчать
                    print(f"[office-live] could not turn AutoSave back on for {wb.Name}; ask the user to re-enable it", file=sys.stderr)
    if not data:
        raise ToolError("Could not render the range (the workbook window must be visible and not minimized)" + (f": {last_error}" if last_error else "."))
    if blank and has_content:
        raise ToolError("The rendered image is blank although the range has data: the workbook window is probably hidden or minimized. Make it visible and retry.")
    return data


@office_tool("excel_format", "read", title="Render range as image", unstructured=True, read_only=False, file_args=("save_to",))
def excel_render_range_image(workbook: str = "", sheet: str = "", cells: str = "", max_cells: int = 1500,
                             save_to: str = "", overwrite: bool = False) -> list:
    """Render a range exactly as it looks on screen (fonts, fills, borders, conditional formats) and return it as a PNG image - use it to visually verify formatting. Briefly uses the Windows clipboard and a temporary chart object on the sheet (removed at once; the workbook content is unchanged, but Excel marks the workbook as modified and the Undo history is cleared; AutoSave is paused meanwhile).

    Args:
        workbook: exact name or '' for the active workbook.
        sheet: sheet name or '' for active.
        cells: range to render ('' = the used range); at most max_cells cells.
        max_cells: safety cap on range size.
        save_to: optional PNG file path within OFFICE_LIVE_ALLOWED_DIRS; creates parent folders. Returns only metadata
            (absolute path, png_bytes, size_px=[width, height]) instead of an inline image. Not available in read-only mode.
        overwrite: allow save_to to replace an existing file (default false).
    """
    dest = None
    if save_to:
        if config.SETTINGS.readonly:
            raise ToolError("save_to writes a file, which read-only mode does not allow; omit it to just get the image back.")
        dest = check_path(save_to, "write", {"png"})
        if os.path.exists(dest) and not overwrite:
            raise ToolError(f"File already exists: {dest}. Pass overwrite=true to replace it.")
    app, wb = pick_workbook(workbook)
    ws, rng = get_range(wb, sheet, cells)
    data = render_range_png(app, wb, ws, rng, max_cells)
    info = {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "cells": addr_of(rng), "png_bytes": len(data)}
    if dest:
        from PIL import Image as PILImage

        with PILImage.open(io.BytesIO(data)) as image:
            size_px = list(image.size)
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        try:
            with open(dest, "wb" if overwrite else "xb") as output:
                output.write(data)
        except FileExistsError:
            raise ToolError(f"File already exists: {dest}. Pass overwrite=true to replace it.") from None
        info.update(path=dest, size_px=size_px)
        return [json.dumps(info, ensure_ascii=False)]
    return [json.dumps(info, ensure_ascii=False), Image(data=data, format="png")]


SPARK_TYPES = {"line": 1, "column": 2, "win_loss": 3}
SPARK_POINTS = {"high": "Highpoint", "low": "Lowpoint", "first": "Firstpoint", "last": "Lastpoint", "negative": "Negative", "all": "Markers"}
SPARK_PROPERTIES = ("DateRange", "DisplayBlanksAs", "DisplayHidden", "LineWeight", "PlotBy",
                    "Axes.Horizontal.RightToLeftPlotOrder", "Axes.Vertical.MinScaleType", "Axes.Vertical.MaxScaleType",
                    "Axes.Vertical.CustomMinScaleValue", "Axes.Vertical.CustomMaxScaleValue")
SPARK_COLORS = ("SeriesColor", "Axes.Horizontal.Axis", *("Points." + p for p in SPARK_POINTS.values()))


def _spark_read(obj, member):
    try:
        return getattr(obj, member)
    except pywintypes.com_error as exc:
        if com.is_busy(exc) or com.is_dead(exc):
            raise
        if member == "DateRange":
            return ""  # live probe: reading DateRange raises until a date axis is set
        raise


def spark_state(group, full=False):
    # Live probe (Excel 16): SparklineGroup.SeriesColor IS a FormatColor; Points.* and Axes.Horizontal.Axis have .Color.
    result = {"location": xl.addr_of(group.Location), "source": group.SourceData, "type": int(group.Type),
              "color": int(group.SeriesColor.Color),
              "markers": [k for k, p in SPARK_POINTS.items() if bool(getattr(group.Points, p).Visible)]}
    if full:
        result["properties"] = {}
        for p in SPARK_PROPERTIES:
            obj, member = undo._property_object(group, p)
            result["properties"][p] = _spark_read(obj, member)
        for p in SPARK_COLORS:
            obj, member = undo._property_object(group, p)
            owner = getattr(obj, member)
            fc, prefix = (owner, p) if p == "SeriesColor" else (owner.Color, p + ".Color")
            try:
                theme = int(fc.ThemeColor)
            except pywintypes.com_error as exc:
                if com.is_busy(exc) or com.is_dead(exc):
                    raise
                theme = 0  # FormatColor has no ThemeColor for explicit RGB (same contract as undo._property_state).
            result["properties"][prefix + (".ThemeColor" if 1 <= theme <= 12 else ".Color")] = theme if 1 <= theme <= 12 else int(fc.Color)
            result["properties"][prefix + ".TintAndShade"] = fc.TintAndShade
            if p != "SeriesColor":
                result["properties"][p + ".Visible"] = bool(owner.Visible)
    return result


def spark_groups(ws):
    coll = ws.Range("A1:XFD1048576").SparklineGroups
    return [coll.Item(i) for i in range(1, int(coll.Count) + 1)]


@office_tool("excel_format", "write", title="Sparklines", read_actions=("list",))
def excel_sparklines(workbook: str = "", sheet: str = "", action: str = "list", data: str = "", location: str = "",
                     type: str = "line", color: str = "", markers: list[str] | None = None) -> dict:
    """List, add or clear sparkline groups. list works in readonly. Add undo deletes the exact created group; clear restores readable properties or reports a barrier.

    Args:
        workbook, sheet: selected workbook and location sheet.
        action: list/add/clear.
        data: source rectangle, optionally prefixed by a sheet name in the same workbook.
        location: one row or column; clear requires whole groups, never a partial group.
        type: line/column/win_loss.
        color: optional series color; markers: high/low/first/last/negative/all.
    """
    if action not in {"list", "add", "clear"} or type not in SPARK_TYPES:
        raise ToolError("Invalid sparkline action or type.")
    if markers is not None and (not isinstance(markers, list) or any(m not in SPARK_POINTS for m in markers)):
        raise ToolError("Invalid sparkline markers.")
    rgb = parse_color(color) if color else None
    app, wb = xl.pick_workbook(workbook)
    ws = xl.pick_sheet(wb, sheet)
    if action == "list":
        return {"workbook": wb.Name, "sheet": ws.Name, "groups": [spark_state(g) for g in spark_groups(ws)]}
    ws, dest = xl.get_range(wb, ws.Name, location, empty_means_used=False)
    xl.validate_rectangle(dest, ws, wb)
    if int(dest.Count) > config.SETTINGS.undo_max_cells:
        raise ToolError("Sparkline location exceeds OFFICE_LIVE_UNDO_MAX_CELLS.")
    if bool(ws.ProtectContents):
        raise ToolError("The worksheet is protected.")
    if dest.MergeCells is not False:
        raise ToolError("Sparkline location must not contain merged cells.")
    affected = [g for g in spark_groups(ws) if app.Intersect(dest, g.Location) is not None]
    if action == "add":
        if affected:
            raise ToolError("location already contains sparklines.")
        source_ws, source = xl.get_range(wb, ws.Name, data, empty_means_used=False, allow_other_sheet=True)
        xl.validate_rectangle(source, source_ws, wb)
        dr, dc = int(dest.Rows.Count), int(dest.Columns.Count)
        if dr != 1 and dc != 1:
            raise ToolError("location must be one row or one column.")
        sr, sc = int(source.Rows.Count), int(source.Columns.Count)
        plot = 1 if dc == 1 and (dr > 1 or sr == 1) else 2
        if (plot == 1 and dr != sr) or (plot == 2 and dc != sc):
            raise ToolError("data and location dimensions do not match.")
        source_ref = xl.ref_label(source_ws, source)
        content = sum(v is not None for row in to_grid(dest.Formula) for v in row)
        undo.require_undo("workbook", app, wb, plan={"_spark_sheet": ws.Name, "_spark_location": xl.addr_of(dest), "_spark_before": []})
        try:
            group = dest.SparklineGroups.Add(SPARK_TYPES[type], source_ref)
            if sr == sc:  # live probe: PlotBy can only be set for square data; otherwise Excel orients itself (1004)
                group.PlotBy = plot  # xlSparklineRowsSquare / xlSparklineColumnsSquare
            if rgb is not None:
                group.SeriesColor.Color = rgb
            if markers is not None:
                for key, member in SPARK_POINTS.items():
                    getattr(group.Points, member).Visible = key in markers
            return {"workbook": wb.Name, "sheet": ws.Name, "groups": [spark_state(group)], "cells_with_content": content}
        except (pywintypes.com_error, ToolError) as exc:
            try:
                applied = int(dest.SparklineGroups.Count)
            except (pywintypes.com_error, ToolError):
                applied = "unknown"
            raise PartialChangeError(f"Sparkline add interrupted: applied={applied} groups. {com.translate(exc)}") from None
    for group in affected:
        if int(app.Intersect(dest, group.Location).Count) != int(group.Location.Count):
            raise ToolError("clear must include whole sparkline groups; partial locations are refused.")
    before, barrier = [], ""
    try:
        before = [spark_state(g, full=True) for g in affected]
    except (AttributeError, pywintypes.com_error) as exc:
        if com.is_busy(exc) or com.is_dead(exc):
            raise
        barrier = "Sparkline properties cannot all be read; clear is an undo barrier."
    if not affected:
        return {"workbook": wb.Name, "sheet": ws.Name, "cleared_groups": 0}
    undo.require_undo("workbook", app, wb, plan={"_spark_sheet": ws.Name, "_spark_location": xl.addr_of(dest),
                                                "_spark_before": before, "_operation_barrier": barrier}, barrier_reason=barrier)
    applied = 0
    try:
        for group in affected:
            group.Delete()
            applied += 1
    except pywintypes.com_error as exc:
        raise PartialChangeError(f"Sparkline clear interrupted: applied={applied} groups. {com.translate(exc)}") from None
    return {"workbook": wb.Name, "sheet": ws.Name, "cleared_groups": applied}
