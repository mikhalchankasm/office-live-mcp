"""Excel chart creation, styling and export."""

import contextlib
import json
import os
import shutil
import tempfile
import time
import pywintypes
from mcp.server.mcpserver import Image
from . import config
from .errors import ToolError
from .excel_format import NUMBER_FORMATS
from .registry import office_tool
from .safety import check_path
from .util import a1_cell, parse_color
from .xl_common import (
    addr_of,
    bounds,
    get_range,
    number_format_for_write,
    pick_sheet,
    pick_workbook,
)


CHART_TYPES = {
    "column": 51, "column_stacked": 52, "column_100": 53, "bar": 57, "bar_stacked": 58, "bar_100": 59,
    "line": 4, "line_markers": 65, "line_stacked": 63, "pie": 5, "pie_exploded": 69, "doughnut": -4120,
    "area": 1, "area_stacked": 76, "scatter": -4169, "scatter_lines": 74, "scatter_smooth": 72,
    "radar": -4151, "radar_markers": 81, "bubble": 15, "surface": 83,
}
LEGEND_POS = {"bottom": -4107, "top": -4160, "left": -4131, "right": -4152}
NO_AXES = {"pie", "pie_exploded", "doughnut"}


def _chart_type_code(chart_type: str) -> int:
    t = chart_type.lower().replace("-", "_").replace(" ", "_")
    if t not in CHART_TYPES:
        raise ToolError(f"chart_type must be one of {sorted(CHART_TYPES)}")
    return CHART_TYPES[t]


def _find_chart(wb, chart_name: str, sheet: str = ""):
    sheets = [pick_sheet(wb, sheet)] if sheet else [wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1)]
    seen = []
    for ws in sheets:
        cos = ws.ChartObjects()
        for j in range(1, int(cos.Count) + 1):
            co = cos(j)
            seen.append(co.Name)
            if co.Name.lower() == chart_name.lower():
                return ws, co
    raise ToolError(f"Chart '{chart_name}' not found. Charts: {seen}")


def _chart_summary(co) -> dict:
    ch = co.Chart
    info = {"name": co.Name, "chart_type": int(ch.ChartType), "left": float(co.Left), "top": float(co.Top), "width": float(co.Width), "height": float(co.Height)}
    info["title"] = ch.ChartTitle.Text if ch.HasTitle else None
    series = []
    sc = ch.SeriesCollection()
    for k in range(1, int(sc.Count) + 1):
        s = sc(k)
        series.append({"index": k, "name": s.Name, "formula": s.Formula})
    info["series"] = series
    return info


def _style_chart(ch, kind_key: str, title, x_title, y_title, legend, data_labels, y_format=None, y_min=None, y_max=None, colors=None, label_format=None):
    if title is not None:
        if title == "":
            ch.HasTitle = False
        else:
            ch.HasTitle = True
            ch.ChartTitle.Text = title
    if kind_key not in NO_AXES:
        if x_title is not None:
            ax = ch.Axes(1)
            ax.HasTitle = bool(x_title)
            if x_title:
                ax.AxisTitle.Text = x_title
        if y_title is not None:
            ax = ch.Axes(2)
            ax.HasTitle = bool(y_title)
            if y_title:
                ax.AxisTitle.Text = y_title
    if legend is not None:
        if legend.lower() == "none":
            ch.HasLegend = False
        else:
            if legend.lower() not in LEGEND_POS:
                raise ToolError("legend must be bottom|top|left|right|none")
            ch.HasLegend = True
            ch.Legend.Position = LEGEND_POS[legend.lower()]
    if data_labels is not None:
        sc = ch.SeriesCollection()
        for k in range(1, int(sc.Count) + 1):
            sc(k).HasDataLabels = bool(data_labels)
    app = ch.Application
    if kind_key not in NO_AXES and (y_format or y_min is not None or y_max is not None):
        ax = ch.Axes(2)
        if y_format:
            # формат диаграммы хранится как есть, без учёта LCID: буквы дат английские, а разделители местные (живая проверка)
            ax.TickLabels.NumberFormat = number_format_for_write(app, y_format, NUMBER_FORMATS)
        if y_min is not None:
            ax.MinimumScale = float(y_min)
        if y_max is not None:
            ax.MaximumScale = float(y_max)
    if label_format:
        sc = ch.SeriesCollection()
        for k in range(1, int(sc.Count) + 1):
            try:
                sc(k).DataLabels().NumberFormat = number_format_for_write(app, label_format, NUMBER_FORMATS)
            except pywintypes.com_error:
                pass  # у ряда нет подписей данных
    if colors:
        sc = ch.SeriesCollection()
        for k, col in enumerate(colors, start=1):
            if k > int(sc.Count):
                break
            bgr = parse_color(col)
            s_k = sc(k)
            try:
                s_k.Format.Fill.ForeColor.RGB = bgr
            except pywintypes.com_error:
                pass
            try:
                s_k.Format.Line.ForeColor.RGB = bgr
            except pywintypes.com_error:
                pass


def _chart_pivot(ch):
    try:
        layout = ch.PivotLayout
        return None if layout is None else layout.PivotTable.Name
    except pywintypes.com_error:
        return None  # обычная диаграмма: у неё нет PivotLayout


@contextlib.contextmanager
def _chart_selection(app, wb, ws, src):
    """AddChart2 привязывает новую диаграмму к данным вокруг АКТИВНОЙ ячейки, в том числе к другой сводной таблице;
    SetSourceData затем падает с 0x80004005 (демо 2026-10-05). На время создания выделяем левую верхнюю ячейку источника."""
    from .excel_cf import _remember_view, _restore_view

    pts = ws.PivotTables()
    if not int(pts.Count):
        yield  # без сводных на листе автопривязка безвредна: SetSourceData задаёт данные явно
        return
    prev = _remember_view(app)
    try:
        wb.Activate()
        ws.Activate()
        src.Cells(1, 1).Select()
        yield
    finally:
        _restore_view(app, prev)


@office_tool("excel_analysis", "write", title="Create chart")
def excel_create_chart(
    workbook: str,
    source: str,
    sheet: str = "",
    chart_type: str = "column",
    title: str = "",
    anchor_cell: str = "",
    width: float = 480,
    height: float = 288,
    x_axis_title: str = "",
    y_axis_title: str = "",
    legend: str = "bottom",
    data_labels: bool = False,
    series_in_rows: bool | None = None,
    name: str = "",
    value_axis_number_format: str = "",
    value_axis_min: float | None = None,
    value_axis_max: float | None = None,
    series_colors: list[str] | None = None,
    data_label_number_format: str = "",
) -> dict:
    """Create a chart from a data block. Include the header row and the label column in `source`: the first column becomes the category axis and each other column a series (use series_in_rows=true if the series are laid out in rows).

    Args:
        workbook: exact workbook name.
        source: data range such as 'A1:D6' (or 'Sheet2!A1:D6'), or the name of a pivot table to get a pivot chart.
        sheet: sheet of the source and where the chart is placed ('' = active).
        chart_type: column | column_stacked | column_100 | bar | bar_stacked | line | line_markers | pie | doughnut | area | area_stacked | scatter | scatter_lines | scatter_smooth | radar | bubble.
        title: chart title ('' = no title).
        anchor_cell: top-left cell for the chart (default: two columns right of the data).
        width, height: size in points.
        x_axis_title, y_axis_title: axis captions ('' = none).
        legend: bottom | top | left | right | none.
        data_labels: show value labels.
        series_in_rows: force series orientation (null = automatic).
        name: chart object name.
        value_axis_number_format: number format of the value axis, e.g. '#,##0' (no decimals). value_axis_min, value_axis_max: axis limits.
        series_colors: one color per series, e.g. ['#1F4E78', '#C00000']. data_label_number_format: format of the value labels.
    """
    app, wb = pick_workbook(workbook)
    expanded = None  # диапазон внутри сводной расширяется до всей сводной — сообщаем об этом в ответе
    pivot_hit = None
    for pws in [wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1)]:
        pts = pws.PivotTables()
        for j in range(1, int(pts.Count) + 1):
            if pts(j).Name.lower() == source.strip().lower():
                pivot_hit = (pws, pts(j))
    if pivot_hit:  # источник — сводная таблица: получится сводная диаграмма
        ws, src = pivot_hit[0], pivot_hit[1].TableRange1
    else:
        ws, src = get_range(wb, sheet, source, empty_means_used=False)
        pts = ws.PivotTables()
        inside = [pts(j) for j in range(1, int(pts.Count) + 1) if app.Intersect(src, pts(j).TableRange2) is not None]
        if len(inside) > 1:
            raise ToolError(f"source {addr_of(src)} overlaps several pivot tables ({', '.join(p.Name for p in inside)}); pass one pivot name.")
        if inside:  # диапазон внутри сводной — та же сводная диаграмма по всей сводной
            expanded = {"requested": addr_of(src), "pivot": inside[0].Name, "note": "Excel builds a pivot chart from the whole pivot table"}
            pivot_hit, src = (ws, inside[0]), inside[0].TableRange1
    r1, c1, r2, c2 = bounds(src)
    key = chart_type.lower().replace("-", "_").replace(" ", "_")
    code = _chart_type_code(key)
    anchor = ws.Range(anchor_cell) if anchor_cell else ws.Range(a1_cell(r1, min(c2 + 2, 16384)))
    left, top = float(anchor.Left), float(anchor.Top)
    target = pivot_hit[1].Name if pivot_hit else None
    shape = None
    try:
        with _chart_selection(app, wb, ws, src):
            try:
                shape = ws.Shapes.AddChart2(-1, code, left, top, float(width), float(height))  # AddChart2(Style, Type, Left, Top, W, H)
                ch = shape.Chart
            except (pywintypes.com_error, AttributeError):
                shape = ws.ChartObjects().Add(left, top, float(width), float(height))
                ch = shape.Chart
                ch.ChartType = code
        if series_in_rows is None:
            ch.SetSourceData(src)
        else:
            ch.SetSourceData(src, 1 if series_in_rows else 2)  # PlotBy: xlRows=1, xlColumns=2
        bound = _chart_pivot(ch)
        if bound != target:
            raise ToolError(f"Excel bound the new chart to pivot '{bound}' instead of " + (f"pivot '{target}'." if target else f"range {addr_of(src)}."))
        if name:
            shape.Name = name
        _style_chart(ch, key, title, x_axis_title, y_axis_title, legend, data_labels, value_axis_number_format or None, value_axis_min, value_axis_max, series_colors, data_label_number_format or None)
    except Exception:
        if shape is not None:  # недоделанная диаграмма не остаётся на листе
            with contextlib.suppress(pywintypes.com_error):
                shape.Delete()
        raise
    series = []
    sc = ch.SeriesCollection()
    for k in range(1, int(sc.Count) + 1):
        series.append(sc(k).Name)
    out = {"ok": True, "workbook": wb.Name, "sheet": ws.Name, "chart": shape.Name, "type": key, "series": series, "position": {"left": left, "top": top, "width": float(width), "height": float(height)}}
    if expanded:
        out["source_expanded_to_pivot"] = expanded
    return out


_PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"


def _export_chart_png(chart, png: str, name: str) -> bytes:
    """Chart.Export sometimes writes an empty file (live finding #32: long-lived test Excel). Retry once after
    Chart.Refresh; never report success for a file that is not a PNG."""
    for attempt in range(2):
        if attempt:
            chart.Refresh()
            time.sleep(0.3)
        chart.Export(png, "PNG")
        with open(png, "rb") as f:
            data = f.read()
        if data.startswith(_PNG_SIGNATURE):
            return data
    raise ToolError(f"Excel returned an empty image for chart '{name}' twice. Make sure the workbook window is not "
                    "minimized or hidden (office_window) and try again; nothing was saved.")


@office_tool("excel_analysis", "write", title="Manage charts", unstructured=True, read_actions=("list", "export_image"), destructive=True, file_args=("export_path",))
def excel_manage_charts(
    workbook: str,
    action: str = "list",
    chart_name: str = "",
    sheet: str = "",
    chart_type: str = "",
    title: str | None = None,
    x_axis_title: str | None = None,
    y_axis_title: str | None = None,
    legend: str | None = None,
    data_labels: bool | None = None,
    source: str = "",
    series_overrides: list[dict] | None = None,
    new_series: dict | None = None,
    width: float | None = None,
    height: float | None = None,
    anchor_cell: str = "",
    export_path: str = "",
    overwrite: bool = False,
    value_axis_number_format: str | None = None,
    value_axis_min: float | None = None,
    value_axis_max: float | None = None,
    series_colors: list[str] | None = None,
    data_label_number_format: str | None = None,
) -> list | dict:
    """List, restyle, move, delete or export charts.

    Args:
        workbook: exact workbook name.
        action: 'list' | 'update' | 'delete' | 'export_image' (returns the chart as a PNG image; also saved to export_path when given).
        chart_name: chart name from 'list' (required except for list).
        sheet: sheet to search ('' = all sheets).
        chart_type, title, x_axis_title, y_axis_title, legend, data_labels: new settings for 'update' ('' title removes it).
        source: new data range for 'update'.
        series_overrides: combo-chart control, e.g. [{"index": 2, "chart_type": "line", "secondary_axis": true}].
        new_series: add a series, e.g. {"name": "Plan", "values": "B2:B6", "x_values": "A2:A6"} (ranges on the chart's sheet or 'Sheet!A1:A5').
        width, height, anchor_cell: resize/move the chart.
        export_path: optional PNG path for export_image (refuses to replace an existing file unless overwrite=true; not available in read-only mode).
        overwrite: allow export_path to replace an existing file.
        value_axis_number_format ('#,##0'), value_axis_min, value_axis_max, series_colors (one per series), data_label_number_format: axis and look settings for 'update'.
    """
    app, wb = pick_workbook(workbook)
    act = action.lower()
    if act == "list":
        sheets = [pick_sheet(wb, sheet)] if sheet else [wb.Worksheets(i) for i in range(1, wb.Worksheets.Count + 1)]
        items = []
        for ws in sheets:
            cos = ws.ChartObjects()
            for j in range(1, int(cos.Count) + 1):
                s = _chart_summary(cos(j))
                s["sheet"] = ws.Name
                items.append(s)
        return {"workbook": wb.Name, "charts": items}
    if not chart_name:
        raise ToolError("'chart_name' is required.")
    ws, co = _find_chart(wb, chart_name, sheet)
    ch = co.Chart
    if act == "delete":
        co.Delete()
        return {"ok": True, "workbook": wb.Name, "deleted": chart_name}
    if act == "export_image":
        dest = None
        if export_path:
            if config.SETTINGS.readonly:
                raise ToolError("export_path writes a file, which read-only mode does not allow; omit it to just get the image back.")
            dest = check_path(export_path, "write", {"png"})
            if os.path.exists(dest) and not overwrite:
                raise ToolError(f"File already exists: {dest}. Pass overwrite=true to replace it.")
        tmpdir = tempfile.mkdtemp(prefix="office_live_")
        png = os.path.join(tmpdir, "chart.png")
        try:
            data = _export_chart_png(ch, png, co.Name)
            if dest:
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                shutil.copyfile(png, dest)
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
        return [json.dumps({"ok": True, "chart": co.Name, "png_bytes": len(data), "saved_to": export_path or None}, ensure_ascii=False), Image(data=data, format="png")]
    if act != "update":
        raise ToolError("action must be 'list', 'update', 'delete' or 'export_image'.")
    applied = []
    key = None
    if chart_type:
        key = chart_type.lower().replace("-", "_").replace(" ", "_")
        ch.ChartType = _chart_type_code(key)
        applied.append(f"chart_type={key}")
    if source:
        _, src = get_range(wb, ws.Name, source, empty_means_used=False, allow_other_sheet=True)
        ch.SetSourceData(src)
        applied.append(f"source={addr_of(src)}")
    _style_chart(ch, key or "", title, x_axis_title, y_axis_title, legend, data_labels, value_axis_number_format, value_axis_min, value_axis_max, series_colors, data_label_number_format)
    for k, v in (("title", title), ("x_axis_title", x_axis_title), ("y_axis_title", y_axis_title), ("legend", legend), ("data_labels", data_labels)):
        if v is not None:
            applied.append(f"{k}={v}")
    for k, v in (("value_axis_number_format", value_axis_number_format), ("value_axis_min", value_axis_min), ("value_axis_max", value_axis_max), ("series_colors", series_colors), ("data_label_number_format", data_label_number_format)):
        if v not in (None, ""):
            applied.append(f"{k}={v}")
    for ov in series_overrides or []:
        idx = int(ov.get("index", 0))
        s = ch.SeriesCollection(idx)
        if ov.get("chart_type"):
            s.ChartType = _chart_type_code(ov["chart_type"])
        if ov.get("secondary_axis") is not None:
            s.AxisGroup = 2 if ov["secondary_axis"] else 1
        applied.append(f"series{idx}")
    if new_series:
        s = ch.SeriesCollection().NewSeries()
        if new_series.get("name"):
            s.Name = new_series["name"]
        if new_series.get("values"):
            s.Values = get_range(wb, ws.Name, new_series["values"], empty_means_used=False, allow_other_sheet=True)[1]
        if new_series.get("x_values"):
            s.XValues = get_range(wb, ws.Name, new_series["x_values"], empty_means_used=False, allow_other_sheet=True)[1]
        applied.append("new_series")
    if width:
        co.Width = float(width)
        applied.append(f"width={width}")
    if height:
        co.Height = float(height)
        applied.append(f"height={height}")
    if anchor_cell:
        a = ws.Range(anchor_cell)
        co.Left, co.Top = float(a.Left), float(a.Top)
        applied.append(f"moved_to={anchor_cell}")
    if not applied:
        raise ToolError("Nothing to update.")
    return {"ok": True, "workbook": wb.Name, "chart": co.Name, "applied": applied, "now": _chart_summary(co)}
