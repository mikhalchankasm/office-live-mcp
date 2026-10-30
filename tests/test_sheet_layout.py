"""Issue #13: batch layout operations, using fake COM only."""

from functools import wraps

import pytest

from office_live import excel_sheets as excel_core, xl_common as xl
from office_live.errors import PartialChangeError, ToolError
from tests.history_fakes import Range, Sheet, fake_office  # noqa: F401
from tests.test_undo import revert, stack


def dimensions(o, **kwargs):
    return o.call("excel_set_dimensions", workbook=o.wb.Name, sheet="Data", **kwargs)


def formats(o, ranges):
    return o.call("excel_format_ranges", workbook=o.wb.Name, sheet="Data", ranges=ranges)


@pytest.fixture
def qualified_sheets(office):
    office.ws.Name = "Sheet1"
    target = office.wb.Worksheets.Add()
    target.Name = "Sheet2"
    office.ws.Activate()
    assert office.wb.ActiveSheet is office.ws
    return office.ws, target


@pytest.fixture
def cross_sheet_names(office, qualified_sheets, monkeypatch):
    _, target = qualified_sheets
    names = {}
    for name, address in (("LayoutTarget", "A1:C3"), ("LayoutOverlap", "B2:C3")):
        defined = office.wb.Names.Add(name, f"=Sheet2!{address}")
        defined.RefersToRange = target.Range(address)
        names[name] = defined
    original = Sheet.Range

    def resolve(self, address):
        if self.Parent is office.wb and address in names:
            return names[address].RefersToRange
        return original(self, address)

    monkeypatch.setattr(Sheet, "Range", resolve)


@pytest.mark.parametrize("cells", ["Sheet2!A1:C3", "LayoutTarget"])
@pytest.mark.parametrize("kwargs,widths,heights", [
    ({"column_width": 22}, [22, 22, 22], [15, 15, 15]),
    ({"row_height": 30}, [8.43, 8.43, 8.43], [30, 30, 30]),
    ({"column_widths": {"A": 22, "C": 24}}, [22, 8.43, 24], [15, 15, 15]),
    ({"row_heights": {"1": 30, "3": 32}}, [8.43, 8.43, 8.43], [30, 15, 32]),
    ({"autofit": "columns"}, [22, 22, 22], [15, 15, 15]),
    ({"autofit": "rows"}, [8.43, 8.43, 8.43], [30, 30, 30]),
    ({"autofit": "both"}, [22, 22, 22], [30, 30, 30]),
])
def test_dimensions_resolve_qualified_cells_with_blank_sheet(
    office, qualified_sheets, cross_sheet_names, monkeypatch, cells, kwargs, widths, heights,
):
    active, target = qualified_sheets
    calls, fit_sheets = [], []

    def effect(rng, axis):
        fit_sheets.append(rng.Worksheet)
        if axis == "Columns":
            rng.ColumnWidth = 22
        else:
            rng.RowHeight = 30

    autofit_spy(monkeypatch, calls, effect)
    result = office.call("excel_set_dimensions", workbook=office.wb.Name,
                         sheet="", cells=cells, **kwargs)
    assert result["sheet"] == "Sheet2" and result["undo"] == "available"
    assert [target.Columns(i).ColumnWidth for i in range(1, 4)] == widths
    assert [target.Rows(i).RowHeight for i in range(1, 4)] == heights
    if "autofit" in kwargs:
        mode = kwargs["autofit"]
        expected = []
        if mode in {"columns", "both"}:
            expected += [("Columns", f"{col}1:{col}3") for col in "ABC"]
        if mode in {"rows", "both"}:
            expected += [("Rows", f"A{row}:C{row}") for row in range(1, 4)]
        assert calls == expected
        assert all(ws is target for ws in fit_sheets)
    else:
        assert calls == []
    assert [active.Columns(i).ColumnWidth for i in range(1, 4)] == [active.StandardWidth] * 3
    assert [active.Rows(i).RowHeight for i in range(1, 4)] == [15] * 3
    assert len(stack(office)) == 1
    assert all(op["sheet"] == "Sheet2" for op in stack(office)[0].ops)
    revert(office)
    assert [target.Columns(i).ColumnWidth for i in range(1, 4)] == [target.StandardWidth] * 3
    assert [target.Rows(i).RowHeight for i in range(1, 4)] == [15] * 3
    assert [active.Columns(i).ColumnWidth for i in range(1, 4)] == [active.StandardWidth] * 3
    assert [active.Rows(i).RowHeight for i in range(1, 4)] == [15] * 3
    assert office.wb.ActiveSheet is active


@pytest.mark.parametrize("cells,overlap", [
    ("Sheet2!A1:C3", "Sheet2!B2:C3"), ("LayoutTarget", "LayoutOverlap"),
])
def test_format_ranges_resolve_qualified_cells_with_blank_sheet(
    office, qualified_sheets, cross_sheet_names, cells, overlap,
):
    active, target = qualified_sheets
    result = office.call("excel_format_ranges", workbook=office.wb.Name, sheet="", ranges=[
        {"cells": cells, "format": {"bold": True}},
        {"cells": overlap, "format": {"italic": True}},
    ])
    assert result["sheet"] == "Sheet2" and result["undo"] == "available"
    assert all(r["sheet"] == "Sheet2" for r in result["ranges"])
    assert target.Range("A1:C3").Font.Bold is True
    assert target.Range("B2:C3").Font.Italic is True
    assert active.Range("A1:C3").Font.Bold is False
    assert active.Range("A1:C3").Font.Italic is False
    assert len(stack(office)) == 1
    assert all(op["sheet"] == "Sheet2" for op in stack(office)[0].ops)
    revert(office)
    assert target.Range("A1:C3").Font.Bold is False
    assert target.Range("A1:C3").Font.Italic is False
    assert active.Range("A1:C3").Font.Bold is False
    assert active.Range("A1:C3").Font.Italic is False
    assert office.wb.ActiveSheet is active


@pytest.mark.parametrize("tool", ["excel_set_dimensions", "excel_format_ranges"])
def test_layout_still_rejects_explicit_conflicting_sheets(office, qualified_sheets, tool):
    active, target = qualified_sheets
    kwargs = ({"cells": "Sheet2!A1:C3", "column_width": 22}
              if tool == "excel_set_dimensions" else
              {"ranges": [{"cells": "Sheet2!A1:C3", "format": {"bold": True}}]})
    with pytest.raises(ToolError, match="Conflicting sheets"):
        office.call(tool, workbook=office.wb.Name, sheet="Sheet1", **kwargs)
    for ws in (active, target):
        assert ws.Range("A1:C3").Font.Bold is False
        assert [ws.Columns(i).ColumnWidth for i in range(1, 4)] == [ws.StandardWidth] * 3
    assert stack(office) == []


def test_ten_widths_and_row_heights_one_call_one_undo(office):
    o = office
    widths = {chr(65 + i): 12 + i for i in range(10)}
    before = [o.ws.Columns(i).ColumnWidth for i in range(1, 11)]
    result = dimensions(o, column_widths=widths, row_heights={"1": 24, "3": 32})
    assert result["undo"] == "available"
    assert [o.ws.Columns(i).ColumnWidth for i in range(1, 11)] == list(widths.values())
    assert o.ws.Rows(1).RowHeight == 24 and o.ws.Rows(3).RowHeight == 32
    assert len(stack(o)) == 1
    revert(o)
    assert [o.ws.Columns(i).ColumnWidth for i in range(1, 11)] == before
    assert o.ws.Rows(1).RowHeight == o.ws.Rows(3).RowHeight == 15


def test_four_formats_one_call_and_undo_restores_all(office):
    o = office
    o.ws.Range("A1:D4").Value = "preserved"
    o.ws.Range("A1:A2").Font.Bold = True  # Mixed before-state must also survive undo.
    specs = [
        {"cells": "A1:A4", "format": {"bold": False}},
        {"cells": "B1:B4", "format": {"italic": True}},
        {"cells": "C1:C4", "format": {"number_format": "0.00"}},
        {"cells": "D1:D4", "format": {"wrap_text": True}},
    ]
    result = formats(o, specs)
    assert result["undo"] == "available" and len(stack(o)) == 1
    assert o.ws.Range("A1:A4").Font.Bold is False
    assert o.ws.Range("B1:B4").Font.Italic is True
    assert xl.read_number_format(o.excel, o.ws.Range("C1:C4")) == "0.00"
    assert o.ws.Range("D1:D4").WrapText is True
    revert(o)
    assert o.ws.Range("A1:A2").Font.Bold is True
    assert o.ws.Range("A3:A4").Font.Bold is False
    assert o.ws.Range("B1:B4").Font.Italic is False
    assert o.ws.Range("C1:C4").NumberFormat == "General"
    assert o.ws.Range("D1:D4").WrapText is False
    assert o.ws.Range("A1:D4").Value == (("preserved",) * 4,) * 4


def autofit_spy(monkeypatch, calls, effect=None):
    original = Range.__init__

    def init(self, ws, address):
        original(self, ws, address)
        for axis in ("Columns", "Rows"):
            def fit(axis=axis):
                calls.append((axis, self.Address))
                if effect:
                    effect(self, axis)
            getattr(self, axis).AutoFit = fit

    monkeypatch.setattr(Range, "__init__", init)


@pytest.mark.parametrize("mode", ["columns", "rows", "both"])
def test_blank_autofit_only_touches_used_range_and_has_undo(office, monkeypatch, mode):
    o = office
    o.ws.Range("C7:L342").Value = "content"  # 336 by 10, offset from A1.
    calls = []
    def effect(rng, axis):
        if axis == "Columns":
            rng.ColumnWidth = 22
        else:
            rng.RowHeight = 30
    autofit_spy(monkeypatch, calls, effect)
    result = dimensions(o, autofit=mode)
    assert result["cells"] == "C7:L342" and result["undo"] == "available"
    assert calls
    for axis, address in calls:
        rng = o.ws.Range(address)
        assert 7 <= rng.Row <= rng.r2 <= 342
        assert 3 <= rng.Column <= rng.c2 <= 12
        assert axis == ("Columns" if mode == "columns" else "Rows") or mode == "both"
    assert o.ws.Columns(2).ColumnWidth == o.ws.StandardWidth
    assert o.ws.Rows(6).RowHeight == 15
    assert len(stack(o)) == 1
    revert(o)
    assert o.ws.Columns(3).ColumnWidth == o.ws.StandardWidth
    assert o.ws.Rows(7).RowHeight == 15


def test_autofit_refuses_oversized_used_range_before_com(office, monkeypatch):
    o = office
    o.ws.Range("A1").Value = "start"
    o.ws.Range("Z10000").Value = "end"
    calls = []
    autofit_spy(monkeypatch, calls)
    with pytest.raises(ToolError, match="[Aa]utofit.*limit"):
        dimensions(o, autofit="columns")
    assert calls == [] and stack(o) == []


@pytest.mark.parametrize("mode,strip_count", [("columns", 3), ("rows", 2), ("both", 5)])
def test_final_autofit_strip_crossing_deadline_succeeds(office, monkeypatch, mode, strip_count):
    o = office
    now = [0.0]
    monkeypatch.setattr(excel_core.time, "monotonic", lambda: now[0])
    calls = []

    def fit(rng, axis):
        if axis == "Columns":
            rng.ColumnWidth = 22
        else:
            rng.RowHeight = 30
        if len(calls) == strip_count:
            now[0] = 6.0

    autofit_spy(monkeypatch, calls, fit)
    result = dimensions(o, cells="A1:C2", autofit=mode)
    assert now[0] == 6.0
    assert result["ok"] is True and result["applied"] == [f"autofit={mode}"]
    assert result["undo"] == "available" and len(stack(o)) == 1
    expected = []
    if mode in {"columns", "both"}:
        expected += [("Columns", f"{col}1:{col}2") for col in "ABC"]
    if mode in {"rows", "both"}:
        expected += [("Rows", f"A{row}:C{row}") for row in range(1, 3)]
    assert calls == expected
    assert [o.ws.Columns(i).ColumnWidth for i in range(1, 4)] == (
        [22] * 3 if mode in {"columns", "both"} else [o.ws.StandardWidth] * 3
    )
    assert [o.ws.Rows(i).RowHeight for i in range(1, 3)] == (
        [30] * 2 if mode in {"rows", "both"} else [15] * 2
    )
    revert(o)
    assert [o.ws.Columns(i).ColumnWidth for i in range(1, 4)] == [o.ws.StandardWidth] * 3
    assert [o.ws.Rows(i).RowHeight for i in range(1, 3)] == [15] * 2


def test_autofit_deadline_stops_following_calls_and_preserves_undo(office, monkeypatch):
    o = office
    partial_errors = []
    original = excel_core.excel_set_dimensions

    @wraps(original)
    def observe_partial_error(*args, **kwargs):
        try:
            return original(*args, **kwargs)
        except PartialChangeError as exc:
            partial_errors.append(exc)
            raise

    # Observe the core error before the undo/registry layers translate it to ToolError.
    monkeypatch.setattr(excel_core, "excel_set_dimensions", observe_partial_error)
    o.ws.Range("A1:J336").Value = "content"
    now = [0.0]
    monkeypatch.setattr(excel_core.time, "monotonic", lambda: now[0])
    calls = []
    def slow(rng, axis):
        rng.ColumnWidth = 22
        now[0] += 6
    autofit_spy(monkeypatch, calls, slow)
    with pytest.raises(ToolError, match="[Aa]utofit.*time.*office_undo"):
        dimensions(o, autofit="both")
    assert len(partial_errors) == 1
    assert len(calls) == 1 and len(stack(o)) == 1
    assert o.ws.Columns(1).ColumnWidth == 22
    assert [o.ws.Columns(i).ColumnWidth for i in range(2, 11)] == [o.ws.StandardWidth] * 9
    assert [o.ws.Rows(i).RowHeight for i in range(1, 337)] == [15] * 336
    revert(o)
    assert o.ws.Columns(1).ColumnWidth == o.ws.StandardWidth


@pytest.mark.parametrize("kwargs", [
    {"column_widths": {"A": 12, "XFE": 15}},
    {"column_widths": {"A": 12, "B": -1}},
    {"row_heights": {"1": 24, "0": 10}},
    {"row_heights": {"1": 24, "2": float("nan")}},
    {"column_width": 12, "autofit": "invalid"},
])
def test_invalid_dimensions_change_nothing(office, kwargs):
    with pytest.raises(ToolError):
        dimensions(office, **kwargs)
    assert office.ws.Columns(1).ColumnWidth == office.ws.StandardWidth
    assert office.ws.Rows(1).RowHeight == 15
    assert stack(office) == []


@pytest.mark.parametrize("bad", [
    {"cells": "B1", "format": {"font_size": 0}},
    {"cells": "B1", "format": {"unknown": True}},
    {"cells": "", "format": {"italic": True}},
    {"cells": "B1", "format": {}},
])
def test_invalid_later_format_prevalidated(office, bad):
    with pytest.raises(ToolError):
        formats(office, [{"cells": "A1", "format": {"bold": True}}, bad])
    assert office.ws.Range("A1").Font.Bold is False
    assert stack(office) == []


def test_overlapping_formats_restore_original_state(office):
    formats(office, [
        {"cells": "A1:C1", "format": {"bold": True}},
        {"cells": "B1:D1", "format": {"bold": False, "italic": True}},
    ])
    assert office.ws.Range("A1").Font.Bold is True
    assert office.ws.Range("B1:D1").Font.Bold is False
    revert(office)
    assert office.ws.Range("A1:D1").Font.Bold is False
    assert office.ws.Range("A1:D1").Font.Italic is False


def test_partial_batch_failure_retains_one_undo(office, monkeypatch):
    from tests.history_fakes import RangeFormat
    import pywintypes
    original = RangeFormat.__setattr__
    def fail(self, name, value):
        if name == "Italic" and value is True:
            raise pywintypes.com_error(-2147352567, "simulated failure", None, None)
        original(self, name, value)
    monkeypatch.setattr(RangeFormat, "__setattr__", fail)
    with pytest.raises(ToolError, match="office_undo"):
        formats(office, [
            {"cells": "A1", "format": {"bold": True}},
            {"cells": "B1", "format": {"italic": True}},
        ])
    assert office.ws.Range("A1").Font.Bold is True and len(stack(office)) == 1
    revert(office)
    assert office.ws.Range("A1").Font.Bold is False


def test_partial_dimension_failure_retains_one_undo(office, monkeypatch):
    import pywintypes
    original = Range.ColumnWidth.fset
    def fail(rng, value):
        if rng.Column == 2:
            raise pywintypes.com_error(-2147352567, "simulated failure", None, None)
        original(rng, value)
    monkeypatch.setattr(Range, "ColumnWidth", property(Range.ColumnWidth.fget, fail))
    with pytest.raises(ToolError, match="office_undo"):
        dimensions(office, column_widths={"A": 12, "B": 24})
    assert office.ws.Columns(1).ColumnWidth == 12 and len(stack(office)) == 1
    revert(office)
    assert office.ws.Columns(1).ColumnWidth == office.ws.StandardWidth


def test_batch_undo_preserves_later_content_and_unrelated_formatting(office):
    formats(office, [
        {"cells": "A1", "format": {"bold": True}},
        {"cells": "B1", "format": {"italic": True}},
    ])
    office.ws.Range("A1").Value = "later user text"
    office.ws.Range("A1:B1").WrapText = True
    revert(office)
    assert office.ws.Range("A1").Value == "later user text"
    assert office.ws.Range("A1:B1").WrapText is True
    assert office.ws.Range("A1:B1").Font.Bold is False
    assert office.ws.Range("A1:B1").Font.Italic is False


def test_batch_total_undo_limit_rejects_before_editing(office, monkeypatch):
    from dataclasses import replace
    from office_live import config
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo_max_cells=3))
    with pytest.raises(ToolError, match="MAX_CELLS"):
        formats(office, [
            {"cells": "A1:B1", "format": {"bold": True}},
            {"cells": "C1:D1", "format": {"italic": True}},
        ])
    assert office.ws.Range("A1:D1").Font.Bold is False
    assert office.ws.Range("A1:D1").Font.Italic is False
    assert stack(office) == []


def test_empty_sheet_autofit_stays_at_a1(office, monkeypatch):
    calls = []
    autofit_spy(monkeypatch, calls)
    result = dimensions(office, autofit="both")
    assert result["cells"] == "A1"
    assert calls == [("Columns", "A1"), ("Rows", "A1")]


def test_legacy_scalar_sizes_keep_existing_undo_barrier_behavior(office, monkeypatch):
    from dataclasses import replace
    from office_live import config
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo_max_cells=2))
    result = dimensions(office, cells="A1:C2", column_width=20)
    assert result["undo"].startswith("not available:")
    assert [office.ws.Columns(i).ColumnWidth for i in range(1, 4)] == [20] * 3
    assert len(stack(office)) == 1 and not stack(office)[0].undoable
