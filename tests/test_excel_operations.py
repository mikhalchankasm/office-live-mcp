"""Office operations Excel preflight, mutations, partial failures and structural undo on fakes."""

import copy
from dataclasses import replace
from types import SimpleNamespace as NS

import pytest

from office_live import config, excel_analysis as analysis, excel_core as core, registry, excel_format as formatting, undo
from office_live.errors import ToolError
from tests.history_fakes import Collection, fake_office  # noqa: F401
from tests.office_operations_fakes import Spark, AnalysisRange, error, excel


@pytest.fixture
def o(office, monkeypatch):
    excel(office, monkeypatch)
    return office


def goal(o, **kw):
    return o.call("excel_goal_seek", **{"workbook": o.wb.Name, "sheet": "Data", "set_cell": "B1", "changing_cell": "A1", "target_value": 100, **kw})


def subtotal(o, **kw):
    return o.call("excel_subtotals", **{"workbook": o.wb.Name, "sheet": "Data", "cells": "A1:B4", "group_by": "Key", "columns": ["Value"], **kw})


def spark(o, **kw):
    return o.call("excel_sparklines", **{"workbook": o.wb.Name, "sheet": "Data", **kw})


def history(o):
    return undo.STACKS.get(undo.stack_key("workbook", o.excel, o.wb), [])


def goal_data(o):
    o.ws.Range("A1").Value2 = 2
    o.ws.Range("B1").Formula = "=A1*2"
    o.ws.results[1, 2] = 4


def subtotal_data(o):
    o.ws.Range("A1:B4").Value2 = (("Key", "Value"), ("b", 2), ("a", 3), ("b", 4))
    o.ws.Range("D10").Value2 = "below/right"
    o.ws.Range("D2").Value2 = "right"


def test_goal_convergence_settings_undo(o):
    goal_data(o)
    o.ws.Range("A1").NumberFormat = "0.00"
    result = goal(o, max_iterations=15, max_change=0.01)
    assert result["converged"] and result["new_value"] == 50 and result["difference"] == 0
    assert result["parameters"] == {"max_iterations": 15, "max_change": 0.01}
    assert (o.excel.MaxIterations, o.excel.MaxChange) == (100, 0.001)
    assert history(o)[-1].areas == [{"sheet": "Data", "address": "A1"}]
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1").Value2 == 2 and o.ws.Range("B1").Formula == "=A1*2"
    assert o.ws.Range("A1").NumberFormat == "0.00"


@pytest.mark.parametrize("keep", [False, True])
@pytest.mark.parametrize("original", [2, None])
def test_goal_nonconvergence_local_rollback_or_keep(o, keep, original):
    goal_data(o)
    o.ws.Range("A1").Value2 = original
    o.ws.converged = False
    o.excel.Calculation = -4135
    result = goal(o, keep_if_not_converged=keep)
    assert not result["converged"] and result["manual_calculation"] and result["kept"] == keep
    assert o.ws.Range("A1").Value2 == (50 if keep else original)
    assert bool(history(o)) == keep
    if keep:
        o.call("office_undo", file=o.wb.Name)
        assert o.ws.Range("A1").Value2 == original


@pytest.mark.parametrize("kw", [{"target_value": float("nan")}, {"target_value": True}, {"max_change": 0}, {"max_change": float("inf")},
                               {"max_iterations": 0}, {"max_iterations": True}, {"set_cell": "B1:B2"}, {"changing_cell": "A1:A2"},
                               {"changing_cell": "Other!A1"}])
def test_goal_invalid_inputs_before_snapshot(o, monkeypatch, kw):
    goal_data(o)
    monkeypatch.setattr(undo, "_snapshots", lambda *a, **k: pytest.fail("snapshot before preflight"))
    with pytest.raises(ToolError):
        goal(o, **kw)
    assert not o.ws.calls and not history(o)


@pytest.mark.parametrize("problem", ["no_formula", "formula", "text", "bool", "error", "protected", "pivot", "merged", "spill", "independent"])
def test_goal_state_refusals(o, problem):
    goal_data(o)
    if problem == "no_formula":
        o.ws.Range("B1").Value2 = 4
    elif problem in {"formula", "text", "bool", "error"}:
        o.ws.Range("A1").Formula = {"formula": "=1", "text": "text", "bool": True, "error": -2146826281}[problem]
    elif problem == "protected":
        o.ws.ProtectContents = True
    elif problem == "pivot":
        o.ws.pivots = Collection([NS(Name="Pivot", TableRange2=o.ws.Range("A1:B2"))])
    elif problem == "merged":
        o.ws.Range("A1").MergeCells = True
    elif problem == "spill":
        o.ws.spills = {(1, 1)}
    elif problem == "independent":
        o.ws.Range("B1").Formula = "=C1*2"
    else:
        o.wb.ActiveSheet = o.wb.Sheets.Add()
    before = copy.deepcopy(o.ws.data)
    with pytest.raises(ToolError):
        goal(o)
    assert o.ws.data == before and not history(o)


def test_goal_partial_failure_restores_settings_and_has_undo(o):
    goal_data(o)
    o.ws.fail = "goal"
    with pytest.raises(ToolError, match="applied=1.*office_undo will restore"):
        goal(o, max_iterations=10, max_change=0.1)
    assert (o.excel.MaxIterations, o.excel.MaxChange) == (100, 0.001)
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1").Value2 == 2


def test_goal_failed_progress_read_keeps_snapshot(o, monkeypatch):
    goal_data(o)
    o.ws.unreadable = False
    original_goal = AnalysisRange.GoalSeek
    original_value = AnalysisRange.Value2
    def fail(rng, value, changing):
        original_goal(rng, value, changing)
        o.ws.unreadable = True
        raise error()
    def read_value(rng):
        if o.ws.unreadable and rng.Address == "A1":
            raise error()
        return original_value.fget(rng)
    monkeypatch.setattr(AnalysisRange, "GoalSeek", fail)
    monkeypatch.setattr(AnalysisRange, "Value2", property(read_value, original_value.fset))
    with pytest.raises(ToolError, match="applied=unknown.*office_undo will restore"):
        goal(o)
    o.ws.unreadable = False
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1").Value2 == 2


def test_goal_conflict_and_unlocked_protected_cell(o):
    goal_data(o)
    o.ws.ProtectContents = True
    o.ws.Range("A1").Locked = False
    goal(o)
    o.ws.Range("A1").Value2 = 12
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.wb.Name)


def test_goal_number_format_user_edit_conflicts(o):
    goal_data(o)
    goal(o)
    o.ws.Range("A1").NumberFormat = "0.00"
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.wb.Name)
    o.call("office_undo", file=o.wb.Name, force=True)
    assert o.ws.Range("A1").Value2 == 2 and o.ws.Range("A1").NumberFormat == "General"


@pytest.mark.parametrize("sort", [False, True])
def test_subtotals_full_rows_and_order_undo(o, sort):
    subtotal_data(o)
    before = copy.deepcopy(o.ws.data)
    o.ws.Range("B3").NumberFormat = "0.00"
    result = subtotal(o, sort_first=sort, summary_below=False)
    assert result["groups"] == (2 if sort else 3) and result["inserted_rows"] == result["groups"] + 1
    assert result["unsorted_groups"] == (0 if sort else 1) and result["undo"] == "available"
    assert o.ws.calls[-1][-1] == 0  # xlSummaryAbove
    o.call("office_undo", file=o.wb.Name)
    assert {k: v for k, v in o.ws.data.items() if v is not None} == before
    assert o.ws.Range("B3").NumberFormat == "0.00"


@pytest.mark.parametrize("function,value", list(core.FUNCTIONS.items()))
def test_subtotal_function_and_absolute_column_mapping(o, function, value):
    subtotal_data(o)
    subtotal(o, group_by="A", columns=["B"], function=function)
    assert o.ws.calls[-1][1:4] == (1, value, (2,))


@pytest.mark.parametrize("problem", ["protected", "merged", "pivot", "table", "filter", "limit", "output_limit", "duplicate", "column", "group", "columns", "undo_off"])
def test_subtotal_preflight_refusals(o, monkeypatch, problem):
    subtotal_data(o)
    kw = {}
    if problem == "protected":
        o.ws.ProtectContents = True
    elif problem == "merged":
        o.ws.Range("B2").MergeCells = True
    elif problem == "pivot":
        o.ws.pivots = Collection([NS(Name="P", TableRange2=o.ws.Range("B2:D3"))])
    elif problem == "table":
        o.ws.ListObjects = Collection([NS(Name="T", Range=o.ws.Range("A2:B3"))])
    elif problem == "filter":
        o.ws.AutoFilterMode = True
        o.ws.AutoFilter = NS(Range=o.ws.Range("A1:B4"), Filters=Collection([NS(On=True)]))
    elif problem in {"limit", "output_limit", "undo_off"}:
        monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo=problem != "undo_off", undo_max_cells=7 if problem == "limit" else 10 if problem == "output_limit" else 10000))
    elif problem == "duplicate":
        o.ws.Range("B1").Value2 = "Key"
    else:
        kw = {"column": {"columns": ["Unknown"]}, "group": {"group_by": "Unknown"}, "columns": {"columns": []}}[problem]
    before = copy.deepcopy(o.ws.data)
    monkeypatch.setattr(undo, "_snapshots", lambda *a, **k: pytest.fail("snapshot before checks"))
    with pytest.raises(ToolError):
        subtotal(o, **kw)
    assert o.ws.data == before and not history(o)


def test_subtotals_partial_failure_and_conflict(o):
    subtotal_data(o)
    before = copy.deepcopy(o.ws.data)
    o.ws.fail = "subtotal"
    with pytest.raises(ToolError, match="office_undo will restore"):
        subtotal(o, sort_first=True)
    o.ws.Range("B3").Value2 = 99
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.wb.Name)
    o.call("office_undo", file=o.wb.Name, force=True)
    assert {k: v for k, v in o.ws.data.items() if v is not None} == before


def test_subtotals_barriers_for_existing_outline_and_remove(o):
    subtotal_data(o)
    o.ws.Rows(2).OutlineLevel = 2
    result = subtotal(o)
    assert "barrier" in result["undo"]
    result = subtotal(o, action="remove", cells=result["range"])
    assert "barrier" in result["undo"]
    with pytest.raises(ToolError, match="not available"):
        o.call("office_undo", file=o.wb.Name)


def test_subtotal_snapshot_failure_refuses_before_sort(o, monkeypatch):
    subtotal_data(o)
    monkeypatch.setattr(undo, "_snapshots", lambda *a, **kw: (_ for _ in ()).throw(RuntimeError("snapshot failed")))
    with pytest.raises(ToolError, match="snapshot failed"):
        subtotal(o, sort_first=True)
    assert not o.ws.calls


# PlotBy is only set for square data (live probe: Excel rejects it otherwise and orients non-square data itself)
@pytest.mark.parametrize("location,data,plot", [("D1:D2", "A1:C2", 0), ("A5:C5", "A1:C2", 0), ("D5", "A1:A2", 0),
                                                ("C1:C2", "A1:B2", 1), ("A5:B5", "A1:B2", 2)])
def test_spark_add_list_exact_undo_keeps_cell_contents(o, location, data, plot):
    o.ws.Range("A1:C2").Value2 = ((1, 2, 3), (4, 5, 6))
    o.ws.Range(location).Formula = "=1"
    before = copy.deepcopy(o.ws.data)
    result = spark(o, action="add", data=data, location=location, color="#ff0000", markers=["high", "all"])
    assert result["cells_with_content"] == int(o.ws.Range(location).Count)
    assert o.ws.sparks[0].PlotBy == plot
    assert spark(o)["groups"][0]["markers"] == ["high", "all"]
    assert o.ws.data == before
    o.call("office_undo", file=o.wb.Name)
    assert not o.ws.sparks and o.ws.data == before


def test_spark_clear_restores_all_properties_and_grouping(o):
    spark(o, action="add", data="A1:C2", location="D1:D2", markers=["negative"])
    group = o.ws.sparks[0]
    group.LineWeight, group.Axes.Vertical.CustomMinScaleValue = 2.5, -10
    group.SeriesColor.ThemeColor = 4
    before = formatting.spark_state(group, full=True)
    result = spark(o, action="clear", location="D1:D2")
    assert result["cleared_groups"] == 1 and not o.ws.sparks
    o.call("office_undo", file=o.wb.Name)
    assert formatting.spark_state(o.ws.sparks[0], full=True) == before


@pytest.mark.parametrize("kw", [{"data": "A1:C2", "location": "D1:D3"}, {"data": "A1:C2", "location": "D1:E2"},
                               {"data": "", "location": "D1"}, {"data": "A1", "location": ""}, {"type": "bad"}, {"markers": ["bad"]}])
def test_spark_invalid_inputs(o, kw):
    with pytest.raises(ToolError):
        spark(o, action="add", **kw)
    assert not o.ws.sparks and not history(o)


def test_spark_occupied_partial_clear_and_user_edit_conflicts(o):
    spark(o, action="add", data="A1:C2", location="D1:D2")
    with pytest.raises(ToolError, match="already"):
        spark(o, action="add", data="A1:C2", location="D1:D2")
    with pytest.raises(ToolError, match="whole"):
        spark(o, action="clear", location="D1")
    o.ws.sparks[0].SourceData = "Data!A1:B2"
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.wb.Name)
    # force must still refuse deleting a user's expanded group outside the recorded location.
    o.ws.sparks[0].Location = o.ws.Range("D1:D3")
    with pytest.raises(ToolError, match="outside"):
        o.call("office_undo", file=o.wb.Name, force=True)


def test_spark_partial_add_still_undoable(o):
    o.ws.fail = "add_spark"
    with pytest.raises(ToolError, match="applied=1.*office_undo will restore"):
        spark(o, action="add", data="A1:C2", location="D1:D2")
    o.call("office_undo", file=o.wb.Name)
    assert not o.ws.sparks


@pytest.mark.parametrize("problem", ["protected", "merged", "limit", "autosave"])
def test_spark_state_refusals(o, monkeypatch, problem):
    if problem == "protected":
        o.ws.ProtectContents = True
    elif problem == "merged":
        o.ws.Range("D1").MergeCells = True
    elif problem == "autosave":
        o.wb.AutoSaveOn = True
    else:
        monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo_max_cells=1))
    with pytest.raises(ToolError):
        spark(o, action="add", data="A1:C2", location="D1:D2")
    assert not o.ws.sparks and not history(o)


def test_spark_cross_sheet_source_and_clear_partial_failure(o):
    source = o.wb.Worksheets.Add(None, o.ws)
    source.Name = "Данные 'тест'"
    spark(o, action="add", data="'Данные ''тест'''!A1:C2", location="D1:D2")
    assert o.ws.sparks[0].SourceData == "'Данные ''тест'''!A1:C2"
    o.ws.fail = "delete_spark"
    with pytest.raises(ToolError, match="applied=0.*office_undo will restore"):
        spark(o, action="clear", location="D1:D2")
    assert len(o.ws.sparks) == 1
    o.ws.fail = ""
    o.call("office_undo", file=o.wb.Name)
    assert len(o.ws.sparks) == 1


def test_spark_unreadable_clear_barrier(o):
    group = Spark(o.ws.Range("D1:D2"), 1, "Data!A1:C2")
    o.ws.sparks.append(group)
    del group.LineWeight
    result = spark(o, action="clear", location="D1:D2")
    assert "barrier" in result["undo"] and not o.ws.sparks


def test_readonly_partial_actions(o, monkeypatch):
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, mode="readonly"))
    registered = {}
    monkeypatch.setattr(registry, "mcp", NS(add_tool=lambda fn, name, **kw: registered.__setitem__(name, fn)))
    registry.office_tool("excel_format", "write", read_actions=("list",))(formatting.excel_sparklines)
    assert registered["excel_sparklines"](workbook=o.wb.Name, action="list")["groups"] == []
    with pytest.raises(ToolError, match="Read-only"):
        registered["excel_sparklines"](workbook=o.wb.Name, action="add")
    for name in ("excel_goal_seek", "excel_subtotals"):
        registry.office_tool(registry.CATALOG[name].group, "write")(getattr(analysis if name == "excel_goal_seek" else core, name))
        assert not registry.CATALOG[name].registered


@pytest.mark.parametrize("formula,verified,reason", [("=A1*2", True, None), ("=C1*2", True, None),
                                                   ('=INDIRECT("A1")*2', False, "indirect"), ("=OFFSET(A1,0,0)*2", False, "offset"),
                                                   ("=NamedInput*2", True, None)])
def test_goal_inactive_sheet_parsed_dependencies(o, formula, verified, reason, monkeypatch):
    goal_data(o)
    o.ws.Range("B1").Formula = formula
    o.ws.Range("C1").Formula = "=A1"
    named = o.wb.Names.Add("NamedInput", "=Data!$A$1")
    named.RefersToRange = o.ws.Range("A1")
    other = o.wb.Worksheets.Add()
    other.Name = "Other"
    monkeypatch.setattr(AnalysisRange, "Precedents", property(lambda self: pytest.fail("Precedents forbidden")))
    monkeypatch.setattr(o.ws, "Activate", lambda: pytest.fail("Activate forbidden"))
    result = goal(o)
    assert result["dependency_verified"] is verified
    if reason:
        assert reason in result["dependency_reasons"]
    assert o.wb.ActiveSheet is other


def test_goal_dependency_node_limit_allows_attempt(o, monkeypatch):
    from office_live import formula_trace

    goal_data(o)
    o.ws.Range("B1").Formula = "=SUM(C1:C2001)"
    result = goal(o)
    assert not result["dependency_verified"] and "node_limit" in result["dependency_reasons"]
    # Independent complete parse rejects; a bounded unknown parse never claims independence.
    found, complete, reasons = formula_trace.same_sheet_dependency(o.excel, o.wb, o.ws.Range("B1"), o.ws.Range("A1"), max_nodes=2)
    assert not found and not complete and "node_limit" in reasons
