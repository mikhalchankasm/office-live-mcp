"""What-if contracts and undo without Office activation."""

import copy
from dataclasses import replace
from types import SimpleNamespace as NS

import pytest

from office_live import config, undo, xl_common as xl
from office_live.errors import ToolError
from tests.history_fakes import Collection, fake_office  # noqa: F401
from tests.office_operations_fakes import AnalysisRange, error


class WhatRange(AnalysisRange):
    @property
    def HasArray(self):
        return bool(self.Worksheet.array and set(self.positions()) & self.Worksheet.array)

    def Table(self, row, column, /):
        ws = self.Worksheet
        ws.calls.append((row if row is None else row.Address, column if column is None else column.Address))
        target = xl.sub_range(ws, self.Row + 1, self.Column + 1, self.r2, self.c2)
        ws.array = set(target.positions())
        target.Formula = "=TABLE()"
        if ws.fail == "data_table":
            raise error()

    def ClearContents(self, /):
        assert set(self.positions()) >= self.Worksheet.array
        self.Worksheet.array = set()
        self.Value2 = None


class Scenarios(Collection):
    def __init__(self, ws):
        super().__init__()
        self.ws = ws

    def Add(self, name, changing, values, comment, locked, hidden, /):
        item = NS(Name=name, ChangingCells=changing, Values=tuple(values), Comment=comment, Locked=locked, Hidden=hidden)
        def delete():
            self.items.remove(item)
            if self.ws.fail == "delete_scenario":
                raise error()
        def show():
            cells = [p for a in changing.Areas for p in a.positions()]
            for (r, c), value in zip(cells, values, strict=True):
                self.ws.Cells(r, c).Value2 = value
                self.ws.Cells(r, c).NumberFormat = "0.0"
            if self.ws.fail == "show_scenario":
                raise error()
        item.Delete, item.Show = delete, show
        self.items.append(item)
        if self.ws.fail == "add_scenario":
            raise error()
        return item

    def CreateSummary(self, kind, results, /):
        assert kind == 1 and results.HasFormula
        sheet = self.ws.Parent.Sheets.Add()
        sheet.Range("A1").Value2 = "Summary"
        if self.ws.fail == "scenario_summary":
            raise error()


@pytest.fixture
def o(office, monkeypatch):
    ws = office.ws
    ws.array, ws.calls, ws.fail = set(), [], ""
    def get_range(address):
        if "," not in address:
            return WhatRange(ws, address)
        areas = [WhatRange(ws, a) for a in address.split(",")]
        return NS(Count=sum(a.Count for a in areas), Areas=Collection(areas), Address=address)
    monkeypatch.setattr(ws, "Range", get_range)
    collection = Scenarios(ws)
    ws.Scenarios = lambda: collection
    ws.Range("A1:A2").Value2 = ((2,), (3,))
    ws.Range("B3").Formula = "=A1*A2"
    ws.Range("C3:D3").Value2 = ((4, 5),)
    ws.Range("B4:B5").Value2 = ((6,), (7,))
    return office


def call(o, action="data_table", **kw):
    args = {"cells": "B3:D5", "row_input": "A1", "column_input": "A2"} if action == "data_table" else {}
    return o.call("excel_what_if", **{"workbook": o.wb.Name, "sheet": "Data", "action": action, **args, **kw})


def add(o, **kw):
    return call(o, "add_scenario", **{"name": "Case", "changing_cells": "A1,A2", "values": [10, "007"], "comment": "saved", **kw})


@pytest.mark.parametrize("mode", [-4105, -4135, 2])
@pytest.mark.parametrize("direction", ["both", "row", "column"])
def test_table_position_dependency_modes_and_whole_array_undo(o, mode, direction):
    o.excel.Calculation = mode
    kw = {}
    if direction == "row":
        o.ws.Range("B4:B5").Formula = "=A1*2"
        kw["column_input"] = ""
    if direction == "column":
        o.ws.Range("C3:D3").Formula = "=A2*2"
        kw["row_input"] = ""
    before = copy.deepcopy(o.ws.data)
    result = call(o, **kw)
    assert result["dependency_verified"] and result["results"] == "C4:D5"
    assert bool(result["calculation_note"]) == (mode != -4105)
    # None, not Missing: Missing in the first position drops the next argument (com.py rule 1, review finding)
    assert o.ws.calls == [(None if direction == "column" else "A1", None if direction == "row" else "A2")]
    o.call("office_undo", file=o.wb.Name)
    assert not o.ws.array and all(o.ws.data.get(k) == v for k, v in before.items())
    assert o.ws.Range("C4:D5").Value2 == ((None, None), (None, None))


@pytest.mark.parametrize("problem", ["occupied", "formula_empty", "no_formula", "independent", "merged", "array", "spill", "table", "pivot", "protected", "inside", "cross_sheet", "same_input", "large", "absent_input"])
def test_table_refuses_before_undo_preparation(o, monkeypatch, problem):
    kw = {}
    if problem in {"occupied", "formula_empty"}:
        o.ws.Range("C4").Formula = "x" if problem == "occupied" else '=IF(1,"",0)'
    elif problem in {"no_formula", "independent"}:
        o.ws.Range("B3").Formula = 4 if problem == "no_formula" else "=A1*2"
    elif problem == "merged":
        o.ws.Range("C4").MergeCells = True
    elif problem == "array":
        o.ws.array = {(4, 3)}
    elif problem == "spill":
        o.ws.spills = {(4, 3)}
    elif problem in {"table", "pivot"}:
        setattr(o.ws, "ListObjects" if problem == "table" else "pivots", Collection([NS(Name="Object", Range=o.ws.Range("C4"), TableRange2=o.ws.Range("C4"))]))
    elif problem == "protected":
        o.ws.ProtectContents = True
    else:
        kw = {"inside": {"row_input": "C4"}, "cross_sheet": {"row_input": "Other!A1"}, "same_input": {"row_input": "A2"},
              "large": {"cells": "B3:XFD100"}, "absent_input": {"row_input": "", "column_input": ""}}[problem]
    monkeypatch.setattr(undo, "require_undo", lambda *a, **k: pytest.fail("preflight was incomplete"))
    with pytest.raises(ToolError):
        call(o, **kw)
    assert not o.ws.calls


def test_incomplete_dependency_proceeds_and_conflicting_edit_blocks_undo(o):
    o.ws.Range("B3").Formula = '=INDIRECT("A1")+A2'
    assert not call(o)["dependency_verified"]
    o.ws.Range("C4").NumberFormat = "0.00"
    with pytest.raises(ToolError):
        o.call("office_undo", file=o.wb.Name)


@pytest.mark.parametrize("action", ["data_table", "add_scenario", "show_scenario", "delete_scenario", "scenario_summary"])
@pytest.mark.parametrize("fail", [False, True])
def test_actions_roundtrip_and_partial_failure(o, action, fail):
    if action in {"show_scenario", "delete_scenario", "scenario_summary"}:
        add(o)
        o.ws.Scenarios().Item("Case").Locked = True
        o.ws.Scenarios().Item("Case").Hidden = True
    if action == "show_scenario":
        o.ws.Range("A1").Formula = "=1+1"
        o.ws.Range("A1").NumberFormat = "0.00"
    before = call(o, "list_scenarios")["scenarios"]
    original = o.ws.Range("A1:A2").Formula, [o.ws.Cells(r, 1).NumberFormat for r in (1, 2)]
    sheets = xl.sheet_names(o.wb)
    o.ws.fail = action if fail else ""
    def run():
        return add(o) if action == "add_scenario" else call(o, action, name="Case", result_cells="B3")
    if fail:
        with pytest.raises(ToolError, match="office_undo"):
            run()
    else:
        assert run()["undo"] == "available"
    o.ws.fail = ""
    o.call("office_undo", file=o.wb.Name)
    assert call(o, "list_scenarios")["scenarios"] == before
    assert (o.ws.Range("A1:A2").Formula, [o.ws.Cells(r, 1).NumberFormat for r in (1, 2)]) == original
    assert xl.sheet_names(o.wb) == sheets


@pytest.mark.parametrize("kw", [{"values": [1]}, {"values": [True, 2]}, {"values": [float("nan"), 2]}, {"values": [None, 2]}, {"changing_cells": "A1:A33", "values": [0] * 33}, {"name": ""}])
def test_scenario_validation(o, kw):
    with pytest.raises(ToolError):
        add(o, **kw)
    assert call(o, "list_scenarios")["scenarios"] == []


def test_scenario_duplicate_conflict_and_readonly(o, monkeypatch):
    add(o)
    with pytest.raises(ToolError):
        add(o, name="case")
    o.ws.Scenarios().Item("Case").Comment = "edited"
    with pytest.raises(ToolError):
        o.call("office_undo", file=o.wb.Name)
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, mode="readonly"))
    assert call(o, "list_scenarios")["scenarios"][0]["comment"] == "edited"
    with pytest.raises(ToolError):
        add(o)


@pytest.mark.parametrize("problem", ["locked", "structure", "result", "spill", "undo_limit", "autosave"])
def test_scenario_and_undo_preflight_refusals(o, monkeypatch, problem):
    add(o)
    if problem == "locked":
        o.ws.Scenarios().Item("Case").Locked = o.ws.ProtectScenarios = True
    elif problem == "structure":
        o.wb.ProtectStructure = True
    elif problem == "spill":
        o.ws.spills = {(1, 1)}
    elif problem == "undo_limit":
        monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo_max_cells=1))
    elif problem == "autosave":
        o.wb.AutoSaveOn = True
    before = copy.deepcopy(o.ws.data)
    action = {"locked": "delete_scenario", "structure": "scenario_summary", "result": "scenario_summary", "spill": "show_scenario"}.get(problem, "data_table")
    with pytest.raises(ToolError):
        call(o, action, name="Case", result_cells="A1" if problem == "result" else "B3")
    assert o.ws.data == before and o.ws.Scenarios().Count == 1


def test_scenario_undo_rebuilds_long_union_addresses():
    # Review finding: Range() refuses addresses over 255 characters; a 32-cell union is rebuilt area by area.
    from office_live import undo

    address = ",".join(f"$AAA${10000 + i}" for i in range(32))
    assert len(address) > 255
    calls = []
    ws = NS(Range=lambda a: (calls.append(a), NS(addr=a))[1])
    app = NS(Union=lambda x, y: NS(addr=x.addr + "," + y.addr))
    assert undo._union_range(app, ws, address).addr == address and len(calls) == 32
    assert undo._union_range(app, ws, "$A$1,$B$2") is not None and calls[-1] == "$A$1,$B$2"
