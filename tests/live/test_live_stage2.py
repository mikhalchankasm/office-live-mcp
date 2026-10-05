"""Stage 2 live contract tests. Written only; explicit OFFICE_LIVE_LIVE_TESTS=1 required.

Only the fixture's own marked temporary workbook is touched. Journals/audit are isolated.
"""

import json
from pathlib import Path

import pytest

from tests.live.conftest import read
from tests.live.mcpclient import Stdio, ToolFailed
from tests.live.test_live_stage1 import in_excel

pytestmark = pytest.mark.live


@pytest.fixture
def srv(tmp):
    client = Stdio(env={"OFFICE_LIVE_MODE": "full", "OFFICE_LIVE_JOURNAL_DIR": str(Path(tmp) / "journal"),
                        "OFFICE_LIVE_AUDIT_LOG": str(Path(tmp) / "audit.jsonl"), "OFFICE_LIVE_AUDIT_CONTENT": "1"})
    yield client
    client.close()


@pytest.fixture
def stage_book(srv, wb, tmp):
    path = str(Path(tmp) / "stage2.xlsx")
    srv.call("excel_save_as", workbook=wb, path=path)
    return path


def test_live_split_preview_literal_types_full_undo_and_conflict(srv, stage_book):
    book = stage_book
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=[["007|=x|12,5"], ["TRUE|1,234| text "]], value_mode="text")
    srv.call("excel_format_range", workbook=book, sheet="Data", cells="C1", number_format="text")
    before = read(srv, book, "A1:C2")
    args = dict(workbook=book, sheet="Data", cells="A:A", delimiter="pipe", types=["text", "number", "number"])
    previous = srv.call("office_undo", file=book, action="history")
    assert srv.call("excel_split_column", **args)["output_range"] == "A1:C2"
    assert srv.call("office_undo", file=book, action="history") == previous and read(srv, book, "A1:C2") == before
    result = srv.call("excel_split_column", **args, preview=False)
    assert result["ambiguous_numbers"] == ["B2"] and result["undo"] == "available"
    assert read(srv, book, "A1:C2") == [["007", "=x", 12.5], ["TRUE", "1,234", "text"]]
    assert not in_excel(book, lambda app, ws: bool(ws.Range("B1").HasFormula))
    in_excel(book, lambda app, ws: setattr(ws.Range("C2"), "Value2", "user edit"))
    with pytest.raises(ToolFailed, match="later edits"):
        srv.call("office_undo", file=book)
    srv.call("office_undo", file=book, force=True)
    assert read(srv, book, "A1:C2") == before
    assert in_excel(book, lambda app, ws: ws.Range("C1").NumberFormat) == "@"


def test_live_split_fixed_width_destination(srv, stage_book):
    srv.call("excel_write_range", workbook=stage_book, sheet="Data", cells="A1", values=[["007ABCrest"]], value_mode="text")
    result = srv.call("excel_split_column", workbook=stage_book, sheet="Data", cells="A1", widths=[3, 3], destination="D3", preview=False)
    assert result["parts"] == 3 and read(srv, stage_book, "D3:F3") == [["007", "ABC", "rest"]]
    srv.call("office_undo", file=stage_book)
    assert read(srv, stage_book, "D3:F3") == [[None, None, None]]


@pytest.mark.parametrize("problem", ["occupied", "formula", "merge", "table", "pivot", "protected"])
def test_live_split_refusals(srv, stage_book, problem):
    book = stage_book
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=[["a|b|c"], ["x|y|z"]], value_mode="text")
    if problem == "occupied":
        in_excel(book, lambda app, ws: setattr(ws.Range("C2"), "Value2", "keep"))
    elif problem == "formula":
        in_excel(book, lambda app, ws: setattr(ws.Range("C2"), "Formula", '=IF(TRUE,"",1)'))
    elif problem == "merge":
        srv.call("excel_merge_cells", workbook=book, sheet="Data", cells="C2:D2")
    elif problem == "table":
        in_excel(book, lambda app, ws: ws.ListObjects.Add(1, ws.Range("C1:D3"), None, 1))
    elif problem == "pivot":
        srv.call("excel_write_range", workbook=book, sheet="Data", cells="G1", values=[["Key", "Value"], ["a", 1], ["b", 2]])
        srv.call("excel_create_pivot_table", workbook=book, source="G1:H3", source_sheet="Data", dest_sheet="Data", dest_cell="C1",
                 name="Stage2Pivot", rows=["Key"], values=[{"field": "Value", "function": "sum"}])
    else:
        srv.call("excel_protection", workbook=book, sheet="Data", action="protect")
    before = read(srv, book, "A1:D3")
    try:
        with pytest.raises(ToolFailed):
            srv.call("excel_split_column", workbook=book, sheet="Data", cells="A1:A2", delimiter="pipe", preview=False)
        assert read(srv, book, "A1:D3") == before
    finally:
        if problem == "protected":
            srv.call("excel_protection", workbook=book, sheet="Data", action="unprotect")


@pytest.mark.parametrize("date1904", [False, True])
def test_live_dates_epoch_strict_format_and_undo(srv, stage_book, date1904):
    from office_live.util import excel_date_serial

    book = stage_book
    in_excel(book, lambda app, ws: setattr(ws.Parent, "Date1904", date1904))
    in_excel(book, lambda app, ws: setattr(ws.Range("A1"), "ColumnWidth", 28))
    text = "29.02.2024 12:30:15"
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=[[text], ["29.02.2023 12:30:15"]], value_mode="text")
    srv.call("excel_format_range", workbook=book, sheet="Data", cells="A1", number_format="text")
    result = srv.call("excel_clean_text", workbook=book, sheet="Data", cells="A1:A2", operations=["text_to_date"],
                      date_format="DD.MM.YYYY HH:MM:SS", preview=False)
    assert result["not_date"] == ["A2"]
    actual = in_excel(book, lambda app, ws: ws.Range("A1").Value2)
    assert actual == pytest.approx(excel_date_serial(text, "DD.MM.YYYY HH:MM:SS", date1904))
    assert in_excel(book, lambda app, ws: str(ws.Range("A1").Text)) == text
    srv.call("office_undo", file=book)
    assert read(srv, book, "A1") == [[text]] and in_excel(book, lambda app, ws: ws.Range("A1").NumberFormat) == "@"


def test_live_trace_cross_sheet_names_table_and_no_ui(srv, stage_book):
    book = stage_book

    def prepare(app, ws):
        other = ws.Parent.Worksheets.Add(None, ws)
        other.Name = "Other sheet"
        other.Range("A1:A3").Value2 = (("Amount",), (2,), (3,))
        table = other.ListObjects.Add(1, other.Range("A1:A3"), None, 1)
        table.Name = "Stage2Table"
        ws.Parent.Names.Add("Stage2Input", "='Other sheet'!$A$2", True)
        ws.Range("B1").Formula = "=SUM('Other sheet'!A2:A3,Stage2Input,Stage2Table[Amount])"
        other.Range("C1").Formula = "=SUM(Data!B1:B3)"
        ws.Activate()  # setup only; the trace itself must not activate any sheet
        ws.Range("D8").Select()

    in_excel(book, prepare)

    def state(app, ws):
        return app.ActiveWorkbook.FullName, app.ActiveSheet.Name, app.Selection.Address, ws.Parent.Saved, ws.Range("B1").Formula

    before = in_excel(book, state)
    result = srv.call("excel_trace_formula", workbook=book, sheet="Data", cell="B1", depth=2)
    assert {"'Other sheet'!A2", "'Other sheet'!A3"} <= {n["address"] for n in result["nodes"]}
    assert not result["complete"] and "inactive_sheet_com_unavailable" in result["incomplete_reasons"]
    dependent = srv.call("excel_trace_formula", workbook=book, sheet="Data", cell="B1", direction="dependents")
    assert "'Other sheet'!C1" in {n["address"] for n in dependent["nodes"]}
    assert in_excel(book, state) == before


def test_live_trace_inactive_local_dependents_no_cells_and_limits(srv, stage_book):
    book = stage_book

    def prepare(app, ws):
        ws.Range("A1:B1").Value2 = ((1, 2),)
        ws.Range("C1").Formula = "=SUM(A1:B1)"
        other = ws.Parent.Worksheets.Add(None, ws)
        other.Name = "ActiveOther"
        other.Activate()

    in_excel(book, prepare)
    result = srv.call("excel_trace_formula", workbook=book, sheet="Data", cell="A1", direction="dependents")
    assert "Data!C1" in {n["address"] for n in result["nodes"]} and not result["complete"]
    assert in_excel(book, lambda app, ws: app.ActiveSheet.Name) == "ActiveOther"
    in_excel(book, lambda app, ws: ws.Activate())
    empty = srv.call("excel_trace_formula", workbook=book, sheet="Data", cell="Z10")
    assert len(empty["nodes"]) == 1 and empty["edges"] == []
    limited = srv.call("excel_trace_formula", workbook=book, sheet="Data", cell="C1", max_nodes=2)
    assert limited["truncated"] and not limited["complete"]


@pytest.mark.parametrize("formula,reason", [('=INDIRECT("A1")', "indirect"), ("=OFFSET(A1,0,0)", "offset"),
                                          ("=INDEX(A1:A2,1)", "index"), ("=LET(x,A1,x)", "let")])
def test_live_trace_dynamic_formulas_incomplete(srv, stage_book, formula, reason):
    in_excel(stage_book, lambda app, ws: setattr(ws.Range("B1"), "Formula", formula))
    result = srv.call("excel_trace_formula", workbook=stage_book, sheet="Data", cell="B1")
    assert not result["complete"] and reason in result["incomplete_reasons"]


@pytest.mark.parametrize("case,reason", [("spill", "spill_reference"), ("name_formula", "name_not_range"),
                                      ("lambda_name", "named_function"), ("3d", "3d_reference")])
def test_live_trace_spill_formula_name_and_3d(srv, stage_book, case, reason):
    def prepare(app, ws):
        if case == "spill":
            ws.Range("A1").Formula2 = "=SEQUENCE(2)"
            ws.Range("B1").Formula2 = "=SUM(A1#)"
        elif case == "name_formula":
            ws.Range("B1").Formula = "=Stage2Dynamic"
        elif case == "lambda_name":
            ws.Range("B1").Formula = "=Stage2Lambda(A1)"
        else:
            other = ws.Parent.Worksheets.Add(None, ws)
            other.Name = "Other"
            ws.Range("B1").Formula = "=SUM(Data:Other!A1)"
        ws.Activate()

    # имена — через сам инструмент: Names.Add по COM разбирает формулу по локальным правилам (в русском Excel ',' — ошибка)
    names = {"name_formula": ("Stage2Dynamic", "=OFFSET(Data!$A$1,0,0)"), "lambda_name": ("Stage2Lambda", "=LAMBDA(x,x+Data!$C$1)")}
    if case in names:
        srv.call("excel_manage_names", workbook=stage_book, action="add", name=names[case][0], formula=names[case][1])
    in_excel(stage_book, prepare)
    result = srv.call("excel_trace_formula", workbook=stage_book, sheet="Data", cell="B1")
    assert not result["complete"] and reason in result["incomplete_reasons"]


@pytest.mark.parametrize("scope", ["sheet", "workbook"])
def test_live_protection_no_password_inverse_flags_locked(srv, stage_book, scope):
    book = stage_book
    args = dict(workbook=book, sheet="Data", scope=scope)
    if scope == "sheet":
        in_excel(book, lambda app, ws: setattr(ws.Range("A2"), "Locked", False))
    extras = {"unlocked_cells": ["A1:A2"], "allow": ["sort", "filter", "format_cells", "edit_objects"]} if scope == "sheet" else {}
    result = srv.call("excel_protection", **args, action="protect", **extras)
    assert result["undo"] == "available"
    state = srv.call("excel_protection", **args)["state"]
    srv.call("excel_protection", **args, action="unprotect")
    srv.call("office_undo", file=book)
    assert srv.call("excel_protection", **args)["state"] == state
    srv.call("office_undo", file=book)
    if scope == "sheet":
        assert in_excel(book, lambda app, ws: (bool(ws.Range("A1").Locked), bool(ws.Range("A2").Locked))) == (True, False)
    else:
        assert not srv.call("excel_protection", **args)["state"]["structure"]


@pytest.mark.parametrize("scope", ["sheet", "workbook"])
def test_live_password_barrier_redaction_and_explicit_empty_no_prompt(srv, stage_book, tmp, scope):
    book, secret, wrong = stage_book, "Stage2-Pass-42", "Wrong-Stage2"
    args = dict(workbook=book, sheet="Data", scope=scope)
    srv.call("office_journal", workbook=book, action="enable_sheet")
    result = srv.call("excel_protection", **args, action="protect", password=secret)
    assert result["undo"].startswith("not available")
    try:
        for password in (wrong, ""):
            with pytest.raises(ToolFailed, match="Неверный пароль") as error:
                srv.call("excel_protection", **args, action="unprotect", password=password)
            assert secret not in str(error.value) and wrong not in str(error.value)
        history = srv.call("office_undo", file=book, action="history")
        assert secret not in json.dumps(history)
        assert secret not in json.dumps(read(srv, book, "A1:E10", sheet="Лог"))
        for path in [Path(tmp) / "audit.jsonl", *(Path(tmp) / "journal").glob("*.md")]:
            assert secret not in path.read_text("utf-8") and wrong not in path.read_text("utf-8")
    finally:
        srv.call("excel_protection", **args, action="unprotect", password=secret)
