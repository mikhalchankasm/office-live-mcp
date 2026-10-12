"""Compact response contracts, using fake Office objects only."""

import copy
import inspect
import json
from types import SimpleNamespace as NS

import pytest

from office_live import excel_analysis, excel_core, links, navigation, undo, word_core
from office_live.errors import ToolError
from tests.history_fakes import fake_office  # noqa: F401


def response_bytes(result):
    return len(json.dumps(result, ensure_ascii=False).encode("utf-8"))


def test_find_ten_matches_halves_legacy_response_and_keeps_cells(office):
    office.ws.Range("A1:A10").Value = "match"
    compact = office.call("excel_find", workbook=office.wb.Name, query="match")
    legacy = copy.deepcopy(compact)
    legacy.update(links=[navigation.item("excel", office.wb, sheet=hit["sheet"], range=hit["cell"])
                         for hit in compact["matches"]], links_truncated=False)
    print(f"find: legacy={response_bytes(legacy)} bytes, compact={response_bytes(compact)} bytes")
    assert response_bytes(compact) <= response_bytes(legacy) * 0.5
    assert [hit["cell"] for hit in compact["matches"]] == [f"A{i}" for i in range(1, 11)]
    assert len(compact["links"]) == 1 and compact["links_truncated"] is False
    expanded = office.call("excel_find", workbook=office.wb.Name, query="match", per_match_links=True)
    assert expanded == legacy
    assert compact.keys() == expanded.keys()


def test_find_across_sheets_links_to_whole_workbook(office):
    office.ws.Range("A1").Value = "match"
    second = office.wb.Worksheets.Add()
    second.Name = "Other"
    second.Range("B2").Value = "match"
    result = office.call("excel_find", workbook=office.wb.Name, query="match", all_sheets=True)
    assert len(result["matches"]) == 2
    assert [links.parse(link["uri"])[1] for link in result["links"]] == [{"book": office.wb.Name}]


def test_compare_across_workbooks_has_one_link_and_unchanged_differences(office):
    office.ws.Range("A1:A2").Value = "left"
    other = office.excel.Workbooks.Add()
    other.ActiveSheet.Range("B1:B2").Value = "right"
    kwargs = dict(workbook_a=office.wb.Name, sheet_a="Data", cells_a="A1:A2",
                  workbook_b=other.Name, sheet_b="Data", cells_b="B1:B2")
    compact = office.call("excel_compare_ranges", **kwargs)
    expanded = office.call("excel_compare_ranges", **kwargs, per_match_links=True)
    assert len(compact["links"]) == 1 and len(expanded["links"]) == 4
    assert links.parse(compact["links"][0]["uri"])[1] == {"book": office.wb.Name}
    assert compact.keys() == expanded.keys()
    assert compact["differences"] == expanded["differences"]


def test_find_no_matches_keeps_previous_keys(office):
    result = office.call("excel_find", workbook=office.wb.Name, query="absent")
    assert result == {"workbook": office.wb.Name, "query": "absent", "total_matches": 0,
                      "returned": 0, "truncated": False, "matches": []}


@pytest.mark.parametrize("fn", [excel_core.excel_find, excel_analysis.excel_compare_ranges,
                                  excel_analysis.excel_find_issues, excel_analysis.excel_trace_formula,
                                  word_core.word_find, word_core.word_get_structure])
def test_multi_location_tools_expose_compact_default(fn):
    assert inspect.signature(fn).parameters["per_match_links"].default is False


@pytest.mark.parametrize("tool,data", [
    ("excel_compare_ranges", {"target_a": {"sheet": "Data"}, "differences": [{"cell_a": "A1"}, {"cell_a": "B2"}]}),
    ("excel_find_issues", {"sheet": "Data", "issues": [{"examples": [{"cell": "A1"}, {"cell": "B2"}]}]}),
    ("excel_trace_formula", {"nodes": [{"address": "Data!A1"}, {"address": "Data!B2"}]}),
    ("word_find", {"matches": [{"paragraph": 1}, {"paragraph": 2}]}),
    ("word_get_structure", {"headings": [{"paragraph": 1}, {"paragraph": 2}]}),
])
def test_other_multi_location_links_are_compact_and_opt_in(office, monkeypatch, tool, data):
    monkeypatch.setattr(navigation, "word_range", lambda *args: None)
    if tool == "excel_compare_ranges":
        data["target_a"]["workbook"] = office.wb.Name
    app, obj = (office.excel, office.wb) if tool.startswith("excel") else (office.word, office.doc)
    kind = "workbook" if tool.startswith("excel") else "document"
    recording = NS(tool=tool, kind="read", args={}, targets={(kind, obj.Name, 1): (app, obj)}, pending={})
    compact, expanded = copy.deepcopy(data), copy.deepcopy(data)
    navigation.result_links(recording, compact)
    recording.args["per_match_links"] = True
    navigation.result_links(recording, expanded)
    assert len(compact["links"]) == 1 and not compact["links_truncated"]
    assert len(expanded["links"]) == 2 and not expanded["links_truncated"]
    for key, value in data.items():
        assert compact[key] == expanded[key] == value


def test_write_summary_twenty_by_ten_reports_errors_without_grid_echo(office):
    values = [[r * 10 + c for c in range(10)] for r in range(20)]
    office.ws.results[(21, 11)] = -2146826281  # #DIV/0!, beyond the five samples
    full = office.call("excel_write_range", workbook=office.wb.Name, sheet="Data", cells="B2", values=values)
    office.call("office_undo", file=office.wb.Name)
    result = office.call("excel_write_range", workbook=office.wb.Name, sheet="Data", cells="B2",
                         values=values, readback="summary")
    print(f"write 20x10: full={response_bytes(full)} bytes, summary={response_bytes(result)} bytes")
    summary = result["readback"]
    assert isinstance(summary, dict)
    assert summary["shape"] == [20, 10] and summary["total_cells"] == 200
    assert summary["nonempty_cells"] == 200 and summary["empty_cells"] == 0
    assert len(summary["samples"]) == 5
    assert result["readback_types"] == {"number": 199, "error": 1}
    assert result["errors"] == [{"cell": "K21", "error": "#DIV/0!"}]
    assert "values" not in summary
    assert result["undo"] == "available"
    entry = undo.STACKS[undo.stack_key("workbook", office.excel, office.wb)][-1]
    assert entry.areas == [{"sheet": "Data", "address": "B2:K21"}]
    office.call("office_undo", file=office.wb.Name)
    assert office.ws.Range("B2").Value is None


def test_write_summary_counts_entire_expanded_target(office):
    office.ws.results[(31, 11)] = -2146826281
    result = office.call("excel_write_range", workbook=office.wb.Name, sheet="Data", cells="B2:K31",
                         values=7, readback="summary")
    assert result["written_shape"] == [1, 1]
    assert result["readback"]["shape"] == [30, 10]
    assert result["readback"]["total_cells"] == 300
    assert result["readback_types"] == {"number": 299, "error": 1}
    assert result["errors"] == [{"cell": "K31", "error": "#DIV/0!"}]


def test_write_summary_blank_counts_and_samples_use_actual_cells(office):
    result = office.call("excel_write_range", workbook=office.wb.Name, sheet="Data", cells="C4",
                         values=[["text", None], [True, ""]], readback="summary")
    assert result["readback"]["shape"] == [2, 2]
    assert result["readback"]["nonempty_cells"] == 2
    assert result["readback"]["empty_cells"] == 2
    assert result["readback"]["samples"][0] == {"cell": "C4", "value": "text"}


def test_write_default_keeps_full_response_keys_and_none_keeps_errors(office):
    kwargs = dict(workbook=office.wb.Name, sheet="Data", cells="A1", values=[[1, 2]])
    full = office.call("excel_write_range", **kwargs)
    office.call("office_undo", file=office.wb.Name)
    explicit = office.call("excel_write_range", **kwargs, readback="full")
    assert full == explicit
    assert full.keys() == {"ok", "workbook", "sheet", "cells", "written_shape", "overwritten_nonempty_cells",
                           "readback", "readback_types", "errors", "links", "links_truncated", "undo"}
    assert full["readback"] == [[1, 2]]
    office.ws.results[(1, 2)] = -2146826281
    minimal = office.call("excel_write_range", **kwargs, readback="none")
    assert "readback" not in minimal and "readback_types" not in minimal
    assert minimal["errors"] == [{"cell": "B1", "error": "#DIV/0!"}]
    assert minimal["undo"] == "available"


def test_invalid_readback_refuses_before_write(office):
    with pytest.raises(ToolError, match="readback"):
        office.call("excel_write_range", workbook=office.wb.Name, sheet="Data", cells="A1", values=99, readback="invalid")
    assert office.ws.Range("A1").Value is None
    assert not undo.STACKS
