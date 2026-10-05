"""Сравнение и очистка: только COM-фейки, включая обёртки, отказ до снимка и частичную запись."""

import datetime
from dataclasses import replace

import pytest
import pywintypes

from office_live import config, excel_analysis, excel_core, registry, undo
from office_live.errors import AppBusyError, ToolError
from office_live.util import EXCEL_ERRORS
from tests.history_fakes import Collection, Excel, Range, fake_office  # noqa: F401


def compare(o, **kwargs):
    return o.call("excel_compare_ranges", **{"workbook_a": o.wb.Name, "sheet_a": "Data", "cells_a": "A1:A4", "cells_b": "B1:B4", **kwargs})


def clean(o, **kwargs):
    return o.call("excel_clean_text", **{"workbook": o.wb.Name, "sheet": "Data", "cells": "A1:A4", "operations": ["trim"], **kwargs})


def history(o):
    return undo.STACKS.get(undo.stack_key("workbook", o.excel, o.wb), [])


def test_compare_empty_types_errors_tolerance_and_bools(office):
    o = office
    o.ws.data = {(1, 1): None, (1, 2): "", (2, 1): None, (2, 2): 0, (3, 1): 12, (3, 2): "12", (4, 1): True, (4, 2): 1}
    result = compare(o)
    assert result["summary"]["differences"] == {"value": 0, "type": 3, "formula": 0, "format": 0}
    errors = list(EXCEL_ERRORS)
    o.ws.data = {(1, 1): errors[0], (1, 2): errors[0], (2, 1): errors[1], (2, 2): errors[2],
                 (3, 1): 1.1, (3, 2): 1.2, (4, 1): False, (4, 2): True}
    result = compare(o, tolerance=0.11)
    assert result["summary"]["differences"]["value"] == 2
    assert result["differences"][0]["a"] == EXCEL_ERRORS[errors[1]]
    assert not history(o) and o.excel.Workbooks.Count == 1


def test_compare_dates_match_read_range_and_precision(office, monkeypatch):
    o = office
    o.ws.data = {(1, 1): 45352, (1, 2): 45352, (2, 1): 1.12345678, (2, 2): 1.12345679}
    o.ws.results = {(1, 1): datetime.datetime(2024, 3, 1), (1, 2): datetime.datetime(2024, 3, 1)}
    monkeypatch.setattr(Range, "Value2", property(lambda self: self.Formula))
    assert compare(o)["summary"]["differences"]["value"] == 1
    o.ws.results[1, 2] = datetime.datetime(2024, 3, 2)
    assert compare(o)["differences"][0]["a"] == "2024-03-01"


def test_compare_normalization_truncation_and_extra_rectangles(office):
    o = office
    o.ws.Range("A1:B3").Value = ((" A\u00a0  b ", "a b"), ("x" * 220, "y" * 250), ("last", "other"))
    assert compare(o, ignore_case=True, ignore_whitespace=True)["summary"]["differences"]["value"] == 2
    result = compare(o, ignore_case=True, ignore_whitespace=True, max_differences=1)
    assert result["truncated"] and result["remaining_differences"] == 1
    assert len(result["differences"][0]["a"]) <= 200 and "truncated" in result["differences"][0]["a"]
    result = compare(o, cells_a="A1:C4", cells_b="E1:F2")
    assert result["only_in_a"] == ["A3:C4", "C1:C2"] and result["only_in_b"] == []
    assert result["summary"]["compared_cells"] == 4 and result["summary"]["rows_only_in_a"] == 2


def test_compare_formulas_and_literal_formula_text(office):
    o = office
    o.ws.Range("A1:B3").Value = (("=A2", "=B2"), ("'=A2", "=A2"), (5, "5"))
    result = compare(o, compare="formulas")
    assert [d["kind"] for d in result["differences"]] == ["formula", "formula", "type"]
    assert result["summary"]["compared_cells"] == 4


def test_compare_formats_boundary_and_properties(office):
    o = office
    o.ws.Range("A2").Font.Bold = True
    o.ws.Range("B3").Font.Italic = True
    o.ws.Range("B4").Font.Color = 255
    o.ws.Range("B5").Interior.Color = 255
    o.ws.Range("B6").NumberFormat = "@"
    result = compare(o, compare="formats", cells_a="A1:A5000", cells_b="B1:B5000")
    assert result["summary"]["compared_cells"] == 5000
    assert result["summary"]["differences"]["format"] == 5


@pytest.mark.parametrize("args,message", [
    ({"cells_a": ""}, "required"), ({"cells_b": "A1:A4"}, "different"), ({"compare": "oops"}, "compare"),
    ({"match": "oops"}, "match"), ({"tolerance": -1}, "tolerance"), ({"tolerance": float("nan")}, "tolerance"),
    ({"tolerance": float("inf")}, "tolerance"), ({"max_differences": 0}, "max_differences"),
    ({"max_differences": 2001}, "max_differences"), ({"max_differences": 1.5}, "max_differences"),
    ({"match": "key"}, "key_column"), ({"cells_b": "B1:B5001", "compare": "formats"}, "5000"),
    ({"cells_a": "A1:A200001"}, "200000"), ({"cells_b": "B1:B200001", "compare": "formulas"}, "200000"),
])
def test_compare_refuses_before_read(office, monkeypatch, args, message):
    monkeypatch.setattr(excel_analysis, "_compare_values", lambda *a: pytest.fail("must not read"))
    with pytest.raises(ToolError, match=message):
        compare(office, **args)
    assert not undo.STACKS


def test_compare_keys_reordered_columns_rows_and_missing(office):
    o = office
    o.ws.Range("A1:C5").Value = (("id", "value", "old"), (1, "same", 0), (2, "changed", 0), (3, "deleted", 0), (None, "skip", 0))
    o.ws.Range("E1:G5").Value = (("value", "id", "new"), ("inserted", 4, 0), ("change", 2, 0), ("same", 1, 0), ("skip", "", 0))
    result = compare(o, cells_a="A1:C5", cells_b="E1:G5", match="key", key_column="id")
    assert result["summary"] == {"compared_cells": 4, "compared_rows": 2, "differences": {"value": 1, "type": 0, "formula": 0, "format": 0},
                                 "rows_only_in_a": 1, "rows_only_in_b": 1, "skipped_empty_keys_a": 1, "skipped_empty_keys_b": 1}
    assert result["differences"][0] == {"cell_a": "B3", "cell_b": "E3", "key": 2, "column": "value", "kind": "value", "a": "changed", "b": "change"}
    assert result["only_in_a"] == ["A4:C4"] and result["only_in_b"] == ["E2:G2"]
    assert result["headers_only_in_a"] == ["old"] and result["headers_only_in_b"] == ["new"]


@pytest.mark.parametrize("problem,message", [("header", "Duplicate header"), ("key", "Duplicate keys"), ("missing", "not found")])
def test_compare_key_rejections(office, problem, message):
    o = office
    o.ws.Range("A1:B3").Value = (("ID", "Name"), (" a ", 2), ("A", 3))
    o.ws.Range("D1:E3").Value = (("ID", "Name"), ("a", 2), ("b", 3))
    if problem == "header":
        o.ws.Range("E1").Value = " id "
        o.ws.Range("A3").Value = "b"
    with pytest.raises(ToolError, match=message):
        compare(o, cells_a="A1:B3", cells_b="D1:E3", match="key", key_column="unknown" if problem == "missing" else "ID",
                ignore_case=True, ignore_whitespace=True)


def test_compare_two_books_instances_and_clipping(office, monkeypatch):
    from office_live import com

    o = office
    other = Excel()
    other.Hwnd = 55
    book = other.Workbooks.Add()
    book.Name = book.FullName = "Second"
    o.ws.Range("A1:A2").Value = (("ID",), (5,))
    book.ActiveSheet.Range("A1:A2").Value = (("ID",), (5,))
    monkeypatch.setattr(com, "apps", lambda *a, **kw: [o.excel, other])
    result = compare(o, workbook_b="Second", cells_a="A:A", cells_b="A:A", match="key", key_column="A")
    assert not result["differences"] and result["summary"]["compared_rows"] == 1
    result = compare(o, workbook_b="Second", cells_a="5:5", cells_b="5:5")
    assert result["summary"]["compared_cells"] == 0


def test_compare_prefixed_references_and_inherited_addresses(office):
    o = office
    other = o.wb.Sheets.Add(None, o.ws)
    other.Name = "Other"
    o.ws.Range("A1").Value = other.Range("A1").Value = 1
    assert not compare(o, cells_a="Data!A1", cells_b="", sheet_b="Other")["differences"]
    assert not compare(o, cells_a="A1", cells_b="Other!A1")["differences"]


@pytest.mark.parametrize("tool", ["compare", "clean"])
def test_new_excel_tools_reject_multi_area_before_read_or_snapshot(office, monkeypatch, tool):
    original = Range.__init__

    def multiple(self, *args):
        original(self, *args)
        self.Areas = Collection([self, self])

    monkeypatch.setattr(Range, "__init__", multiple)
    monkeypatch.setattr(undo, "_snapshots", lambda *a, **kw: pytest.fail("must validate first"))
    with pytest.raises(ToolError, match="Multi-area"):
        compare(office) if tool == "compare" else clean(office, preview=False)


def test_compare_exact_values_limit(office):
    result = compare(office, cells_a="A1:A200000", cells_b="B1:B200000")
    assert not result["differences"] and result["summary"]["compared_cells"] == 200000


@pytest.mark.parametrize("operations,source,expected", [
    (["trim"], "  a\u00a0 \u202fb  ", "a b"), (["clean"], "a\t\x00\u200b\u200c\u200d\ufeff\nB", "a\nB"),
    (["upper"], "Straße", "STRASSE"), (["lower"], "ТЕКСТ", "текст"), (["proper"], "иВАН пЕТРОВ", "Иван Петров"),
    (["text_to_number", "trim", "clean"], "\u200b 12,50  ", 12.5), (["lower", "clean", "trim"], " A\u200b   B ", "a b"),
])
def test_clean_operations_fixed_order(office, operations, source, expected):
    o = office
    o.ws.Range("A1").Value = source
    result = clean(o, operations=operations)
    assert result["examples"][0]["after"] == expected
    assert o.ws.Range("A1").Value == source and not history(o) and o.excel.Workbooks.Count == 1


def test_clean_preview_has_no_journal_audit_events_or_snapshot(office, monkeypatch):
    from office_live import journal

    o = office
    o.ws.Range("A1").Value = " x "
    o.wb.AutoSaveOn = True
    monkeypatch.setattr(journal, "append", lambda *a, **kw: pytest.fail("preview journal"))
    monkeypatch.setattr(journal, "append_sheet", lambda *a, **kw: pytest.fail("preview sheet"))
    monkeypatch.setattr(registry, "_audit", lambda *a, **kw: pytest.fail("preview audit"))
    monkeypatch.setattr(undo, "_snapshots", lambda *a, **kw: pytest.fail("preview snapshot"))
    assert clean(o)["preview"] is True
    assert o.excel.EnableEvents and not history(o)


def test_clean_identifiers_and_ambiguity(office):
    o = office
    texts = ["007", "+7 999 123 45 67", "+79991234567", "1.2.3", "1234567890123456", "12.03.2026", "2026-10-05", "1,234", "1.234", "12,5"]
    o.ws.Range("A1:A10").Value = tuple((x,) for x in texts)
    result = clean(o, cells="A1:A10", operations=["text_to_number"])
    assert result["changed"] == 1 and result["skipped_leading_zero"] == 1
    assert result["skipped_ambiguous"] == 2 and result["ambiguous"] == ["A8", "A9"]
    result = clean(o, cells="A8", operations=["text_to_number"], decimal_separator=",")
    assert result["examples"][0]["after"] == 1.234


def test_clean_preview_examples_limit_and_clipped_column(office):
    o = office
    o.ws.Range("A5:A64").Value = " x "
    result = clean(o, cells="A:A")
    assert result["changed"] == 60 and len(result["examples"]) == 50 and result["examples"][0]["cell"] == "A5"
    assert clean(o, cells="C:C")["changed"] == 0


def test_clean_nontexts_and_empty_result(office):
    o = office
    o.ws.data = {(1, 1): 7, (2, 1): True, (3, 1): next(iter(EXCEL_ERRORS)), (4, 1): None, (5, 1): "   ", (6, 1): "'001"}
    result = clean(o, cells="A1:A6", preview=False)
    assert result["skipped_non_text"] == 4 and result["applied"] == 1 and o.ws.Range("A5").Value == ""
    result = clean(o, cells="A1:A6", operations=["lower"], preview=False)
    assert result["skipped_non_text"] == 4 and o.ws.data[3, 1] == next(iter(EXCEL_ERRORS))
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.data[3, 1] == next(iter(EXCEL_ERRORS)) and o.ws.Range("A5").Value == "   "


def test_clean_expanded_unicode_too_long_refuses_before_snapshot(office, monkeypatch):
    o = office
    o.ws.Range("A1").Value = "ß" * 20000
    monkeypatch.setattr(undo, "_snapshots", lambda *a, **kw: pytest.fail("must validate first"))
    with pytest.raises(ToolError, match="32767"):
        clean(o, operations=["upper"], preview=False)
    assert o.ws.Range("A1").Value == "ß" * 20000


def test_clean_write_groups_formulas_formats_and_undo(office, monkeypatch):
    o = office
    o.ws.Range("A1:B4").Value = (("  12,5 ", "keep"), ("  00123 ", 8), ("=1+1", True), ("  =x  ", None))
    o.ws.Range("A1").NumberFormat = "@"
    before = dict(o.ws.data)
    writes = []
    assign = Range._assign

    def observed(self, value, formula=False):
        writes.append((self.Address, value))
        return assign(self, value, formula)

    monkeypatch.setattr(Range, "_assign", observed)
    result = clean(o, cells="A1:B4", operations=["trim", "text_to_number"], preview=False)
    assert result["applied"] == 3 and result["format_changed"] == ["A1"] and result["undo"] == "available"
    assert result["skipped_formulas"] == 1 and result["skipped_non_text"] == 3
    assert [a for a, _ in writes] == ["A1:A2", "A4"]
    assert o.ws.Range("A2").Value == "00123" and o.ws.Range("A4").Value == "=x" and not o.ws.Range("A4").HasFormula
    assert o.ws.Range("A1").NumberFormat == "General"
    assert history(o)[-1].areas == [{"sheet": "Data", "address": "A1:B4"}]
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.data == before and o.ws.Range("A1").NumberFormat == "@"


@pytest.mark.parametrize("edit", ["value", "format", "prefix"])
def test_clean_undo_conflict_whole_range(office, edit):
    o = office
    o.ws.Range("A1:A2").Value = ((" x ",), ("untouched",))
    clean(o, cells="A1:A2", preview=False)
    if edit == "value":
        o.ws.Range("A2").Value = "user"
    elif edit == "format":
        o.ws.Range("A2").NumberFormat = "@"
    else:
        o.ws.prefixes[2, 1] = "'"
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.wb.Name)


def test_clean_rewrites_unexpected_coercion_as_text(office, monkeypatch):
    o = office
    o.ws.Range("A1").Value = " abc "
    assign = Range._assign

    def coerce(self, value, formula=False):
        assign(self, value, formula)
        if value == (("abc",),):
            self.Worksheet.data[1, 1] = 42

    monkeypatch.setattr(Range, "_assign", coerce)
    assert clean(o, preview=False)["rewritten_as_text"] == ["A1"]
    assert o.ws.Range("A1").Value == "abc"


def test_clean_merged_anchor_and_followers(office):
    o = office
    o.ws.merges = ["A1:B1"]
    o.ws.Range("A1:B1").MergeCells = True
    o.ws.Range("A1").Value = " x "
    result = clean(o, cells="A1:B1", preview=False)
    assert result["changed"] == 1 and result["skipped_merged"] == 1
    assert o.ws.Range("A1").Value == "x"
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1").Value == " x "


@pytest.mark.parametrize("args,message", [
    ({"operations": []}, "operations"), ({"operations": ["oops"]}, "operations"), ({"operations": ["upper", "lower"]}, "at most one"),
    ({"operations": ["trim", "trim"]}, "unique"), ({"decimal_separator": ";"}, "decimal_separator"), ({"cells": ""}, "required"),
    ({"cells": "A1:A100001"}, "100000"),
])
def test_clean_input_refusals_no_snapshot(office, monkeypatch, args, message):
    monkeypatch.setattr(undo, "_snapshots", lambda *a, **kw: pytest.fail("must validate first"))
    with pytest.raises(ToolError, match=message):
        clean(office, preview=False, **args)
    assert not history(office)


@pytest.mark.parametrize("problem,message", [("protected", "protected"), ("pivot", "Pivot1"), ("autosave", "AutoSave"), ("undo_limit", "MAX_CELLS"), ("snapshot", "prepare office_undo")])
def test_clean_state_refusals_leave_data_and_prior_history(office, monkeypatch, problem, message):
    from types import SimpleNamespace as NS

    o = office
    o.ws.Range("A1").Value = " x "
    if problem == "protected":
        o.ws.ProtectContents = True
    elif problem == "pivot":
        o.ws.pivots = Collection([NS(Name="Pivot1", TableRange2=o.ws.Range("A1"))])
    elif problem == "autosave":
        o.wb.AutoSaveOn = True
    elif problem == "undo_off":
        monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo=False))
    elif problem == "undo_limit":
        monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo_max_cells=1))
    else:
        monkeypatch.setattr(undo, "_snapshots", lambda *a, **kw: (_ for _ in ()).throw(ValueError("snapshot failed")))
    with pytest.raises(ToolError, match=message):
        clean(o, preview=False)
    assert o.ws.Range("A1").Value == " x " and not history(o)


@pytest.mark.parametrize("after_assign", [False, True])
def test_clean_partial_com_failure_retains_undo(office, monkeypatch, after_assign):
    o = office
    o.ws.Range("A1:A3").Value = ((" a ",), ("=1",), (" 12 ",))
    o.ws.Range("A3").NumberFormat = "@"
    original = dict(o.ws.data)
    assign = Range._assign

    def fail(self, value, formula=False):
        if self.Address == "A3":
            if after_assign:
                assign(self, value, formula)
            raise pywintypes.com_error(-2147352567, "N-th group failed", None, None)
        assign(self, value, formula)

    monkeypatch.setattr(Range, "_assign", fail)
    with pytest.raises(ToolError, match=f"applied={2 if after_assign else 1}.*office_undo will restore"):
        clean(o, cells="A1:A3", operations=["trim", "text_to_number"], preview=False)
    assert history(o)[-1].undoable
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.data == original and o.ws.Range("A3").NumberFormat == "@"


@pytest.mark.parametrize("unreadable_snapshot", [False, True])
def test_clean_busy_readback_keeps_snapshot_for_explicit_force(office, monkeypatch, unreadable_snapshot):
    o = office
    o.ws.Range("A1:A3").Value = ((" a ",), ("=1",), (" b ",))
    before = dict(o.ws.data)
    busy = False
    assign, matches, fingerprint, finish = Range._assign, excel_core._clean_matches, undo.excel_fingerprint, undo._finish_excel

    def fail(self, value, formula=False):
        nonlocal busy
        if self.Address == "A3":
            busy = True
            raise AppBusyError("Office busy")
        assign(self, value, formula)

    def reading(fn, *args):
        if busy:
            raise AppBusyError("Office busy")
        return fn(*args)

    monkeypatch.setattr(Range, "_assign", fail)
    monkeypatch.setattr(excel_core, "_clean_matches", lambda *args: reading(matches, *args))
    monkeypatch.setattr(undo, "excel_fingerprint", lambda *args: reading(fingerprint, *args))
    if unreadable_snapshot:
        monkeypatch.setattr(undo, "_finish_excel", lambda *args: reading(finish, *args))
    with pytest.raises(ToolError, match="applied=1, remaining=1, unverified=1.*force=true"):
        clean(o, cells="A1:A3", preview=False)
    assert history(o)[-1].undoable and history(o)[-1].backup_sheets and history(o)[-1].force_reason
    busy = False
    with pytest.raises(ToolError, match="force=true"):
        o.call("office_undo", file=o.wb.Name)
    o.call("office_undo", file=o.wb.Name, force=True)
    assert o.ws.data == before


def test_stage1_readonly_registration(office, monkeypatch):
    from office_live import word_core, word_tables

    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, mode="readonly"))
    for module, name in ((excel_analysis, "excel_compare_ranges"), (excel_core, "excel_clean_text"),
                         (word_core, "word_compare_documents"), (word_tables, "word_sort_table")):
        info = registry.CATALOG[name]
        registry.office_tool(info.group, info.kind)(getattr(module, name))
        assert registry.CATALOG[name].registered is (name == "excel_compare_ranges")


def test_compare_key_only_lists_follow_max_differences(office):
    o = office
    o.ws.Range("A1:A6").Value = (("id",), (1,), (2,), (3,), (4,), (5,))
    o.ws.Range("B1:B2").Value = (("id",), (9,))
    result = compare(o, cells_a="A1:A6", cells_b="B1:B2", match="key", key_column="id", max_differences=2)
    assert result["only_in_a"] == ["A2", "A3"] and result["only_truncated"]
    assert result["summary"]["rows_only_in_a"] == 5


def test_clean_skips_spilled_cells(office):
    o = office
    o.ws.data = {(1, 1): " a ", (2, 1): " b "}
    o.ws.spills = {(2, 1)}
    result = clean(o, preview=False, cells="A1:A2")
    assert result["changed"] == 1 and result["skipped_formulas"] == 1
    assert o.ws.data[1, 1] == "a" and o.ws.data[2, 1] == " b "


def test_clean_works_without_undo_when_owner_disabled_it(office, monkeypatch):
    o = office
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo=False))
    o.ws.data = {(1, 1): " a "}
    assert clean(o, preview=False, cells="A1")["applied"] == 1
    assert o.ws.data[1, 1] == "a" and not history(o)
