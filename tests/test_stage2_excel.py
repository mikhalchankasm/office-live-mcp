"""Stage 2: in-memory COM only, including real tool wrappers and undo snapshots."""

import copy
import json
from dataclasses import asdict, replace
from types import SimpleNamespace as NS

import pytest
import pywintypes

from office_live import config, excel_core, excel_format, formula_trace, journal, registry, undo
from office_live.errors import AppBusyError, ToolError
from office_live.util import DATE_FORMATS, excel_date_serial
from tests.history_fakes import Collection, Range, Sheet, fake_office  # noqa: F401


def split(o, **kw):
    return o.call("excel_split_column", **{"workbook": o.wb.Name, "sheet": "Data", "cells": "A1:A2", "delimiter": "pipe", **kw})


def protect(o, **kw):
    return o.call("excel_protection", **{"workbook": o.wb.Name, "sheet": "Data", **kw})


def trace(o, **kw):
    return o.call("excel_trace_formula", **{"workbook": o.wb.Name, "sheet": "Data", "cell": "B1", **kw})


def dates(o, **kw):
    return o.call("excel_clean_text", **{"workbook": o.wb.Name, "sheet": "Data", "cells": "A1:A2", "operations": ["text_to_date"],
                                        "date_format": "DD.MM.YYYY", **kw})


def history(o):
    return undo.STACKS.get(undo.stack_key("workbook", o.excel, o.wb), [])


def test_split_preview_and_full_output_undo(office):
    o = office
    o.ws.Range("A1:A2").Value = (("007|=x|12,5",), (" text |1,234|TRUE",))
    before = copy.deepcopy(o.ws.data)
    result = split(o, types=["text", "text", "number"])
    assert result["parts"] == 3 and result["output_range"] == "A1:C2" and result["output_cells"] == 6
    assert o.ws.data == before and not history(o) and o.excel.Workbooks.Count == 1
    o.ws.Range("C1").NumberFormat = "@"
    result = split(o, types=["text", "number", "number"], preview=False)
    assert result["ambiguous_numbers"] == ["B2"] and result["converted_numbers"] == 1
    assert result["format_changed"] == ["C1"] and o.ws.Range("C1").Value2 == 12.5
    assert o.ws.Range("A1").Value2 == "007" and o.ws.Range("B1").Value2 == "=x" and not o.ws.Range("B1").HasFormula
    assert history(o)[-1].areas == [{"sheet": "Data", "address": "A1:C2"}]
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1:A2").Value == (("007|=x|12,5",), (" text |1,234|TRUE",))
    assert o.ws.Range("B1:C2").Value == ((None, None), (None, None)) and o.ws.Range("C1").NumberFormat == "@"


@pytest.mark.parametrize("args,source,expected", [
    ({"delimiter": "::", "max_parts": 2}, " a::b::c ", ["a", "b::c"]),
    ({"delimiter": "pipe", "consecutive_as_one": True}, "a|||b", ["a", "b"]),
    ({"delimiter": "pipe", "consecutive_as_one": False}, "a||b", ["a", "", "b"]),
    ({"delimiter": "", "widths": [2, 3]}, "ABCDEFGH", ["AB", "CDE", "FGH"]),
    ({"delimiter": "", "widths": [2, 3], "max_parts": 2}, "ABCDEFGH", ["AB", "CDEFGH"]),
    ({"max_parts": 1, "trim": False}, " a|b ", [" a|b "]),
    ({"delimiter": "comma"}, '"a,b",c', ['"a', 'b"', 'c']),
    ({"delimiter": "tab"}, "a\tb", ["a", "b"]),
])
def test_split_parser(office, args, source, expected):
    o = office
    o.ws.Range("A1").Value = source
    result = o.call("excel_split_column", **{"workbook": o.wb.Name, "sheet": "Data", "cells": "A1", "delimiter": "pipe", **args})
    assert result["examples"][0]["after"] == expected


@pytest.mark.parametrize("args,message", [
    ({"delimiter": ""}, "exactly one"), ({"widths": [1]}, "exactly one"), ({"delimiter": "", "widths": [0]}, "positive"),
    ({"delimiter": "", "widths": [True]}, "positive"), ({"delimiter": "", "widths": []}, "positive"),
    ({"max_parts": -1}, "max_parts"), ({"max_parts": 1.5}, "max_parts"), ({"types": ["date"]}, "types"),
    ({"types": ["text"] * 4}, "more entries"), ({"cells": ""}, "required"), ({"cells": "A:B"}, "one column"),
    ({"destination": "B1:B3"}, "one A1"), ({"destination": "Other!A1"}, "one A1"), ({"destination": "XFD1"}, "boundaries"),
    ({"destination": "B1048576"}, "boundaries"), ({"cells": "A1:A100001"}, "100000"),
])
def test_split_input_refusals_before_snapshot(office, monkeypatch, args, message):
    o = office
    o.ws.Range("A1:A2").Value = "a|b|c"
    before = copy.deepcopy(o.ws.data)
    monkeypatch.setattr(undo, "_snapshots", lambda *a, **kw: pytest.fail("preflight first"))
    with pytest.raises(ToolError, match=message):
        o.call("excel_split_column", **{"workbook": o.wb.Name, "sheet": "Data", "cells": "A1:A2", "delimiter": "pipe", "preview": False, **args})
    assert o.ws.data == before and not history(o)


@pytest.mark.parametrize("problem", ["occupied", "empty_formula", "source_formula", "nontext", "merged", "source_merge", "table", "pivot", "protected", "spill", "autosave"])
def test_split_state_refusals(office, monkeypatch, problem):
    o = office
    o.ws.Range("A1:A2").Value = "a|b|c"
    if problem == "occupied":
        o.ws.Range("C2").Value = ""
    elif problem == "empty_formula":
        o.ws.Range("C2").Formula = '=""'
        o.ws.results[2, 3] = None
    elif problem == "source_formula":
        o.ws.Range("A2").Formula = '=TEXT(1,"0")'
    elif problem == "nontext":
        o.ws.Range("A2").Value = 123
    elif problem in {"merged", "source_merge"}:
        o.ws.Range("C2" if problem == "merged" else "A2").MergeCells = True
    elif problem == "table":
        o.ws.ListObjects = Collection([NS(Name="T", Range=o.ws.Range("C2:D3"))])
    elif problem == "pivot":
        o.ws.pivots = Collection([NS(Name="P", TableRange2=o.ws.Range("C2:D3"))])
    elif problem == "protected":
        o.ws.ProtectContents = True
    elif problem == "spill":
        o.ws.spills = {(2, 3)}
    elif problem == "autosave":
        o.wb.AutoSaveOn = True
    before = copy.deepcopy(o.ws.data)
    monkeypatch.setattr(undo, "_snapshots", lambda *a, **kw: pytest.fail("preflight first"))
    with pytest.raises(ToolError):
        split(o, preview=False)
    assert o.ws.data == before and not history(o)


def test_split_output_limit_and_explicit_destination(office):
    o = office
    o.ws.Range("A1:A2").Value = "a|b"
    with pytest.raises(ToolError, match="not empty"):
        split(o, destination="A1")
    assert split(o, destination="D7", preview=False)["output_range"] == "D7:E8"
    assert o.ws.Range("A1").Value == "a|b"
    o.ws.Range("E8").Value = "user"
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.wb.Name)
    o.call("office_undo", file=o.wb.Name, force=True)
    o.ws.Range("A1").Value = "|" * 100
    with pytest.raises(ToolError, match="100000"):
        o.call("excel_split_column", workbook=o.wb.Name, sheet="Data", cells="A1:A1000", delimiter="pipe")


def test_split_snapshot_failure_and_partial_failure(office, monkeypatch):
    o = office
    o.ws.Range("A1:A2").Value = "a|12"
    with monkeypatch.context() as patch:
        patch.setattr(undo, "_snapshots", lambda *a, **kw: (_ for _ in ()).throw(ValueError("snapshot failed")))
        with pytest.raises(ToolError, match="prepare office_undo"):
            split(o, preview=False)
    assign = Range._assign

    def failing(self, value, formula=False):
        if self.Address == "B1:B2":
            raise pywintypes.com_error(-2147352567, "write failed", None, None)
        return assign(self, value, formula)

    monkeypatch.setattr(Range, "_assign", failing)
    with pytest.raises(ToolError, match="applied=2.*office_undo will restore"):
        split(o, preview=False)
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1:A2").Value == (("a|12",), ("a|12",))


@pytest.mark.parametrize("fmt", DATE_FORMATS)
def test_all_date_formats_strict(fmt):
    import datetime

    value = datetime.datetime(2024, 2, 29, 17, 23, 45)
    text = value.strftime(DATE_FORMATS[fmt])
    assert excel_date_serial(text, fmt) is not None
    assert excel_date_serial(text + " ", fmt) is None
    assert excel_date_serial(text.replace("2024", "24"), fmt) is None


@pytest.mark.parametrize("text,expected", [("01.01.1900", 1), ("28.02.1900", 59), ("01.03.1900", 61),
                                          ("29.02.1900", None), ("31.02.2024", None), ("1.01.2024", None), ("31.12.1899", None)])
def test_date_serial_1900(text, expected):
    assert excel_date_serial(text, "DD.MM.YYYY") == expected


@pytest.mark.parametrize("date1904", [False, True])
def test_dates_preview_write_format_undo_and_skips(office, date1904):
    o = office
    o.wb.Date1904 = date1904
    o.ws.Range("A1:A4").Value = (("29.02.2024 12:30:15",), ("31.02.2024 12:30:15",), ("=A1",), (12,))
    o.ws.Range("A1").NumberFormat = "@"
    args = {"cells": "A1:A4", "date_format": "DD.MM.YYYY HH:MM:SS"}
    assert dates(o, **args)["changed"] == 1 and not history(o)
    result = dates(o, preview=False, **args)
    assert result["skipped_not_date"] == 1 and result["not_date"] == ["A2"] and result["skipped_formulas"] == 1
    assert o.ws.Range("A1").Value2 == pytest.approx(excel_date_serial("29.02.2024 12:30:15", args["date_format"], date1904))
    assert o.ws.Range("A1").NumberFormat == "dd.mm.yyyy hh:mm:ss"
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1").Value2 == "29.02.2024 12:30:15" and o.ws.Range("A1").NumberFormat == "@"


@pytest.mark.parametrize("args", [{"date_format": ""}, {"date_format": "DD.MM.YY"}, {"operations": ["text_to_date", "text_to_number"]},
                                  {"operations": ["trim"]}])
def test_date_argument_refusals(office, monkeypatch, args):
    monkeypatch.setattr(undo, "_snapshots", lambda *a, **kw: pytest.fail("preflight first"))
    with pytest.raises(ToolError):
        dates(office, preview=False, **args)


def test_invalid_dates_not_trimmed_and_address_cap(office):
    o = office
    o.ws.Range("A1:A60").Value = " invalid "
    result = dates(o, cells="A1:A60", operations=["trim", "text_to_date"], preview=False)
    assert result["skipped_not_date"] == 60 and len(result["not_date"]) == 50 and not history(o)
    assert o.ws.Range("A1").Value == " invalid "


def test_protect_sheet_inverse_locked_flags_and_conflict(office):
    o = office
    o.ws.Range("A2").Locked = False
    result = protect(o, action="protect", unlocked_cells=["A1:A2"], allow=["sort", "filter", "edit_objects"])
    assert result["undo"] == "available" and o.ws.ProtectContents and not o.ws.ProtectDrawingObjects
    assert o.ws.Protection.AllowSorting and o.ws.Protection.AllowFiltering and not o.ws.Range("A1").Locked
    o.ws.Range("A2").Locked = True
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.wb.Name)
    o.call("office_undo", file=o.wb.Name, force=True)
    assert not o.ws.ProtectContents and o.ws.Range("A1").Locked and not o.ws.Range("A2").Locked


@pytest.mark.parametrize("scope", ["sheet", "workbook"])
def test_protection_unprotect_inverse(office, scope):
    o = office
    args = {"allow": list(excel_format.PROTECTION_ALLOW) + ["edit_scenarios"]} if scope == "sheet" else {}
    protect(o, action="protect", scope=scope, **args)
    before = protect(o, scope=scope)["state"]
    result = protect(o, action="unprotect", scope=scope)
    assert result["undo"] == "available"
    o.call("office_undo", file=o.wb.Name)
    assert protect(o, scope=scope)["state"] == before
    o.call("office_undo", file=o.wb.Name)
    assert not excel_format._protected(protect(o, scope=scope)["state"], scope)


@pytest.mark.parametrize("args", [{"scope": "bad"}, {"action": "bad"}, {"allow": ["bad"]}, {"allow": ["sort", "sort"]},
                                  {"scope": "workbook", "allow": ["sort"]}, {"unlocked_cells": "A1"},
                                  {"unlocked_cells": [""]}, {"unlocked_cells": ["A:A"]}, {"unlocked_cells": ["Else!A1"]}])
def test_protection_refuses_before_changes(office, args):
    with pytest.raises(ToolError):
        protect(office, **{"action": "protect", **args})
    assert not history(office) and not office.ws.protection_calls


def test_password_redacted_everywhere_and_wrong_password(office, monkeypatch):
    o = office
    secret = "small-PASS-Я"
    log = o.tmp / "audit.jsonl"
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, audit_log=log, audit_content=True))
    o.call("office_journal", workbook=o.wb.Name, action="enable_sheet")
    result = protect(o, action="protect", password=secret, unlocked_cells=["A1"])
    assert "not available" in result["undo"]
    assert secret not in json.dumps([asdict(e) for e in history(o)], ensure_ascii=False)
    assert secret not in repr(o.wb.Sheets("Лог").data)
    assert secret not in log.read_text("utf-8") and "<redacted>" in log.read_text("utf-8")
    assert secret not in journal.journal_path(o.wb.Name).read_text("utf-8")
    prior = copy.deepcopy(history(o))
    wrong = "wrong-PASS"
    with pytest.raises(ToolError, match="Неверный пароль") as error:
        protect(o, action="unprotect", password=wrong)
    assert wrong not in str(error.value) and wrong not in log.read_text("utf-8")
    assert history(o) == prior and o.ws.ProtectContents
    protect(o, action="unprotect", password=secret)
    assert not history(o)[-1].undoable and not o.ws.ProtectContents


@pytest.mark.parametrize("error_class", [ToolError, ValueError, RuntimeError])
def test_generic_password_redaction_positional_errors_and_journal_overrides(office, monkeypatch, capsys, error_class):
    o = office
    secret = "short-secret"
    settings = replace(config.SETTINGS, audit_log=o.tmp / "audit.jsonl", audit_content=True)
    monkeypatch.setattr(config, "SETTINGS", settings)
    captured = {}
    monkeypatch.setattr(registry, "mcp", NS(add_tool=lambda fn, name, **kw: captured.setdefault(name, fn)))

    @registry.office_tool("excel_format", "write")
    def password_probe(oldPassword: str):
        raise error_class("Echo: " + oldPassword)

    with pytest.raises(ToolError) as error:
        captured["password_probe"](secret)
    assert secret not in str(error.value) and secret not in settings.audit_log.read_text("utf-8")
    assert secret not in capsys.readouterr().err
    args = {"oldPassword": secret, "nested": {"PASSWORD": secret}}
    assert secret not in str(journal.describe(args))
    journal.append(["workbook:X"], "probe", args, error=secret, where=secret, summary=secret)
    assert secret not in journal.journal_path("X").read_text("utf-8")


def test_protection_failed_call_restores_locked_and_prior_history(office, monkeypatch):
    o = office
    o.ws.Range("A1").Value = "old"
    o.call("excel_write_range", workbook=o.wb.Name, sheet="Data", cells="A1", values="new")
    prior = copy.deepcopy(history(o))
    monkeypatch.setattr(Sheet, "Protect", lambda *a: (_ for _ in ()).throw(pywintypes.com_error(-2147352567, "fail", None, None)))
    with pytest.raises(ToolError):
        protect(o, action="protect", unlocked_cells=["A1"])
    assert o.ws.Range("A1").Locked and history(o) == prior


def test_trace_local_cross_sheet_names_tables_and_no_ui(office, monkeypatch):
    o = office
    other = o.wb.Worksheets.Add(None, o.ws)
    other.Name = "Лист 1"
    o.ws.Activate()
    o.ws.Range("B1").Formula = "=SUM(A1,'Лист 1'!A1:A2,Input,T[Col])+12"
    o.ws.direct["precedents", "B1"] = ["A1"]
    nm = o.wb.Names.Add("Input", "='Лист 1'!C1")
    nm.RefersToRange = other.Range("C1")
    other.ListObjects = Collection([NS(Name="T", ListColumns=Collection([NS(Name="Col", DataBodyRange=other.Range("D1:D2"))]))])
    before = (o.excel.ActiveSheet, o.excel.Selection, copy.deepcopy(o.ws.data))
    monkeypatch.setattr(Sheet, "Activate", lambda *a: pytest.fail("no activation"))
    result = trace(o)
    labels = {n["address"] for n in result["nodes"]}
    assert labels == {"Data!B1", "Data!A1", "'Лист 1'!A1", "'Лист 1'!A2", "'Лист 1'!C1", "'Лист 1'!D1", "'Лист 1'!D2"}
    assert len(result["edges"]) == 6
    assert "inactive_sheet_com_unavailable" in result["incomplete_reasons"]
    assert (o.excel.ActiveSheet, o.excel.Selection, o.ws.data) == before and not history(o)


def test_trace_dependents_contains_ranges_and_ignores_strings(office):
    o = office
    other = o.wb.Worksheets.Add(None, o.ws)
    other.Name = "Other"
    other.Range("B1:B3").Formula = (("=SUM(Data!A1:A5)",), ('="Data!A2"',), ("=Data!A5",))
    o.ws.Range("B1").Formula = "=A2"
    o.ws.Activate()
    result = trace(o, cell="A2", direction="dependents")
    assert {n["address"] for n in result["nodes"]} == {"Data!A2", "Data!B1", "Other!B1"}
    assert {tuple(e.values()) for e in result["edges"]} == {("Data!A2", "Data!B1"), ("Data!A2", "Other!B1")}


@pytest.mark.parametrize("formula,reason", [
    ('=INDIRECT("A1")', "indirect"), ("=OFFSET(A1,1,1)", "offset"), ("=INDEX(A1:A4,2)", "index"),
    ("='[Other.xlsx]Data'!A1", "external_reference"), ("=Data:Other!A1", "3d_reference"),
    ("=A1#", "spill_reference"), ("=LET(x,A1,x)", "let"), ("=LAMBDA(x,x+1)(A1)", "lambda"),
    ("=FormulaName", "name_not_range"), ("=T[[#Totals],[Col]]", "unsupported_structured_reference"),
    ("=MissingName", "unresolved_name"),
])
def test_trace_incomplete_is_explicit(office, formula, reason):
    o = office
    o.wb.Names.Add("FormulaName", "=1+2")
    o.ws.Range("B1").Formula = formula
    result = trace(o)
    assert not result["complete"] and reason in result["incomplete_reasons"]


def test_trace_depth_nodes_cycle_and_scan_limits(office, monkeypatch):
    o = office
    o.ws.Range("B1").Formula = "=A1"
    assert trace(o)["complete"]
    o.ws.Range("A1").Formula = "=C1"
    assert "depth_limit" in trace(o)["incomplete_reasons"]
    assert trace(o, depth=2)["complete"]
    o.ws.Range("C1").Formula = "=B1"
    assert len(trace(o, depth=5)["nodes"]) == 3
    o.ws.Range("B1").Formula = "=SUM(A1:A1000)"
    result = trace(o, max_nodes=3)
    assert len(result["nodes"]) == 3 and result["truncated"] and "node_limit" in result["incomplete_reasons"]
    monkeypatch.setattr(formula_trace, "SCAN_LIMIT", 1)
    assert "scan_limit" in trace(o, direction="dependents")["incomplete_reasons"]


@pytest.mark.parametrize("args", [{"cell": ""}, {"cell": "A1:B2"}, {"direction": "bad"}, {"depth": 0}, {"depth": 6},
                                  {"max_nodes": 0}, {"max_nodes": 2001}, {"max_nodes": True}])
def test_trace_input_refusals(office, args):
    with pytest.raises(ToolError):
        trace(office, **args)


def test_readonly_registration_and_status(office, monkeypatch):
    o = office
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, mode="readonly"))
    assert protect(o)["action"] == "status"
    assert trace(o)["complete"]
    with pytest.raises(ToolError, match="Read-only"):
        protect(o, action="protect")
    info = registry.CATALOG["excel_split_column"]
    registry.office_tool(info.group, info.kind)(excel_core.excel_split_column)
    assert not registry.CATALOG["excel_split_column"].registered


@pytest.mark.parametrize("tool", ["split", "dates", "protection"])
@pytest.mark.parametrize("problem", ["readonly_book", "autosave", "outside", "strict"])
def test_stage2_target_guards(office, monkeypatch, tool, problem):
    o = office
    o.ws.Range("A1:A2").Value = "01.01.2024"
    if problem == "readonly_book":
        o.wb.ReadOnly = True
    elif problem == "autosave":
        o.wb.AutoSaveOn = True
    elif problem == "outside":
        o.wb.Path = str(o.tmp / "outside")
        o.wb.FullName = str(o.tmp / "outside" / "Book.xlsx")
        monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, allowed_dirs=(str(o.tmp / "allowed"),)))
    else:
        monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, strict_target=True))
    kwargs = {"workbook": ""} if problem == "strict" else {}
    with pytest.raises(ToolError):
        if tool == "split":
            split(o, preview=False, **kwargs)
        elif tool == "dates":
            dates(o, preview=False, **kwargs)
        else:
            protect(o, action="protect", **kwargs)
    assert not history(o) and not o.ws.protection_calls and o.ws.Range("A1").Value == "01.01.2024"


def test_split_preview_no_audit_journal_events_and_coercion_repair(office, monkeypatch):
    o = office
    o.ws.Range("A1").Value = "abc|def"
    with monkeypatch.context() as patch:
        for module, name in ((journal, "append"), (journal, "append_sheet"), (registry, "_audit"), (undo, "_snapshots")):
            patch.setattr(module, name, lambda *a, **kw: pytest.fail("preview side effect"))
        assert split(o, cells="A:A")["rows"] == 1 and o.excel.EnableEvents
    assign = Range._assign

    def coerce(self, value, formula=False):
        assign(self, value, formula)
        if self.Address == "B1" and value == (("def",),):
            self.Worksheet.data[1, 2] = 1

    monkeypatch.setattr(Range, "_assign", coerce)
    assert split(o, cells="A1", preview=False)["rewritten_as_text"] == ["B1"]
    assert o.ws.Range("B1").Value == "def"


@pytest.mark.parametrize("tool", ["split", "protection"])
def test_stage2_undo_disabled_explicitly(office, monkeypatch, tool):
    o = office
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo=False))
    o.ws.Range("A1").Value = "a|b"
    result = split(o, preview=False) if tool == "split" else protect(o, action="protect", unlocked_cells=["A1"])
    assert "undo" not in result and not history(o)


def test_protection_status_and_validation_leave_events_and_history(office):
    o = office
    assert protect(o)["state"]["contents"] is False and o.excel.EnableEvents
    assert not protect(o, action="unprotect")["changed"] and not history(o)
    protect(o, action="protect")
    prior = copy.deepcopy(history(o))
    with pytest.raises(ToolError, match="already protected"):
        protect(o, action="protect")
    assert history(o) == prior


def test_protection_partial_mutation_retains_inverse(office, monkeypatch):
    o = office
    original = Sheet.Protect

    def fail_after(self, *args):
        original(self, *args)
        raise pywintypes.com_error(-2147352567, "failure after protecting", None, None)

    monkeypatch.setattr(Sheet, "Protect", fail_after)
    with pytest.raises(ToolError, match="office_undo will restore"):
        protect(o, action="protect", unlocked_cells=["A1:A2"])
    assert o.ws.ProtectContents and history(o)[-1].undoable
    o.call("office_undo", file=o.wb.Name)
    assert not o.ws.ProtectContents and o.ws.Range("A1:A2").Locked


def test_dates_partial_failure_and_number_format_conflict(office, monkeypatch):
    o = office
    o.ws.Range("A1:A3").Value = (("01.01.2024",), ("skip",), ("02.01.2024",))
    assign = Range._assign

    def fail(self, value, formula=False):
        if self.Address == "A3":
            raise pywintypes.com_error(-2147352567, "failure", None, None)
        assign(self, value, formula)

    with monkeypatch.context() as patch:
        patch.setattr(Range, "_assign", fail)
        with pytest.raises(ToolError, match="applied=1.*office_undo will restore"):
            dates(o, cells="A1:A3", preview=False)
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1").Value == "01.01.2024" and o.ws.Range("A3").NumberFormat == "General"
    dates(o, preview=False, date_number_format="yyyy-mm-dd")
    o.ws.Range("A2").NumberFormat = "@"
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.wb.Name)


def test_trace_com_error_and_name_formula_that_evaluates_to_range(office, monkeypatch):
    o = office
    o.ws.Range("B1").Formula = "=Dyn"
    nm = o.wb.Names.Add("Dyn", "=OFFSET(Data!A1,0,0)")
    nm.RefersToRange = o.ws.Range("A1")
    assert "name_not_range" in trace(o)["incomplete_reasons"]

    def no_cells(self):
        raise pywintypes.com_error(-2147352567, "Exception", (0, "Excel", "No cells found", None, 0, -2146827284), None)

    monkeypatch.setattr(Range, "DirectPrecedents", property(no_cells))
    result = trace(o, cell="C1")
    assert result["edges"] == [] and "direct_references_unavailable_or_empty" in result["incomplete_reasons"]


def test_split_multi_area_and_undo_limits(office, monkeypatch):
    o = office
    o.ws.Range("A1:A2").Value = "a|b"
    original = Range.__init__

    def multi(self, *args):
        original(self, *args)
        self.Areas = Collection([self, self])

    with monkeypatch.context() as patch:
        patch.setattr(Range, "__init__", multi)
        with pytest.raises(ToolError, match="Multi-area"):
            split(o, preview=False)
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo_max_cells=3))
    with pytest.raises(ToolError, match="MAX_CELLS"):
        split(o, preview=False)
    assert o.ws.Range("A1").Value == "a|b" and not history(o)


def test_protection_merged_unlock_and_snapshot_limits(office, monkeypatch):
    o = office
    o.ws.Range("B1").MergeCells = True
    with pytest.raises(ToolError, match="merged"):
        protect(o, action="protect", unlocked_cells=["A1", "B1"])
    assert o.ws.Range("A1").Locked and not history(o)
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo_max_cells=1))
    with pytest.raises(ToolError, match="MAX_CELLS"):
        protect(o, action="protect", unlocked_cells=["A1:A2"])
    assert not o.ws.ProtectContents and not history(o)


def test_trace_external_name_and_reference_work_budget(office, monkeypatch):
    o = office
    o.wb.Names.Add("External", "='[Other.xlsx]Data'!$A$1")
    o.ws.Range("B1").Formula = "=External"
    assert "external_reference" in trace(o)["incomplete_reasons"]
    worker = formula_trace.Trace(o.excel, o.wb, "dependents", 1, 10)
    worker.scanned = [(o.ws.Range("B1"), [o.ws.Range("A1")])]
    worker.reference_checks = formula_trace.SCAN_LIMIT * 4
    result = worker.run(o.ws.Range("A1"))
    assert result["truncated"] and "reference_check_limit" in result["incomplete_reasons"]


def test_trace_called_lambda_name_does_not_claim_complete(office):
    o = office
    o.wb.Names.Add("Compute", "=LAMBDA(x,x+Data!C1)")
    o.ws.Range("B1").Formula = "=Compute(A1)"
    result = trace(o)
    assert not result["complete"] and "named_function" in result["incomplete_reasons"]


def test_protection_busy_after_mutation_keeps_snapshot_for_force(office, monkeypatch):
    o = office
    busy = False
    original, state = Sheet.Protect, excel_format._protection_state

    def fail(self, *args):
        nonlocal busy
        original(self, *args)
        busy = True
        raise AppBusyError("Office busy")

    def read(*args):
        if busy:
            raise AppBusyError("Office busy")
        return state(*args)

    monkeypatch.setattr(Sheet, "Protect", fail)
    monkeypatch.setattr(excel_format, "_protection_state", read)
    with pytest.raises(ToolError, match="force=true"):
        protect(o, action="protect", unlocked_cells=["A1"])
    assert history(o)[-1].undoable and history(o)[-1].force_reason
    busy = False
    o.call("office_undo", file=o.wb.Name, force=True)
    assert not o.ws.ProtectContents and o.ws.Range("A1").Locked
