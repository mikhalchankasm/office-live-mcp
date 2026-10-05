"""Регрессии шестого раунда: только fake COM, Win32 подделки и tmp_path."""

import ctypes
import os
import uuid
from types import SimpleNamespace as NS

import pytest

from office_live import com, undo
from office_live.errors import ToolError
from tests.history_fakes import Collection, fake_office  # noqa: F401


def test_value_undo_preserves_later_column_width_and_row_height(office):
    o = office
    o.call("excel_write_range", workbook=o.wb.Name, sheet="Data", cells="A1", values=1)
    o.ws.Columns(1).ColumnWidth = 42
    o.ws.Rows(1).RowHeight = 37
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1").Value is None
    assert o.ws.Columns(1).ColumnWidth == 42 and o.ws.Rows(1).RowHeight == 37


def test_value_undo_preserves_later_cell_formatting(office):
    o = office
    o.call("excel_write_range", workbook=o.wb.Name, sheet="Data", cells="A1", values=1)
    o.ws.Range("A1").Font.Bold = True
    o.ws.Range("A1").NumberFormat = "0.00"
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1").Value is None
    assert o.ws.Range("A1").Font.Bold and o.ws.Range("A1").NumberFormat == "0.00"


def test_format_undo_changes_only_supplied_property_and_checks_it(office):
    o = office
    o.call("excel_format_range", workbook=o.wb.Name, sheet="Data", cells="A1:B1", bold=True)
    o.ws.Range("A1").Value = "user"
    o.ws.Range("A1").Font.Italic = True
    o.ws.Range("A1").Font.Bold = False
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.wb.Name)
    o.call("office_undo", file=o.wb.Name, force=True)
    assert o.ws.Range("A1").Value == "user" and o.ws.Range("A1").Font.Italic
    assert o.ws.Range("A1:B1").Font.Bold is False


def test_mixed_format_property_is_restored_per_cell(office):
    o = office
    o.ws.Range("B1").Font.Bold = True
    o.call("excel_format_range", workbook=o.wb.Name, sheet="Data", cells="A1:B1", bold=False)
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1").Font.Bold is False and o.ws.Range("B1").Font.Bold is True


@pytest.mark.parametrize("options", [{"font_color": "red"}, {"fill_color": "yellow"}, {"borders": "all"}])
def test_color_and_border_undo_preserves_later_user_content(office, options):
    o = office
    original = undo._format_state(o.ws.Range("A1"))
    result = o.call("excel_format_range", workbook=o.wb.Name, sheet="Data", cells="A1", **options)
    assert result["undo"] == "available"
    o.ws.Range("A1").Value = "later user text"
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1").Value == "later user text" and undo._format_state(o.ws.Range("A1")) == original


def test_full_copy_undo_checks_named_style_even_when_looks_are_unchanged(office):
    o = office
    o.call("excel_copy_range", workbook=o.wb.Name, sheet="Data", source="B1", dest_cell="A1")
    o.ws.Range("A1").Style = "User style with same appearance"
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.wb.Name)


@pytest.mark.parametrize("member,cell_type,attribute,old,new", [
    ("Comment", -4144, "Visible", False, True),
    ("Validation", -4174, "Formula1", "=Lists!A1:A3", "=Lists!B1:B3"),
])
def test_full_snapshot_checks_changed_comment_and_validation(office, monkeypatch, member, cell_type, attribute, old, new):
    from tests.history_fakes import Range

    o = office
    value = NS(Text=lambda: "note", Author="user", Visible=False, Type=3, AlertStyle=1, Operator=1,
               Formula1="=Lists!A1:A3", Formula2=None, IgnoreBlank=True, InCellDropdown=True,
               ShowError=True, ErrorTitle="", ErrorMessage="", ShowInput=False, InputTitle="", InputMessage="")
    cell = o.ws.Range("A1")
    setattr(cell, member, value)
    original = Range.SpecialCells

    def special(rng, which, /):
        if which == cell_type:
            return NS(Cells=[cell])
        return original(rng, which)

    monkeypatch.setattr(Range, "SpecialCells", special)
    monkeypatch.setattr(o.excel, "Intersect", lambda rng, selected: selected, raising=False)
    entry = undo.Entry("excel_copy_range", "", "workbook", areas=[{"sheet": "Data", "address": "A1"}])
    setattr(value, attribute, old)
    before = undo.excel_fingerprint(o.excel, o.wb, entry)
    setattr(value, attribute, new)
    assert undo.excel_fingerprint(o.excel, o.wb, entry) != before


def test_single_cell_metadata_scan_intersects_before_enumeration(office, monkeypatch):
    from tests.history_fakes import Range

    o = office
    outside = NS()  # no Cells: scanning the whole sheet instead of Intersect would fail
    monkeypatch.setattr(Range, "SpecialCells", lambda rng, kind: outside)
    calls = []
    monkeypatch.setattr(o.excel, "Intersect", lambda rng, selected: calls.append(selected), raising=False)
    assert undo._metadata_state(o.ws.Range("A1"), "validation") == []
    assert calls == [outside]


@pytest.mark.parametrize("axis,property_name", [("Columns", "ColumnWidth"), ("Rows", "RowHeight")])
def test_captured_dimensions_participate_in_range_fingerprint(office, axis, property_name):
    o = office
    entry = undo.Entry("excel_delete_rows_columns", "", "workbook")
    com.run_com(lambda: undo._snapshots(o.excel, o.wb, entry, [o.ws.Range("A1")], columns=axis == "Columns", rows=axis == "Rows"))
    before = undo.excel_fingerprint(o.excel, o.wb, entry)
    setattr(getattr(o.ws, axis)(1), property_name, 42)
    assert undo.excel_fingerprint(o.excel, o.wb, entry) != before


@pytest.mark.parametrize("address,widths,heights", [("A2:F2", 0, 1), ("B1:B8", 1, 0)])
def test_full_line_copy_snapshots_only_dimensions_it_can_change(office, address, widths, heights):
    o = office
    o.ws.Rows.Count, o.ws.Columns.Count = 8, 6
    entry = undo.Entry("excel_copy_range", "", "workbook")
    com.run_com(lambda: undo._snapshots(o.excel, o.wb, entry, [o.ws.Range(address)]))
    assert len(entry.ops[0]["widths"]) == widths and len(entry.ops[0]["heights"]) == heights


@pytest.mark.parametrize("kind", ["rows", "columns", "both"])
def test_autofit_dimensions_are_fingerprinted(office, monkeypatch, kind):
    o = office
    entry = undo.Entry("excel_set_dimensions", "", "workbook")
    undo._dimensions_plan(o.excel, o.wb, entry, {"sheet": "Data", "cells": "A1:B2", "autofit": kind})
    before = undo.excel_fingerprint(o.excel, o.wb, entry)
    if kind == "columns":
        o.ws.Columns(1).ColumnWidth = 30
    else:
        o.ws.Rows(1).RowHeight = 30
    assert undo.excel_fingerprint(o.excel, o.wb, entry) != before


@pytest.mark.parametrize("creator", ["add", "copy", "bridge"])
@pytest.mark.parametrize("change", ["format", "shape", "table", "condition", "content", "extent"])
def test_created_sheet_user_changes_block_undo(office, creator, change):
    o = office
    if creator == "add":
        o.call("excel_add_worksheet", workbook=o.wb.Name, name="New")
    elif creator == "copy":
        o.call("excel_manage_sheet", workbook=o.wb.Name, sheet="Data", action="copy", new_name="New")
    else:
        o.doc.Paragraphs = Collection([])
        o.call("bridge_word_text_to_excel", document=o.doc.Name, workbook=o.wb.Name, sheet="New", create_sheet=True)
    ws = o.wb.Sheets("New")
    if change == "format":
        ws.Range("A1").Interior.Color = 255
    elif change == "shape":
        ws.Shapes.items.append(NS(Name="Picture"))
    elif change == "table":
        ws.ListObjects.items.append(NS(Name="Table1"))
    elif change == "condition":
        ws.conditions.items.append(NS())
    else:
        ws.Range("A1" if change == "content" else "Z20").Value = "user"
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.wb.Name)
    assert ws in o.wb.Sheets.items
    o.call("office_undo", file=o.wb.Name, force=True)
    assert ws not in o.wb.Sheets.items


def foreign_backup(o, pid, session=None):
    wb = o.excel.Workbooks.Add()
    wb.Names.Add(undo.MARKER, f'="{pid}:{session or uuid.uuid4().hex}:{uuid.uuid4().hex}"', False)
    wb.Windows(1).Visible = False
    o.wb.Activate()
    return wb


def test_backup_reuses_only_this_process_session_and_closes_dead_owners(office, monkeypatch):
    o = office
    dead = [foreign_backup(o, 100 + i) for i in range(2)]
    live = foreign_backup(o, 200)
    same_pid = foreign_backup(o, os.getpid())  # иной session UUID, даже при совпадении PID
    monkeypatch.setattr(undo, "_process_alive", lambda pid: pid not in {100, 101})
    calls = []
    for wb in dead:
        original = wb.Close

        def close(save, /, original=original):
            assert not o.excel.DisplayAlerts and not o.excel.EnableEvents
            calls.append(save)
            original(save)

        monkeypatch.setattr(wb, "Close", close)
    owned = com.run_com(lambda: undo._backup(o.excel))
    assert all(wb not in o.excel.Workbooks.items for wb in dead) and calls == [False, False]
    assert live in o.excel.Workbooks.items and same_pid in o.excel.Workbooks.items
    assert str(os.getpid()) + ":" + undo.SESSION in owned.Names(undo.MARKER).RefersTo
    assert com.run_com(lambda: undo._backup(o.excel)) is owned
    assert o.excel.DisplayAlerts and o.excel.EnableEvents and o.excel.ActiveWorkbook is o.wb


def test_restarted_session_never_reuses_old_entry_backup(office, monkeypatch):
    o = office
    o.call("excel_write_range", workbook=o.wb.Name, sheet="Data", cells="A1", values=1)
    old = o.excel.Workbooks(2)
    old.Names(undo.MARKER).RefersTo = '="101:' + uuid.uuid4().hex + ':' + uuid.uuid4().hex + '"'
    entry = undo.STACKS[undo.stack_key("workbook", o.excel, o.wb)][-1]
    entry.backup_token = old.Names(undo.MARKER).RefersTo
    monkeypatch.setattr(undo, "_process_alive", lambda pid: False)
    history = o.call("office_undo", file=o.wb.Name, action="history")["history"]
    assert old not in o.excel.Workbooks.items and not history[0]["undoable"]
    with pytest.raises(ToolError, match="closed"):
        o.call("office_undo", file=o.wb.Name, force=True)


def test_dead_backup_close_failure_restores_alerts_and_events(office, monkeypatch):
    o = office
    old = foreign_backup(o, 101)
    monkeypatch.setattr(undo, "_process_alive", lambda pid: False)
    monkeypatch.setattr(old, "Close", lambda save: (_ for _ in ()).throw(ValueError("close failed")))
    with pytest.raises(ToolError, match="close failed"):
        com.run_com(lambda: undo._backup(o.excel))
    assert o.excel.DisplayAlerts and o.excel.EnableEvents


@pytest.mark.parametrize("handle,error,code,query_ok,expected", [
    (12, 0, 259, True, True), (12, 0, 0, True, False), (12, 0, 0, False, True),
    (0, 87, 0, True, False), (0, 5, 0, True, True),
])
def test_windows_liveness_uses_query_api_and_closes_handle(monkeypatch, handle, error, code, query_ok, expected):
    calls = []

    def open_process(access, inherit, pid):
        assert (access, inherit, pid) == (0x1000, False, 123)
        return handle

    def exit_code(process, value):
        assert process == handle
        value._obj.value = code
        return query_ok

    def close(process):
        calls.append(process)

    kernel = NS(OpenProcess=open_process, GetExitCodeProcess=exit_code, CloseHandle=close)
    monkeypatch.setattr(ctypes, "WinDLL", lambda *a, **kw: kernel)
    monkeypatch.setattr(ctypes, "get_last_error", lambda: error)
    monkeypatch.setattr(os, "kill", lambda *a: pytest.fail("os.kill must never be used"))
    assert undo._process_alive(123) is expected
    assert calls == ([handle] if handle else [])
