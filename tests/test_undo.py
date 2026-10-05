"""Отмена на подделках COM, включая настоящие обёртки MCP и защитные проверки."""

from dataclasses import replace
from types import SimpleNamespace as NS

import pytest

from office_live import com, config, undo, xl_common as xl
from office_live.errors import ToolError
from tests.history_fakes import fake_office  # noqa: F401 — общая фикстура office


def stack(o):
    return undo.STACKS.get(undo.stack_key("workbook", o.excel, o.wb), [])


def write(o, values, cells="A1"):
    return o.call("excel_write_range", workbook=o.wb.Name, sheet="Data", cells=cells, values=values)


def revert(o, **kwargs):
    return o.call("office_undo", file=o.wb.Name, **kwargs)


def test_range_snapshot_and_restore_repairs_only_formulas(office):
    o = office
    o.ws.Range("A1:C1").Value = (("007", "=Sheet2!A1", "constant"),)
    o.ws.Columns(2).ColumnWidth = 24
    o.ws.Rows(1).RowHeight = 33
    result = write(o, [[1, 2, 3]])
    assert result["undo"] == "available"
    backup = o.excel.Workbooks(2)
    assert backup.Windows(1).Visible is False and backup.Names(undo.MARKER).Visible is False and backup.Saved
    assert backup.Sheets.Count == 2
    o.ws.formula_writes.clear()
    revert(o)
    assert o.ws.Range("A1:C1").Value == (("007", "=Sheet2!A1", "constant"),)
    assert o.ws.formula_writes == [("B1", "=Sheet2!A1")]
    assert o.ws.Columns(2).ColumnWidth == 24 and o.ws.Rows(1).RowHeight == 33
    assert backup.Sheets.Count == 1 and not stack(o)
    assert o.excel.EnableEvents is True


def test_expansion_scalar_fill_and_two_undo_steps(office):
    o = office
    write(o, [[1, 2], [3, 4]], "B2")
    assert stack(o)[-1].areas == [{"sheet": "Data", "address": "B2:C3"}]
    write(o, "007", "B2:C3")
    assert stack(o)[-1].areas[0]["address"] == "B2:C3"
    assert len(revert(o, steps=2)["undone"]) == 2
    assert o.ws.Range("B2:C3").Value == ((None, None), (None, None))


def test_formula_array_expands(office):
    o = office
    o.call("excel_set_formula", workbook=o.wb.Name, sheet="Data", cells="A1", formula=[["=1", "=2"]])
    assert stack(o)[-1].areas[0]["address"] == "A1:B1"
    revert(o)
    assert o.ws.Range("A1:B1").Value == ((None, None),)


def test_fingerprint_refuses_later_edits_and_force(office):
    o = office
    write(o, 2)
    o.ws.Range("A1").Value = 99
    with pytest.raises(ToolError, match="later edits"):
        revert(o)
    assert len(stack(o)) == 1 and o.ws.Range("A1").Value == 99
    revert(o, force=True)
    assert o.ws.Range("A1").Value is None


def test_size_barrier_stops_without_skipping(office, monkeypatch):
    o = office
    write(o, 1)
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo_max_cells=2))
    result = write(o, [[1, 2, 3]])
    assert result["undo"].startswith("not available:")
    assert len(stack(o)) == 2 and not stack(o)[-1].undoable
    with pytest.raises(ToolError, match="excel_write_range.*MAX_CELLS"):
        revert(o, force=True, steps=2)
    assert len(stack(o)) == 2


def test_depth_trims_backup_data_and_reuses_one_workbook(office, monkeypatch):
    o = office
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo_depth=2))
    for i in range(4):
        write(o, i)
    assert len(stack(o)) == 2
    assert o.excel.Workbooks.Count == 2 and o.excel.Workbooks(2).Sheets.Count == 3
    revert(o, steps=2)
    assert o.ws.Range("A1").Value == 1


def test_failed_tool_drops_snapshot_but_does_not_restore(office, monkeypatch):
    from tests.history_fakes import Range

    o = office
    write(o, 1)
    old = Range._assign

    def fail(self, value, formula=False):
        old(self, value, formula)
        raise ValueError("failure after write")

    monkeypatch.setattr(Range, "_assign", fail)
    with pytest.raises(ToolError, match="failure after write"):
        write(o, 8)
    assert len(stack(o)) == 1
    assert o.ws.Range("A1").Value == 8
    assert o.excel.Workbooks(2).Sheets.Count == 2


def test_backup_disappearance_is_a_permanent_barrier(office):
    o = office
    write(o, 1)
    o.excel.Workbooks(2).Close(False)
    history = revert(o, action="history")["history"]
    assert not history[0]["undoable"] and "closed" in history[0]["reason"]
    with pytest.raises(ToolError, match="closed"):
        revert(o, force=True)


def test_scratch_hidden_from_list_pick_and_all_workbooks(office):
    o = office
    write(o, 1)
    scratch = o.excel.Workbooks(2)
    assert [b["name"] for b in o.call("excel_list_workbooks")["workbooks"]] == [o.wb.Name]
    assert xl.all_workbooks() == [(o.excel, o.wb)]
    with pytest.raises(ToolError, match="not found"):
        com.run_com(lambda: xl.pick_workbook(scratch.Name), kind="read")
    scratch.Activate()
    assert com.run_com(lambda: xl.pick_workbook()[1].Name, kind="read") == o.wb.Name


def test_guard_prevents_any_snapshot_on_autosave(office):
    o = office
    o.wb.AutoSaveOn = True
    with pytest.raises(ToolError, match="AutoSave"):
        write(o, 1)
    assert o.excel.Workbooks.Count == 1 and not stack(o)


def test_undo_off_does_not_create_scratch(office, monkeypatch):
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo=False))
    result = write(office, 1)
    assert "undo" not in result and office.excel.Workbooks.Count == 1


def test_word_custom_records_undo_and_top_fingerprint(office):
    o = office
    assert o.word_write("first")["undo"] == "available"
    o.word_write("second")
    assert o.word.starts == ["Office Live: word_insert_text"] * 2 and o.word.ends == 2
    result = o.call("office_undo", file=o.doc.Name, steps=2)
    assert len(result["undone"]) == 2 and o.doc.Content.Text == "original" and o.doc.undo_calls == [1, 1]


def test_word_fingerprint_force_and_boolean_failure(office):
    o = office
    o.word_write("agent")
    o.doc.Content.Text = "user"
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.doc.Name)
    assert not o.doc.undo_calls
    o.doc.undo_ok = False
    with pytest.raises(ToolError, match=r"Undo\(1\)"):
        o.call("office_undo", file=o.doc.Name, force=True)
    assert o.doc.undo_calls == [1]


def test_word_force_and_failed_tool_cleanup(office):
    o = office
    o.word_write("agent")
    o.doc.Content.Text = "user"
    o.call("office_undo", file=o.doc.Name, force=True)
    assert o.doc.Content.Text == "original"
    with pytest.raises(ToolError, match="simulated failure"):
        o.word_write("partial", fail=True)
    assert o.word.ends == 2 and o.doc.Content.Text == "partial"
    assert o.call("office_undo", file=o.doc.Name, action="history")["history"] == []


def test_word_nested_record_is_not_closed_by_server(office):
    o = office
    o.word.UndoRecord.IsRecordingCustomRecord = True
    assert "not available" in o.word_write("agent")["undo"]
    assert o.word.starts == [] and o.word.ends == 0


def test_add_sheet_inverse_detects_created_sheet(office):
    o = office
    assert o.call("excel_add_worksheet", workbook=o.wb.Name, name="New")["undo"] == "available"
    revert(o)
    assert xl.sheet_names(o.wb) == ["Data"]


@pytest.mark.parametrize("axis,start,cell", [("rows", 2, "A3"), ("columns", 2, "C1")])
def test_insert_lines_inverse(office, axis, start, cell):
    o = office
    o.ws.Range(cell).Value = "original"
    o.call("excel_insert_rows_columns", workbook=o.wb.Name, sheet="Data", axis=axis, start=start)
    revert(o)
    assert o.ws.Range(cell).Value == "original"


def test_delete_rows_snapshot_insert_and_copy_back(office):
    o = office
    o.ws.Range("A1:A3").Value = ((1,), ("007",), (3,))
    result = o.call("excel_delete_rows_columns", workbook=o.wb.Name, sheet="Data", axis="rows", start=2)
    assert result["undo"] == "available"
    assert o.ws.Range("A2").Value == 3
    revert(o)
    assert o.ws.Range("A1:A3").Value == ((1,), ("007",), (3,))


def test_delete_column_is_full_column_size_barrier(office):
    o = office
    result = o.call("excel_delete_rows_columns", workbook=o.wb.Name, sheet="Data", axis="columns", start=2)
    assert "not available" in result["undo"] and "1048576" in result["undo"]


def test_delete_sheet_restores_index_name_and_warns_about_ref(office):
    o = office
    other = o.wb.Sheets.Add(None, o.ws)
    other.Name = "Other"
    o.ws.Range("A1").Value = "007"
    result = o.call("excel_delete_sheet", workbook=o.wb.Name, sheet="Data", confirm=True)
    assert "#REF!" in result["undo_warning"]
    result = revert(o)
    assert xl.sheet_names(o.wb) == ["Data", "Other"]
    assert o.wb.Worksheets("Data").Range("A1").Value == "007" and "#REF!" in result["warning"]


@pytest.mark.parametrize("action,kwargs", [
    ("rename", {"new_name": "Renamed"}), ("move", {"position": "end"}), ("tab_color", {"color": "red"}),
    ("hide", {}), ("very_hide", {}), ("unhide", {}), ("copy", {"new_name": "Copy"}),
])
def test_manage_sheet_inverse(office, action, kwargs):
    o = office
    other = o.wb.Sheets.Add(None, o.ws)
    other.Name = "Other"
    if action == "unhide":
        o.ws.Visible = 0
    before = o.ws.Visible
    result = o.call("excel_manage_sheet", workbook=o.wb.Name, sheet="Data", action=action, **kwargs)
    assert result["undo"] == "available"
    revert(o)
    assert xl.sheet_names(o.wb) == ["Data", "Other"] and o.ws.Visible == before
    assert o.ws.Tab.ColorIndex == -4142


@pytest.mark.parametrize("action", ["add", "delete"])
def test_manage_names_inverse_preserves_old_value_and_visibility(office, action):
    o = office
    o.wb.Names.Add("Rate", "=0.1", False)
    o.call("excel_manage_names", workbook=o.wb.Name, action=action, name="Rate", formula="=0.2")
    revert(o)
    assert o.wb.Names("Rate").RefersTo == "=0.1" and o.wb.Names("Rate").Visible is False


def test_add_new_name_inverse_deletes_it(office):
    o = office
    o.call("excel_manage_names", workbook=o.wb.Name, action="add", name="Rate", formula="=0.2")
    revert(o)
    with pytest.raises(KeyError):
        o.wb.Names("Rate")


def test_dimensions_and_hidden_inverse(office):
    o = office
    o.call("excel_set_dimensions", workbook=o.wb.Name, sheet="Data", cells="A1:B3", column_width=20, row_height=30)
    assert o.excel.Workbooks.Count == 1
    revert(o)
    assert o.ws.Columns(1).ColumnWidth == 8.43 and o.ws.Rows(3).RowHeight == 15
    entry = undo.Entry("excel_hide_rows_columns", "Data!1:3", "workbook")
    undo._hidden(o.excel, o.wb, entry, {"sheet": "Data", "axis": "rows", "lines": "1-3"})
    o.ws.Rows(2).Hidden = True
    com.run_com(lambda: undo.restore_excel(o.excel, o.wb, entry))
    assert o.ws.Rows(2).Hidden is False


def test_sheet_view_and_calculation_inverse(office):
    o = office
    o.call("excel_sheet_view", workbook=o.wb.Name, sheet="Data", freeze_at="B2", zoom=150, gridlines=False, headings=False)
    revert(o)
    assert o.ws.view.Zoom == 100 and not o.ws.view.FreezePanes and o.ws.view.DisplayGridlines
    assert "undo" not in o.call("excel_calculate", workbook=o.wb.Name)
    o.call("excel_calculate", workbook=o.wb.Name, mode="manual")
    revert(o)
    assert o.excel.Calculation == -4105


@pytest.mark.parametrize("tool,result_key,op", [("excel_create_chart", "chart", "shape"), ("excel_insert_image", "shape", "shape"), ("excel_manage_slicers", "cache", "slicer"), ("excel_manage_tables", "table", "table")])
def test_created_object_inverse_and_table_format_snapshot(office, tool, result_key, op):
    o = office
    entry = undo.Entry(tool, "Data!A1", "workbook")
    args = {"sheet": "Data", "cells": "A1", "action": "create" if op == "table" else "add"}
    o.excel.EnableEvents = False
    undo.RESOLVERS[tool](o.excel, o.wb, entry, args)
    calls = []
    item = NS(Name="Created", Delete=lambda: calls.append("delete"), Unlist=lambda: calls.append("unlist"))
    coll = o.wb.SlicerCaches if op == "slicer" else o.ws.ListObjects if op == "table" else o.ws.Shapes
    coll.items.append(item)
    o.ws.Range("A1").Value = "table header"
    undo._finish_excel(o.wb, entry, {result_key: item.Name})
    com.run_com(lambda: undo.restore_excel(o.excel, o.wb, entry))
    assert calls == ["unlist" if op == "table" else "delete"]
    if op == "table":
        assert o.ws.Range("A1").Value is None


def test_cross_workbook_copy_records_destination_and_transpose(office):
    o = office
    dest = o.excel.Workbooks.Add()
    o.ws.Range("A1:B1").Value = ((1, 2),)
    result = o.call("excel_copy_range", workbook=o.wb.Name, sheet="Data", source="A1:B1", dest_cell="D2", dest_sheet="Data", dest_workbook=dest.Name, what="values", transpose=True)
    assert result["undo"] == "available" and not stack(o)
    entry = undo.STACKS[undo.stack_key("workbook", o.excel, dest)][-1]
    assert entry.areas == [{"sheet": "Data", "address": "D2:D3"}]
    o.call("office_undo", file=dest.Name)
    assert dest.Worksheets("Data").Range("D2:D3").Value == ((None,), (None,))


def test_move_snapshots_both_source_and_destination(office):
    o = office
    o.ws.Range("A1").Value = "007"
    o.call("excel_copy_range", workbook=o.wb.Name, sheet="Data", source="A1", dest_cell="C1", move=True)
    assert len(stack(o)[-1].areas) == 2
    revert(o)
    assert o.ws.Range("A1:C1").Value == (("007", None, None),)


def test_autofill_bounding_box_and_replace_active_sheet_plan(office):
    o = office
    o.excel.EnableEvents = False
    entry = undo.Entry("excel_autofill", "", "workbook")
    undo._autofill(o.excel, o.wb, entry, {"sheet": "Data", "source": "B2:C3", "dest": "B2:C9"})
    assert entry.areas == [{"sheet": "Data", "address": "B2:C9"}]
    o.wb.Sheets.Add(None, o.ws)
    entry = undo.Entry("excel_replace", "", "workbook")
    undo._replace(o.excel, o.wb, entry, {"sheet": "", "cells": ""})
    assert entry.areas == [{"sheet": o.wb.ActiveSheet.Name, "address": "A1"}]


def test_unknown_excel_write_is_barrier_and_read_action_does_not_record(office):
    from office_live import registry

    o = office
    write(o, 1)

    def unknown(workbook):
        xl.pick_workbook(workbook)
        return {"ok": True}

    unknown.__name__ = "excel_future_tool"
    saved = []
    old = registry.mcp
    registry.mcp = NS(add_tool=lambda fn, **kw: saved.append(fn))
    try:
        registry.office_tool("excel_core", "write")(unknown)
    finally:
        registry.mcp = old
    assert "not available" in saved[0](workbook=o.wb.Name)["undo"]
    o.call("excel_manage_names", workbook=o.wb.Name, action="list")
    assert len(stack(o)) == 2
    with pytest.raises(ToolError, match="excel_future_tool"):
        revert(o)


def test_cross_workbook_move_undo_is_one_coordinated_step(office):
    o = office
    dest = o.excel.Workbooks.Add()
    o.ws.Range("A1").Value = "007"
    dest.Worksheets("Data").Range("C1").Value = "old destination"
    o.call("excel_copy_range", workbook=o.wb.Name, sheet="Data", source="A1", dest_cell="C1", dest_sheet="Data", dest_workbook=dest.Name, move=True)
    result = o.call("office_undo", file=dest.Name)
    assert len(result["undone"]) == 1 and o.ws.Range("A1").Value == "007"
    assert dest.Worksheets("Data").Range("C1").Value == "old destination"
    assert all(not entries for entries in undo.STACKS.values())


def test_cross_workbook_move_refuses_to_skip_newer_source_change(office):
    o = office
    dest = o.excel.Workbooks.Add()
    o.ws.Range("A1").Value = "007"
    o.call("excel_copy_range", workbook=o.wb.Name, sheet="Data", source="A1", dest_cell="C1", dest_sheet="Data", dest_workbook=dest.Name, move=True)
    write(o, 9)
    with pytest.raises(ToolError, match="newer changes"):
        o.call("office_undo", file=dest.Name, force=True)
    revert(o)
    o.call("office_undo", file=dest.Name)
    assert o.ws.Range("A1").Value == "007"


def test_cross_workbook_sheet_copy_undo_targets_destination(office):
    o = office
    dest = o.excel.Workbooks.Add()
    result = o.call("excel_manage_sheet", workbook=o.wb.Name, sheet="Data", action="copy", new_name="Copied", dest_workbook=dest.Name)
    assert result["undo"] == "available" and not stack(o)
    o.call("office_undo", file=dest.Name)
    assert xl.sheet_names(dest) == ["Data"] and xl.sheet_names(o.wb) == ["Data"]


def test_bridge_word_text_snapshot_uses_filtered_source_size(office):
    from tests.history_fakes import Collection

    o = office
    o.doc.Paragraphs = Collection([
        NS(OutlineLevel=1, Range=NS(Text="Heading\r"), Style=NS(NameLocal="Heading 1")),
        NS(OutlineLevel=10, Range=NS(Text="body\r"), Style=NS(NameLocal="Normal")),
        NS(OutlineLevel=1, Range=NS(Text="\r"), Style=NS(NameLocal="Heading 1")),
    ])
    result = o.call("bridge_word_text_to_excel", document=o.doc.Name, workbook=o.wb.Name, sheet="Data", top_left="B3", headings_only=True)
    assert result["undo"] == "available" and stack(o)[-1].areas == [{"sheet": "Data", "address": "B3:D4"}]
    assert not o.word.starts
    revert(o)
    assert o.ws.Range("B3:D4").Value == ((None, None, None), (None, None, None))


def test_bridge_word_table_snapshot_and_new_sheet_inverse(office, monkeypatch):
    from office_live import bridge, word_tables

    o = office
    grid = {"uniform": True, "values": [["007", "5"], ["text", "8"]]}
    for module in (bridge, word_tables):
        monkeypatch.setattr(module, "get_table", lambda *a: None)
        monkeypatch.setattr(module, "table_grid", lambda *a: grid)
    result = o.call("bridge_word_table_to_excel", document=o.doc.Name, table_index=1, workbook=o.wb.Name, sheet="Data", top_left="C2", decimal_separator=".")
    assert result["undo"] == "available" and stack(o)[-1].areas[0]["address"] == "C2:D3"
    revert(o)
    o.call("bridge_word_table_to_excel", document=o.doc.Name, table_index=1, workbook=o.wb.Name, sheet="Imported", create_sheet=True, decimal_separator=".")
    revert(o)
    assert xl.sheet_names(o.wb) == ["Data"] and not o.word.starts


def test_bridge_into_word_uses_native_record_and_no_excel_entry(office):
    from office_live import registry, wd_common

    o = office
    wrappers = []
    old = registry.mcp
    registry.mcp = NS(add_tool=lambda fn, **kw: wrappers.append(fn))
    try:
        @registry.office_tool("bridge", "write")
        def bridge_excel_range_to_word_table(workbook, document):
            xl.pick_workbook(workbook)
            _, doc = wd_common.pick_document(document)
            doc.Content.Text = "table from Excel"
            return {"ok": True}
    finally:
        registry.mcp = old
    assert wrappers[0](workbook=o.wb.Name, document=o.doc.Name)["undo"] == "available"
    assert not stack(o) and o.word.ends == 1
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.Content.Text == "original"


def test_hidden_where_stores_all_data_rows_before_matching(office):
    o = office
    o.ws.Range("A3:A6").Value = (("Header",), ("yes",), ("no",), ("yes",))
    entry = undo.Entry("excel_hide_rows_columns", "", "workbook")
    undo._hidden(o.excel, o.wb, entry, {"sheet": "Data", "cells": "A3:A6", "where": {"column": "A", "value": "yes"}})
    assert entry.ops[0]["values"] == [(3, False), (4, False), (5, False), (6, False)]


def test_readonly_undo_history_does_not_trigger_autosave_or_events(office, monkeypatch):
    o = office
    write(o, 1)
    o.wb.AutoSaveOn = True
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, mode="readonly"))
    assert revert(o, action="history")["history"][0]["undoable"]
    assert o.excel.EnableEvents and len(stack(o)) == 1


def test_doctor_and_live_cleanup_never_expose_backup(office, monkeypatch, capsys):
    from office_live import cli
    from tests.live.conftest import close_own_workbooks

    o = office
    write(o, 1)
    backup = o.excel.Workbooks(2)
    monkeypatch.setattr(cli, "_office_installed", lambda name: True)
    cli.doctor()
    output = capsys.readouterr().out
    assert "open files: ['Book1']" in output and backup.Name not in output
    calls = []

    def call(tool, **kw):
        calls.append(tool)
        return o.call(tool, **kw)

    close_own_workbooks(NS(call=call), backup.Name, "not-a-user-file")
    assert calls == ["excel_list_workbooks"] and backup in o.excel.Workbooks.items


def test_failed_undo_keeps_a_barrier_and_never_skips_it(office, monkeypatch):
    o = office
    write(o, 1)
    monkeypatch.setattr(undo, "restore_excel", lambda *a: (_ for _ in ()).throw(ValueError("restore interrupted")))
    with pytest.raises(ToolError, match="partially applied"):
        revert(o)
    assert len(stack(o)) == 1 and not stack(o)[-1].undoable


def test_snapshot_preserves_other_active_workbook(office):
    o = office
    other = o.excel.Workbooks.Add()
    write(o, 1)
    assert o.excel.ActiveWorkbook is other


def test_noop_word_tool_does_not_claim_native_undo(office):
    o = office
    result = o.word_write(o.doc.Content.Text)
    assert result["undo"] == "not available: no document change detected"
    assert o.call("office_undo", file=o.doc.Name, action="history")["history"] == []


def test_snapshot_service_respects_events_opt_in_for_the_tool_body(office, monkeypatch):
    from tests.history_fakes import Range

    o = office
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, enable_events=True))
    old = Range._assign

    def checked(self, value, formula=False):
        if self.Worksheet.Parent is o.wb:
            assert o.excel.EnableEvents
        return old(self, value, formula)

    monkeypatch.setattr(Range, "_assign", checked)
    assert write(o, 1)["undo"] == "available"


def test_close_drops_history_and_backups_without_adding_a_step(office):
    from office_live import registry

    o = office
    write(o, 1)
    registered = []
    old = registry.mcp
    registry.mcp = NS(add_tool=lambda fn, **kw: registered.append(fn))
    try:
        @registry.office_tool("excel_core", "destructive")
        def excel_close_workbook(workbook):
            _, wb = xl.pick_workbook(workbook)
            wb.Close(False)
            return {"ok": True}
    finally:
        registry.mcp = old
    result = registered[0](workbook=o.wb.Name)
    assert "undo" not in result and not undo.STACKS
    assert o.excel.Workbooks(1).Sheets.Count == 1


def test_save_as_moves_history_to_new_identity_without_new_step(office):
    from office_live import registry

    o = office
    write(o, 1)
    registered = []
    old = registry.mcp
    registry.mcp = NS(add_tool=lambda fn, **kw: registered.append(fn))
    try:
        @registry.office_tool("excel_core", "save")
        def excel_save_as(workbook):
            _, wb = xl.pick_workbook(workbook)
            wb.Name, wb.FullName, wb.Path = "Saved.xlsx", "C:/Docs/Saved.xlsx", "C:/Docs"
            return {"ok": True}
    finally:
        registry.mcp = old
    assert "undo" not in registered[0](workbook=o.wb.Name)
    assert len(stack(o)) == 1
    revert(o)
    assert o.ws.Range("A1").Value is None


def test_partial_word_failure_cannot_be_skipped_after_a_new_success(office):
    o = office
    o.word_write("first")
    with pytest.raises(ToolError):
        o.word_write("failed partial change", fail=True)
    assert len(o.call("office_undo", file=o.doc.Name, action="history")["history"]) == 1
    o.word_write("last")
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.Content.Text == "failed partial change"
    with pytest.raises(ToolError, match="failed word_insert_text"):
        o.call("office_undo", file=o.doc.Name, force=True)


def test_failed_validation_leaves_previous_undo_available(office):
    o = office
    write(o, 1)
    with pytest.raises(ToolError):
        o.call("excel_clear_range", workbook=o.wb.Name, sheet="Data", cells="A1", what="invalid")
    assert len(stack(o)) == 1 and stack(o)[-1].undoable
    revert(o)
    assert o.ws.Range("A1").Value is None


def test_interleaved_user_edit_still_blocks_older_excel_undo(office):
    o = office
    write(o, 1)
    o.ws.Range("A1").Value = "user change"
    write(o, 2, "B1")
    revert(o)
    with pytest.raises(ToolError, match="later edits"):
        revert(o)
    assert o.ws.Range("A1").Value == "user change"


def test_interleaved_user_edit_still_blocks_older_word_undo(office):
    o = office
    o.word_write("first")
    o.doc.save_state()
    o.doc.Content.Text = "user change"
    o.word_write("last")
    o.call("office_undo", file=o.doc.Name)
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.doc.Name)
    assert o.doc.Content.Text == "user change" and o.doc.undo_calls == [1]


def test_word_fingerprint_ignores_spellcheck_marks_and_edit_session_ids():
    """Фоновая проверка орфографии и rsid меняются без пользователя — отпечаток от этого меняться не должен,
    а настоящая правка форматирования (w:b) — должна."""
    from types import SimpleNamespace

    from office_live import undo

    def doc(body, settings="<w:proofState w:spelling='clean'/>"):
        xml = (f'<pkg:package><pkg:part pkg:name="/word/settings.xml">{settings}</pkg:part>'
               f'<pkg:part pkg:name="/word/document.xml"><w:body>{body}</w:body></pkg:part></pkg:package>')
        one = SimpleNamespace(Count=1)
        from tests.history_fakes import Word

        target = Word().ActiveDocument
        target.Content = SimpleNamespace(Text="Hello", WordOpenXML=xml)
        target.Paragraphs = target.Tables = target.InlineShapes = one
        return target

    base = undo.word_fingerprint(doc('<w:p w:rsidR="00A1"><w:r><w:t>Hello</w:t></w:r></w:p>'))
    noisy = doc('<w:p w:rsidR="00B2" w:rsidRDefault="00C3"><w:proofErr w:type="spellStart"/><w:r><w:t>Hello</w:t></w:r></w:p>', settings="<w:proofState/>")
    assert undo.word_fingerprint(noisy) == base
    bold = doc('<w:p w:rsidR="00A1"><w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r></w:p>')
    assert undo.word_fingerprint(bold) != base
