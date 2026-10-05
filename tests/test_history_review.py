"""Независимые регрессии истории: только строгие подделки COM, без приложений Office."""

from types import SimpleNamespace as NS
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from dataclasses import replace

import pytest

from office_live import com, config, journal, undo
from office_live.errors import ToolError
from tests.history_fakes import Collection, Section, fake_office  # noqa: F401 — фикстура office


@pytest.mark.parametrize("sheet", ["", "data", "DATA"])
def test_bridge_create_sheet_respects_existing_case_insensitive_target(office, sheet):
    o = office
    o.doc.Paragraphs = Collection([NS(OutlineLevel=1, Range=NS(Text="Heading\r"), Style=NS(NameLocal="Heading 1"))])
    o.ws.Range("A1:C2").Value = (("old", 2, 3), (4, 5, 6))
    result = o.call("bridge_word_text_to_excel", document=o.doc.Name, workbook=o.wb.Name,
                    sheet=sheet, create_sheet=True, overwrite=True)
    assert result["undo"] == "available"
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1:C2").Value == (("old", 2, 3), (4, 5, 6))


def test_copy_single_cell_colon_destination_expands_snapshot(office):
    o = office
    o.ws.Range("A1:B1").Value = ((1, 2),)
    o.ws.Range("D1:E1").Value = (("keep", "also keep"),)
    o.call("excel_copy_range", workbook=o.wb.Name, sheet="Data", source="A1:B1", dest_cell="D1:D1")
    assert o.ws.Range("D1:E1").Value == ((1, 2),)
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("D1:E1").Value == (("keep", "also keep"),)


@pytest.mark.parametrize("action", ["add", "delete"])
def test_structural_undo_survives_log_sheet_enable(office, action):
    o = office
    if action == "add":
        o.call("excel_add_worksheet", workbook=o.wb.Name, name="Other")
    else:
        other = o.wb.Sheets.Add(None, o.ws)
        other.Name = "Other"
        o.call("excel_delete_sheet", workbook=o.wb.Name, sheet="Other", confirm=True)
    o.call("office_journal", workbook=o.wb.Name, action="enable_sheet")
    o.call("office_undo", file=o.wb.Name)
    assert o.wb.Worksheets(journal.LOG_SHEET).Range("B3").Value == "office_undo"
    assert ("Other" in [s.Name for s in o.wb.Sheets]) == (action == "delete")


def test_replace_does_not_snapshot_other_sheets_or_log(office):
    o = office
    o.call("office_journal", workbook=o.wb.Name, action="enable_sheet")
    other = o.wb.Sheets.Add(None, o.ws)
    other.Name = "Other"
    other.Range("A1").Value = "untouched"
    o.ws.Activate()
    entry = undo.Entry("excel_replace", "", "workbook")
    com.run_com(lambda: undo._replace(o.excel, o.wb, entry, {"sheet": "", "cells": ""}))
    assert entry.areas == [{"sheet": "Data", "address": "A1"}]
    other.Range("A1").Value = "user edit"
    com.run_com(lambda: undo.restore_excel(o.excel, o.wb, entry))
    assert other.Range("A1").Value == "user edit"


def test_formula_repair_preserves_dynamic_array_dialect(office, monkeypatch):
    from tests.history_fakes import Range

    o = office
    formula = "=SQRT(Sheet2!A1:A4)"
    o.ws.Range("A1").Formula2 = formula
    original = Range.Formula

    def legacy_write(rng, value):
        raise AssertionError("Formula setter suppresses dynamic-array spilling")

    monkeypatch.setattr(Range, "Formula", property(original.fget, legacy_write))
    o.call("excel_write_range", workbook=o.wb.Name, sheet="Data", cells="A1", values=1)
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1").Formula2 == formula


def test_formula_snapshot_falls_back_on_legacy_excel(office, monkeypatch):
    from tests.history_fakes import Range

    def unavailable(rng):
        raise AttributeError("Formula2")

    monkeypatch.setattr(Range, "Formula2", property(unavailable))
    o = office
    o.ws.Range("A1").Formula = "=Sheet2!A1"
    o.call("excel_write_range", workbook=o.wb.Name, sheet="Data", cells="A1", values=1)
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1").Formula == "=Sheet2!A1"


def test_journal_failure_isolated_per_target(office, monkeypatch, capsys):
    blocked = journal.journal_path("Blocked.xlsx")
    original = Path.open

    def checked(path, *args, **kwargs):
        if path == blocked:
            raise PermissionError("locked journal")
        return original(path, *args, **kwargs)

    monkeypatch.setattr(Path, "open", checked)
    monkeypatch.setattr(journal, "_warned", False)
    for _ in range(2):
        journal.append(["workbook:Blocked.xlsx", "workbook:Writable.xlsx"], "excel_copy_range", {})
    assert journal.journal_path("Writable.xlsx").read_text(encoding="utf-8").count("excel_copy_range") == 2
    assert capsys.readouterr().err.count("journal is not writable") == 1


def test_service_log_range_is_never_snapshotted(office):
    o = office
    o.call("office_journal", workbook=o.wb.Name, action="enable_sheet")
    result = o.call("excel_write_range", workbook=o.wb.Name, sheet=journal.LOG_SHEET, cells="G1", values="note")
    assert "service log sheet" in result["undo"]
    assert o.excel.Workbooks.Count == 1


def test_service_log_deletion_is_not_a_restorable_sheet_copy(office):
    o = office
    o.call("office_journal", workbook=o.wb.Name, action="enable_sheet")
    result = o.call("excel_delete_sheet", workbook=o.wb.Name, sheet=journal.LOG_SHEET, confirm=True)
    assert "service log sheet" in result["undo"]
    assert o.excel.Workbooks.Count == 1


@pytest.mark.parametrize("events_allowed", [False, True])
def test_snapshot_copy_error_restores_app_state_and_keeps_backup_saved(office, monkeypatch, events_allowed):
    from tests.history_fakes import Range

    o = office
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, enable_events=events_allowed))

    def fail_copy(self, destination):
        assert not o.excel.EnableEvents
        raise ValueError("copy rejected")

    monkeypatch.setattr(Range, "Copy", fail_copy)
    result = o.call("excel_write_range", workbook=o.wb.Name, sheet="Data", cells="A1", values=1)
    assert "not available" in result["undo"] and o.ws.Range("A1").Value == 1
    assert o.excel.EnableEvents and o.excel.DisplayAlerts
    backup = o.excel.Workbooks(2)
    assert backup.Saved and not backup.Windows(1).Visible and backup.Sheets.Count == 1


def test_delete_inverse_error_restores_alerts_and_events(office, monkeypatch):
    o = office
    o.call("excel_add_worksheet", workbook=o.wb.Name, name="Added")

    def fail_delete():
        assert not o.excel.DisplayAlerts and not o.excel.EnableEvents
        raise ValueError("protected workbook")

    monkeypatch.setattr(o.wb.Sheets("Added"), "Delete", fail_delete)
    with pytest.raises(ToolError, match="partially applied"):
        o.call("office_undo", file=o.wb.Name)
    assert o.excel.DisplayAlerts and o.excel.EnableEvents


def test_journal_concurrent_threads_keep_single_header_and_all_rows(office):
    with ThreadPoolExecutor(max_workers=8) as pool:
        list(pool.map(lambda i: journal.append(["workbook:Concurrent.xlsx"], "write", {"index": i}), range(32)))
    lines = journal.journal_path("Concurrent.xlsx").read_text(encoding="utf-8").splitlines()
    assert len(lines) == 33 and sum(line.startswith("# ") for line in lines) == 1


@pytest.mark.parametrize("section_index", [1, 2])
@pytest.mark.parametrize("collection", ["Headers", "Footers"])
@pytest.mark.parametrize("part_index", [1, 2, 3])
def test_word_header_footer_edit_blocks_undo_and_force_reaches_original(office, section_index, collection, part_index):
    o = office
    o.doc.Sections.items.append(Section())
    part = getattr(o.doc.Sections(section_index), collection)(part_index)
    part.Exists = True
    part.Range.Text = "original header/footer"
    before = undo.word_fingerprint(o.doc)
    o.word_write("agent")
    o.doc.save_state()
    part.Range.Text = "user header/footer"
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.doc.Name)
    assert o.doc.undo_calls == []
    result = o.call("office_undo", file=o.doc.Name, force=True)
    assert result["native_steps"] == result["undone"][0]["native_steps"] == 2
    assert undo.word_fingerprint(o.doc) == before


@pytest.mark.parametrize("collection", ["Comments", "Footnotes", "Endnotes", "Shapes"])
def test_word_story_count_edit_blocks_undo(office, collection):
    o = office
    o.word_write("agent")
    getattr(o.doc, collection).items.append(NS(Range=NS(Text="user story")))
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.doc.Name)
    assert not o.doc.undo_calls


def test_word_existing_comment_text_edit_and_header_existence_change_fingerprint(office):
    doc = office.doc
    doc.Comments.items.extend([NS(Range=NS(Text="first")), NS(Range=NS(Text="second"))])
    before = undo.word_fingerprint(doc)
    doc.Comments(2).Range.Text = "changed"
    assert undo.word_fingerprint(doc) != before
    before = undo.word_fingerprint(doc)
    doc.Sections(1).Headers(2).Exists = True
    assert undo.word_fingerprint(doc) != before  # отсутствующий колонтитул отличается от существующего пустого


def test_word_fingerprint_does_not_walk_paragraphs(office):
    office.doc.Paragraphs = NS(Count=100000)  # нет Item/__call__/итерации
    assert undo.word_fingerprint(office.doc)


@pytest.mark.parametrize("user_steps", [0, 1, 3, 19])
def test_word_force_native_steps_include_agent_and_later_user_edits(office, user_steps):
    o = office
    o.word_write("agent")
    for i in range(user_steps):
        o.doc.save_state()
        o.doc.Content.Text = f"user {i}"
    result = o.call("office_undo", file=o.doc.Name, force=True)
    assert result["native_steps"] == user_steps + 1
    assert o.doc.undo_calls == [1] * (user_steps + 1) and o.doc.Content.Text == "original"
    assert o.call("office_undo", file=o.doc.Name, action="history")["history"] == []


def test_word_force_limit_keeps_entry_as_barrier(office):
    o = office
    o.word_write("agent")
    for i in range(20):
        o.doc.save_state()
        o.doc.Content.Text = f"user {i}"
    with pytest.raises(ToolError, match="Could not reach the state before word_insert_text; 20 native steps were undone, including the user's edits"):
        o.call("office_undo", file=o.doc.Name, force=True)
    assert o.doc.Content.Text == "agent" and o.doc.undo_calls == [1] * 20
    history = o.call("office_undo", file=o.doc.Name, action="history")["history"]
    assert len(history) == 1 and not history[0]["undoable"]
    with pytest.raises(ToolError, match="not available"):
        o.call("office_undo", file=o.doc.Name, force=True)
    assert o.doc.undo_calls == [1] * 20


def test_word_force_exhausted_native_history_keeps_entry(office):
    o = office
    o.word_write("agent")
    o.doc.states.clear()  # Word потерял историю агента, позднее появилось ручное действие
    o.doc.save_state()
    o.doc.Content.Text = "user"
    with pytest.raises(ToolError, match=r"1 native steps were undone.*Undo\(1\)"):
        o.call("office_undo", file=o.doc.Name, force=True)
    assert o.doc.Content.Text == "agent" and o.doc.undo_calls == [1, 1]
    assert not o.call("office_undo", file=o.doc.Name, action="history")["history"][0]["undoable"]


def test_word_force_partial_multientry_result_reports_total_and_stop(office):
    o = office
    o.word_write("first")
    key = undo.stack_key("document", o.word, o.doc)
    undo.STACKS[key][-1].before_fingerprint = "unreachable state"
    o.word_write("second")
    result = o.call("office_undo", file=o.doc.Name, force=True, steps=2)
    assert result["native_steps"] == 2 and len(result["undone"]) == 1
    assert result["undone"][0]["native_steps"] == 1
    assert "1 native steps were undone" in result["stopped"]
    assert len(undo.STACKS[key]) == 1 and not undo.STACKS[key][0].undoable


FORMAT_CHANGES = [("NumberFormat", "0.00"), ("Font.Name", "Arial"), ("Font.Size", 14), ("Font.Bold", True),
                  ("Font.Italic", True), ("Font.Underline", 2), ("Font.Color", 255), ("Interior.Color", 255),
                  ("Interior.Pattern", 1), ("HorizontalAlignment", -4108), ("VerticalAlignment", -4108),
                  ("WrapText", True), ("MergeCells", True)]


def set_format(rng, prop, value):
    if "." in prop:
        parent, prop = prop.split(".")
        rng = getattr(rng, parent)
    setattr(rng, prop, value)


@pytest.mark.parametrize("prop,value", FORMAT_CHANGES)
def test_excel_format_edit_blocks_snapshot_undo_and_force_restores_it(office, prop, value):
    o = office
    original = undo._format_state(o.ws.Range("A1:B2"))
    o.call("excel_write_range", workbook=o.wb.Name, sheet="Data", cells="A1:B2", values=1)
    set_format(o.ws.Range("A1"), prop, value)
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1").Value == 1
    o.call("office_undo", file=o.wb.Name, force=True)
    assert undo._format_state(o.ws.Range("A1:B2")) == original


def test_excel_format_signature_reads_each_whole_range_property_once(office, monkeypatch):
    from collections import Counter
    from tests.history_fakes import Range

    reads = []
    for member in ("Font", "Interior"):
        getter = getattr(Range, member).fget

        def counted(rng, name=member, get=getter):
            reads.append((rng.Address, name))
            return get(rng)

        monkeypatch.setattr(Range, member, property(counted))
    o = office
    entry = undo.Entry("excel_write_range", "", "workbook", areas=[{"sheet": "Data", "address": "A1:B2"}])
    o.ws.format_reads.clear()
    undo.excel_fingerprint(o.excel, o.wb, entry)
    assert Counter(o.ws.format_reads) == Counter({("A1:B2", prop): 1 for prop, _ in FORMAT_CHANGES})
    assert reads == [("A1:B2", "Font"), ("A1:B2", "Interior")]


def test_excel_already_mixed_format_signature_has_documented_limit(office):
    o = office
    o.ws.Range("A1").Font.Bold = True
    entry = undo.Entry("excel_write_range", "", "workbook", areas=[{"sheet": "Data", "address": "A1:C1"}])
    assert o.ws.Range("A1:C1").Font.Bold is None
    before = undo.excel_fingerprint(o.excel, o.wb, entry)
    o.ws.Range("B1").Font.Bold = True
    assert o.ws.Range("A1:C1").Font.Bold is None
    assert undo.excel_fingerprint(o.excel, o.wb, entry) == before


@pytest.mark.parametrize("axis,address,kept,skipped", [("Columns", "A2:F2", "heights", "widths"), ("Rows", "C1:C8", "widths", "heights")])
def test_full_line_dimensions_skip_orthogonal_axis_using_sheet_counts(office, monkeypatch, axis, address, kept, skipped):
    o = office
    # Нестандартный размер листа ловит жёстко заданные 16384/1048576. У пропускаемой оси нет Item.
    monkeypatch.setattr(o.ws, axis, NS(Count=6 if axis == "Columns" else 8))
    dimensions = undo._dimensions(o.ws.Range(address))
    assert dimensions[skipped] == [] and len(dimensions[kept]) == 1
    undo._restore_dimensions(o.ws, dimensions)  # пропускаемая ось также не восстанавливается


def test_explicit_dimensions_still_capture_full_axis(office):
    o = office
    o.ws.Columns.Count, o.ws.Rows.Count = 3, 4
    o.call("excel_set_dimensions", workbook=o.wb.Name, sheet="Data", cells="A1:C4", column_width=20, row_height=30)
    assert o.ws.Columns(3).ColumnWidth == 20 and o.ws.Rows(4).RowHeight == 30
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Columns(3).ColumnWidth == 8.43 and o.ws.Rows(4).RowHeight == 15


@pytest.mark.parametrize("axis,address", [("rows", "2:2"), ("columns", "B:B")])
def test_deleted_lines_warn_about_unrestored_dependencies(office, monkeypatch, axis, address):
    from tests.history_fakes import Range
    from office_live import excel_core

    o = office
    # Маленький лист позволяет проверить удаление столбца, не обходя лимит ячеек и не выделяя миллион ячеек.
    o.ws.Rows.Count, o.ws.Columns.Count = 8, 6
    monkeypatch.setattr(excel_core, "_rows_cols_range", lambda ws, axis, start, count: Range(ws, "A2:F2" if axis == "rows" else "B1:B8"))
    o.ws.Range("B2").Value = "007"
    other = o.wb.Sheets.Add(None, o.ws)
    other.Range("A1").Formula = "=Data!B2"
    result = o.call("excel_delete_rows_columns", workbook=o.wb.Name, sheet="Data", axis=axis, start=2)
    assert result["undo"] == "available" and f"deleted {axis}" in result["undo_warning"]
    # Моделируем реакцию Excel на удаление ссылки; restore не затрагивает другой лист.
    other.Range("A1").Formula = "=#REF!"
    restored = o.call("office_undo", file=o.wb.Name)
    assert restored["warning"] == result["undo_warning"] and "#REF!" in restored["warning"]
    assert o.ws.Range("B2").Value == "007" and other.Range("A1").Formula == "=#REF!"


@pytest.mark.parametrize("cross_workbook", [False, True])
def test_move_warns_about_dependencies_outside_snapshots(office, cross_workbook):
    o = office
    dest = o.excel.Workbooks.Add() if cross_workbook else o.wb
    o.ws.Range("A1").Value = "007"
    o.ws.Range("F1").Formula = "=A1"
    result = o.call("excel_copy_range", workbook=o.wb.Name, sheet="Data", source="A1", dest_cell="C1", move=True,
                    dest_workbook=dest.Name if cross_workbook else "")
    assert result["undo"] == "available" and "keep pointing to the destination" in result["undo_warning"]
    for entries in undo.STACKS.values():
        assert entries[-1].warning == result["undo_warning"]
    # Cut перенаправляет зависимую формулу; обратное копирование снимков не выполняет обратный Cut.
    redirected = f"=[{dest.Name}]Data!C1" if cross_workbook else "=C1"
    o.ws.Range("F1").Formula = redirected
    restored = o.call("office_undo", file=dest.Name)
    assert restored["warning"] == result["undo_warning"]
    assert o.ws.Range("A1").Value == "007" and o.ws.Range("F1").Formula == redirected


def test_copy_without_move_has_no_dependency_warning(office):
    o = office
    result = o.call("excel_copy_range", workbook=o.wb.Name, sheet="Data", source="A1", dest_cell="C1")
    assert "undo_warning" not in result
    assert "warning" not in o.call("office_undo", file=o.wb.Name)
