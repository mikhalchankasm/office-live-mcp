"""Журнал и доступ к history: файлы только в tmp_path, листы и приложения — подделки."""

import hashlib
from dataclasses import replace
from types import SimpleNamespace as NS

import pytest

from office_live import com, config, journal, registry, undo
from office_live.errors import ToolError
from tests.history_fakes import fake_office  # noqa: F401 — общая фикстура office


def test_line_format_per_target_and_failure(office, monkeypatch):
    monkeypatch.setattr(journal.time, "strftime", lambda fmt: "2026-10-05 14:03:12")
    targets = [r"workbook:D:\Work\Book.xlsx", r"document:D:\Work\Draft.docx", r"workbook:D:\Work\Book.xlsx"]
    args = {"workbook": "Book.xlsx", "document": "Draft.docx", "sheet": "Data", "cells": "A1:C5", "bold": True, "fill_color": "#FFFF00"}
    journal.append(targets, "excel_format_range", args)
    path = journal.journal_path(r"D:\Work\Book.xlsx")
    text = path.read_text(encoding="utf-8")
    assert text == "# D:\\Work\\Book.xlsx\n- 2026-10-05 14:03:12 · excel_format_range · Data!A1:C5 · bold=True, fill_color='#FFFF00' · ok\n"
    assert len(list(config.SETTINGS.journal_dir.glob("*.md"))) == 2
    journal.append(targets[:1], "failed", {}, ValueError("x" * 230))
    assert path.read_text(encoding="utf-8").splitlines()[-1].endswith("ошибка: " + "x" * 200)


def test_hidden_content_and_audit_content_opt_in(office, monkeypatch):
    identity = "Book.xlsx"
    args = {"values": [["secret"]] * 3, "text": "private" * 30, "config": {"secret": "private"}}
    journal.append([f"workbook:{identity}"], "write", args)
    path = journal.journal_path(identity)
    text = path.read_text(encoding="utf-8")
    assert "private" not in text and "secret" not in text
    assert "[3 items]" in text and "<str len=210>" in text and "{1 keys}" in text
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, audit_content=True))
    journal.append([f"workbook:{identity}"], "write", args)
    assert "private" in path.read_text(encoding="utf-8") and "...(+50)" in path.read_text(encoding="utf-8")


def test_name_hash_sanitization_and_single_physical_line(office):
    identity = r"C:\folder\bad:name.xlsx"
    path = journal.journal_path(identity)
    digest = hashlib.sha1(identity.lower().encode()).hexdigest()[:12]
    assert path.name == f"bad_name.xlsx [{digest}].md"
    journal.append([f"workbook:{identity}"], "write", {"text": "one\ntwo\rthree"})
    assert len(path.read_text(encoding="utf-8").splitlines()) == 2


def test_twelve_digit_hash_separates_known_eight_digit_collision(office):
    first, second = r"C:\project8497\Book.xlsx", r"C:\project73998\Book.xlsx"
    assert hashlib.sha1(first.lower().encode()).hexdigest()[:8] == hashlib.sha1(second.lower().encode()).hexdigest()[:8]
    journal.append([f"workbook:{first}", f"workbook:{second}"], "write", {})
    assert journal.journal_path(first) != journal.journal_path(second)
    for identity in (first, second):
        assert journal.journal_path(identity).read_text(encoding="utf-8").splitlines()[0] == "# " + identity


def test_off_switch_and_warning_once(office, monkeypatch, capsys):
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, journal=False))
    journal.append(["workbook:Book"], "write", {})
    assert not config.SETTINGS.journal_dir.exists()
    blocker = office.tmp / "blocker"
    blocker.write_text("file", encoding="utf-8")
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, journal=True, journal_dir=blocker / "child"))
    monkeypatch.setattr(journal, "_warned", False)
    for _ in range(2):
        journal.append(["workbook:Book"], "write", {})
    assert capsys.readouterr().err.count("journal is not writable") == 1


def test_registry_journals_success_and_failure_actual_targets(office):
    o = office
    o.call("excel_write_range", workbook=o.wb.Name, sheet="Data", cells="A1", values=[[7]])
    with pytest.raises(ToolError):
        o.call("excel_clear_range", workbook=o.wb.Name, sheet="Data", cells="A1", what="invalid")
    path = journal.journal_path(o.wb.Name)
    lines = path.read_text(encoding="utf-8").splitlines()
    assert len(lines) == 3 and lines[1].endswith(" · ok") and "ошибка:" in lines[2]
    with pytest.raises(ToolError):
        o.call("excel_write_range", workbook="absent", sheet="Data", cells="A1", values=[[7]])
    assert len(list(path.parent.glob("*.md"))) == 1


def test_sheet_enable_append_disable_without_recursive_records(office):
    o = office
    o.call("office_journal", workbook=o.wb.Name, action="enable_sheet")
    assert o.wb.Names(journal.LOG_NAME).Visible is False
    assert o.wb.Sheets(o.wb.Sheets.Count).Name == "Лог"
    log = o.wb.Worksheets("Лог")
    assert log.Range("A1:E1").Font.Bold is True
    assert not undo.STACKS
    o.call("excel_write_range", workbook=o.wb.Name, sheet="Data", cells="A1", values=[[7]])
    key = undo.stack_key("workbook", o.excel, o.wb)
    assert len(undo.STACKS[key]) == 1
    lines = journal.journal_path(o.wb.Name).read_text(encoding="utf-8").splitlines()
    assert len(lines) == 3  # заголовок, enable_sheet, единственный вызов записи
    assert log.Range("B3").Value == "excel_write_range"
    o.call("office_undo", file=o.wb.Name)
    assert "отменено: excel_write_range" in log.Range("D4").Value
    assert not undo.STACKS[key]
    o.call("office_journal", workbook=o.wb.Name, action="disable_sheet")
    assert not journal.sheet_enabled(o.wb) and o.wb.Worksheets("Лог") is log
    o.call("office_journal", workbook=o.wb.Name, action="disable_sheet", delete_sheet=True)
    assert o.wb.Sheets.Count == 1 and o.excel.EnableEvents


def test_sheet_warning_does_not_fail_tool(office, monkeypatch):
    o = office
    o.wb.Names.Add(journal.LOG_NAME, "=TRUE", False)
    monkeypatch.setattr(o.wb.Sheets, "Add", lambda *a: (_ for _ in ()).throw(ValueError("protected workbook")))
    result = o.call("excel_write_range", workbook=o.wb.Name, sheet="Data", cells="A1", values=[[7]])
    assert result["log_sheet_warning"] == "protected workbook" and result["ok"]


def test_sheet_logging_forces_events_off_even_with_user_opt_in(office, monkeypatch):
    o = office
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, enable_events=True, undo=False))
    o.call("office_journal", workbook=o.wb.Name, action="enable_sheet")
    old = o.wb.Worksheets("Лог").Range

    def checked(address):
        assert not o.excel.EnableEvents
        return old(address)

    monkeypatch.setattr(o.wb.Worksheets("Лог"), "Range", checked)
    result = o.call("excel_write_range", workbook=o.wb.Name, sheet="Data", cells="A1", values=1)
    assert "log_sheet_warning" not in result and o.excel.EnableEvents


def test_journal_status_tail_and_word_only_disk(office):
    o = office
    assert o.call("office_journal", document=o.doc.Name, action="read")["lines"] == []
    o.word_write("first")
    o.word_write("second")
    status = o.call("office_journal", document=o.doc.Name, action="status")
    assert status["journal_enabled"] and status["sheet_enabled"] is False
    assert len(o.call("office_journal", document=o.doc.Name, lines=1)["lines"]) == 1
    with pytest.raises(ToolError, match="Excel-only"):
        o.call("office_journal", document=o.doc.Name, action="enable_sheet")


def test_history_pseudo_group_and_readonly_actions(office, monkeypatch):
    o = office
    assert "history" not in config.GROUPS and "history" not in config.PRESETS
    for group in ("excel_core", "word_core"):
        s = config.load({"OFFICE_LIVE_TOOLSETS": group, "OFFICE_LIVE_MODE": "readonly"})
        assert s.group_selected("history") and s.enabled("history", "read") and not s.enabled("history", "write")
    assert not config.load({"OFFICE_LIVE_TOOLSETS": "excel_format"}).group_selected("history")
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, mode="readonly"))
    assert o.call("office_undo", file=o.wb.Name, action="history")["history"] == []
    assert o.call("office_journal", workbook=o.wb.Name, action="status")["sheet_enabled"] is False
    for action in ("undo",):
        with pytest.raises(ToolError, match="Read-only"):
            o.call("office_undo", file=o.wb.Name, action=action)
    for action in ("enable_sheet", "disable_sheet"):
        with pytest.raises(ToolError, match="Read-only"):
            o.call("office_journal", workbook=o.wb.Name, action=action)


@pytest.mark.parametrize("kind", ["read", "write"])
def test_read_action_writing_file_is_journaled(office, monkeypatch, kind):
    wrappers = []
    monkeypatch.setattr(registry, "mcp", NS(add_tool=lambda fn, **kw: wrappers.append(fn)))
    audit = office.tmp / "audit.jsonl"
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, audit_log=audit))

    @registry.office_tool("excel_core", kind, read_actions=("export",), file_args=("export_path",))
    def export_test(workbook, action="export", export_path=""):
        com.note_target("workbook:" + workbook)
        return {"ok": True}

    wrappers[0](workbook="Book", export_path="fake-output")
    assert journal.journal_path("Book").exists() and not undo.STACKS
    assert len(audit.read_text(encoding="utf-8").splitlines()) == 2


@pytest.mark.parametrize("name,value", [
    ("OFFICE_LIVE_JOURNAL", "yes"), ("OFFICE_LIVE_UNDO", "invalid"), ("OFFICE_LIVE_UNDO_DEPTH", "0"),
    ("OFFICE_LIVE_UNDO_DEPTH", "x"), ("OFFICE_LIVE_UNDO_MAX_CELLS", "-1"),
])
def test_settings_validate(name, value):
    with pytest.raises(config.ConfigError):
        config.load({name: value})


def test_settings_defaults_and_overrides(tmp_path):
    s = config.load({})
    assert s.journal and s.undo and s.undo_depth == 50 and s.undo_max_cells == 200000
    s = config.load({"OFFICE_LIVE_JOURNAL": "off", "OFFICE_LIVE_UNDO": "off", "OFFICE_LIVE_JOURNAL_DIR": str(tmp_path), "OFFICE_LIVE_UNDO_DEPTH": "2"})
    assert not s.journal and not s.undo and s.journal_dir == tmp_path and s.undo_depth == 2
