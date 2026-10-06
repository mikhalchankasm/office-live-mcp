"""Mail merge planning and partial generation on fakes and synthetic OOXML only."""

import os
import zipfile
from dataclasses import replace
from pathlib import Path
from types import SimpleNamespace as NS

import pytest

from office_live import bridge, com, config, journal, registry, undo
from office_live.errors import ToolError
from tests.history_fakes import fake_office  # noqa: F401


def template(path, body=None):
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("word/document.xml", body or f'<w:document {ns}><w:body><w:p><w:r><w:t>{{{{Co</w:t></w:r><w:r><w:t>de}}}}</w:t></w:r></w:p></w:body></w:document>')
        package.writestr("word/header1.xml", f'<w:hdr {ns}><w:p><w:r><w:t>{{{{ Name }}}}</w:t></w:r></w:p></w:hdr>')
        package.writestr("word/footer1.xml", f'<w:ftr {ns}><w:p><w:r><w:t>{{{{Missing}}}}</w:t></w:r></w:p></w:ftr>')
    return str(path)


@pytest.fixture
def merge(office, monkeypatch):
    data = [["Name", "Code", "Extra"], ["Alice", "007", "x"], ["alice", "0002", "y"], ["", "", ""], ["Bob", "00103", "z"]]
    office.ws.Range("B4:D8").Value2 = data
    monkeypatch.setattr(bridge, "read_grid", lambda rng, mode: data)
    tpl = template(office.tmp / "template.docx")
    args = {"workbook": office.wb.Name, "sheet": "Data", "cells": "B4:D8", "template": tpl,
            "output_dir": str(office.tmp / "output"), "filename_column": "Name", "export_pdf": True}
    return office, args, data


def test_preview_no_files_journal_undo_or_com_mutations(merge, monkeypatch):
    o, args, _ = merge
    before = sorted(o.tmp.rglob("*"))
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, audit_log=o.tmp / "audit.jsonl"))
    monkeypatch.setattr(registry, "_audit_write", lambda *a, **k: pytest.fail("preview audit"))
    monkeypatch.setattr(journal, "append", lambda *a, **k: pytest.fail("preview journal"))
    monkeypatch.setattr(undo, "require_undo", lambda *a, **k: pytest.fail("preview undo"))
    monkeypatch.setattr(com, "primary_app", lambda *a, **k: pytest.fail("Word launch"))
    monkeypatch.setattr(bridge.os, "makedirs", lambda *a, **k: pytest.fail("mkdir"))
    result = o.call("bridge_excel_to_word_preview", **args)
    assert result["ready"] and result["placeholders_analyzed"] and result["total_documents"] == 3
    assert result["matched"] == ["Code", "Name"] and result["unused_columns"] == ["Extra"]
    assert result["unfilled_placeholders"] == ["Missing"]
    assert result["blank_rows"] == [7] and result["sample"][0]["values"]["Code"] == "007"
    assert [p["filename"] for p in result["files"]] == ["Alice.docx", "alice_002.docx", "Bob.docx"]
    assert result["renamed"] == [{"row": 2, "wanted": "alice", "used": "alice_002"}]
    assert result["files"][0]["pdf"] == os.path.join(args["output_dir"], "Alice.pdf")
    assert not result["preview_is_permission"] and sorted(o.tmp.rglob("*")) == before
    assert not undo.STACKS and not o.word.starts
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, mode="readonly"))
    registered = {}
    monkeypatch.setattr(registry, "mcp", NS(add_tool=lambda fn, name, **kw: registered.__setitem__(name, kw)))
    registry.office_tool("bridge", "read")(bridge.bridge_excel_to_word_preview)
    assert registered["bridge_excel_to_word_preview"]["annotations"].read_only_hint


def test_preview_existing_files_and_word_query_without_launch(merge, monkeypatch):
    o, args, _ = merge
    out = Path(args["output_dir"])
    out.mkdir()
    existing = out / "Alice.docx"
    existing.write_text("existing", encoding="utf-8")
    o.doc.FullName, o.doc.Path = str(existing), str(out)
    calls = []
    def apps(kind, launch=False):
        calls.append((kind, launch))
        assert not launch
        return [o.excel if kind == "excel" else o.word]
    monkeypatch.setattr(com, "apps", apps)
    result = o.call("bridge_excel_to_word_preview", **args)
    assert not result["ready"] and len(result["problems"]) == 2
    assert result["existing_files"] == [{"path": str(existing), "open_in_word": True}]
    assert not o.call("bridge_excel_to_word_preview", **{**args, "overwrite": True})["ready"]
    with pytest.raises(ToolError, match="Nothing was created"):
        o.call("bridge_excel_to_word_documents", **args)
    assert not any(launch for _, launch in calls) and existing.read_text(encoding="utf-8") == "existing"


@pytest.mark.parametrize("kind", ["doc", "unsaved", "dirty", "malformed"])
def test_unknown_template_placeholder_analysis(merge, kind):
    o, args, _ = merge
    if kind == "doc":
        path = o.tmp / "old.doc"
        path.write_bytes(b"synthetic")
        args["template"] = str(path)
    elif kind == "malformed":
        Path(args["template"]).write_bytes(b"bad zip")
    else:
        args["template"] = o.doc.Name
        if kind == "dirty":
            o.doc.Path, o.doc.FullName, o.doc.Saved = str(o.tmp), str(o.tmp / "template.docx"), False
    result = o.call("bridge_excel_to_word_preview", **args)
    assert not result["placeholders_analyzed"] and result["placeholders_analysis_reason"]
    if kind == "unsaved":
        assert not result["ready"]
    assert not Path(args["output_dir"]).exists()


def generator(o, monkeypatch, fail=0, fail_pdf=False, fail_after_write=False):
    calls = []
    app = NS(AutomationSecurity=1, DisplayAlerts=1, NormalTemplate=NS(Saved=False), Quit=lambda save=0: calls.append("quit"))
    def add(path, /, *visibility):  # private hidden Word: Add(Template, NewTemplate, DocumentType, Visible)
        idx = len([c for c in calls if c != "quit"]) + 1
        calls.append(path)
        def save(path, format, /):
            assert app.DisplayAlerts == 0 and format == 16
            if idx == fail and not fail_pdf and not fail_after_write:
                raise RuntimeError("fake save failure")
            Path(path).write_text("synthetic document", encoding="utf-8")
            if idx == fail and fail_after_write:
                raise RuntimeError("failure after writing")
        def export(path, *args):
            assert args == (17, False, 0, 0, 0, 0, 0)
            if idx == fail and fail_pdf:
                raise RuntimeError("fake PDF failure")
            Path(path).write_bytes(b"synthetic PDF")
        return NS(SaveAs2=save, ExportAsFixedFormat=export, Close=lambda save: None)
    app.Documents = NS(Add=add)
    from tests.window_fakes import FakeWin32

    monkeypatch.setattr(com, "private_app", lambda kind: app)  # generation runs in a separate hidden Word (0.4.3)
    monkeypatch.setattr(bridge, "Win32", FakeWin32)
    monkeypatch.setattr(bridge, "fill_placeholders", lambda *a: {"unfilled_placeholders": []})
    return calls


def test_generation_rebuilds_plan_and_keeps_legacy_success_fields(merge, monkeypatch):
    o, args, data = merge
    preview = o.call("bridge_excel_to_word_preview", **args)
    data[1][0] = "Changed"
    calls = generator(o, monkeypatch)
    result = o.call("bridge_excel_to_word_documents", **args)
    assert result["ok"] and result["files"][0] == "Changed.docx"
    assert result["files"][0] != preview["files"][0]["filename"]
    assert result["skipped_blank_rows"] == 1 and len([c for c in calls if c != "quit"]) == 3 and calls[-1] == "quit"
    assert len(result["created_files"]) == len(result["created_pdfs"]) == 3
    assert all(os.path.isabs(p) and Path(p).exists() for p in result["created_files"] + result["created_pdfs"])


@pytest.mark.parametrize("pdf", [False, True])
def test_partial_failure_returns_all_full_paths_over_twenty(merge, monkeypatch, pdf):
    o, args, data = merge
    data[:] = [["Name", "Code", "Extra"], *[[f"row{i}", "007", "x"] for i in range(1, 28)]]
    args["cells"] = "A1:C28"
    generator(o, monkeypatch, fail=25, fail_pdf=pdf)
    result = o.call("bridge_excel_to_word_documents", **args)
    assert result["ok"] is False and result["stopped_at_row"] == 25 and "failure" in result["error"]
    assert len(result["created_files"]) == (25 if pdf else 24) and len(result["created_pdfs"]) == 24
    assert all(Path(p).exists() and os.path.isabs(p) for p in result["created_files"] + result["created_pdfs"])


@pytest.mark.parametrize("kw", [{"max_documents": 1}, {"max_documents": 0}, {"left_delimiter": ""}])
def test_preview_limits_before_generation(merge, monkeypatch, kw):
    o, args, _ = merge
    monkeypatch.setattr(com, "primary_app", lambda *a, **k: pytest.fail("launch"))
    with pytest.raises(ToolError):
        o.call("bridge_excel_to_word_preview", **{**args, **kw})
    assert not Path(args["output_dir"]).exists()



def test_partial_failure_after_file_creation_and_error_journal(merge, monkeypatch):
    o, args, _ = merge
    generator(o, monkeypatch, fail=1, fail_after_write=True)
    result = o.call("bridge_excel_to_word_documents", **args)
    assert not result["ok"] and len(result["created_files"]) == 1
    assert Path(result["created_files"][0]).exists()
    assert any("failure after writing" in p.read_text(encoding="utf-8") for p in (o.tmp / "journal").glob("*.md"))
