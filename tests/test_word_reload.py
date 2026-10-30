"""Reload lifecycle and refusal checks through the real registry/undo wrappers; no Office."""

from dataclasses import replace
from types import SimpleNamespace as NS
from urllib.parse import parse_qs, urlsplit

import pytest
import pywintypes

from office_live import config, registry, undo, word_core
from office_live.errors import ToolError
from tests.history_fakes import Document, fake_office  # noqa: F401


@pytest.fixture
def reload_office(office):
    o = office
    path = o.tmp / "reload.docx"
    path.write_text("external text", encoding="utf-8")
    o.doc.Path, o.doc.FullName = str(o.tmp), str(path)
    o.doc.ActiveWindow = NS(Hwnd=84, WindowState=1, Left=10, Top=20, Width=800, Height=600, Visible=True)
    o.open_calls = []
    o.old_doc = o.doc

    def open_document(*args):
        assert not o.word.Documents.items
        assert o.word.AutomationSecurity == 3  # msoAutomationSecurityForceDisable
        o.open_calls.append(args)
        new = Document(o.word)
        new.Name, new.FullName, new.Path = "Reloaded.docx", str(path), str(o.tmp)
        new.ReadOnly = args[2]
        new.Content.Text = path.read_text(encoding="utf-8")
        new.ActiveWindow = NS(Hwnd=85)
        o.word.Documents.items.append(new)
        o.word.ActiveDocument = new
        o.doc = new
        return new

    o.word.Documents.Open = open_document
    return o


def reload(o, **kwargs):
    return o.call("word_reload_document", document=o.old_doc.FullName, **kwargs)


@pytest.mark.parametrize("read_only", [False, True])
def test_reload_reads_disk_preserves_mode_window_and_returns_new_links(reload_office, read_only):
    o = reload_office
    o.old_doc.ReadOnly = read_only
    o.word_write("previous agent edit")
    assert undo.STACKS
    o.old_doc.Saved = False
    starts = len(o.word.starts)
    result = reload(o, discard_changes=True)
    assert result["document"] == o.doc.Name == "Reloaded.docx"
    assert result["path"] == o.old_doc.FullName and result["read_only"] == read_only
    assert "history has been dropped" in result["undo"] and not undo.STACKS
    assert len(o.word.starts) == starts  # no UndoRecord around close/open
    assert o.word.AutomationSecurity == 1
    assert o.open_calls[0][:5] == (o.old_doc.FullName, False, read_only, False, "__office_live_no_password__")
    assert vars(o.doc.ActiveWindow) == {**vars(o.old_doc.ActiveWindow), "Hwnd": 85}
    assert result["links"][0]["label"] == o.doc.Name
    assert parse_qs(urlsplit(result["links"][0]["uri"]).query)["doc"] == [o.doc.FullName]
    assert o.call("word_read_document", document=o.doc.Name)["text"] == "external text"


def test_reload_unsaved_refusal_keeps_document_and_history(reload_office):
    o = reload_office
    o.word_write("unsaved local text")
    o.doc.Saved = False
    before = dict(undo.STACKS)
    with pytest.raises(ToolError, match="unsaved changes.*discard_changes=true"):
        reload(o)
    assert o.word.Documents.items == [o.old_doc] and not o.open_calls
    assert o.doc.Content.Text == "unsaved local text" and not o.doc.Saved
    assert undo.STACKS == before


def test_reload_saved_default_does_not_inspect_closed_target(reload_office, monkeypatch):
    o = reload_office

    def closed_target_must_not_be_read(obj):
        pytest.fail("Undo finalization inspected the old document instance")

    monkeypatch.setattr(undo, "_tracked_protection", closed_target_must_not_be_read)
    result = reload(o)
    assert result["ok"] and o.doc.Content.Text == "external text"


def test_reload_window_failure_is_partial_and_drops_old_history(reload_office):
    o = reload_office
    o.word_write("previous edit")
    original_open = o.word.Documents.Open

    class BrokenWindow:
        def __setattr__(self, name, value):
            raise pywintypes.com_error(-2147352567, "window unavailable", None, None)

    def open_with_broken_window(*args):
        doc = original_open(*args)
        doc.ActiveWindow = BrokenWindow()
        return doc

    o.word.Documents.Open = open_with_broken_window
    with pytest.raises(ToolError, match="interrupted after closing.*window unavailable"):
        reload(o)
    assert not undo.STACKS and o.doc in o.word.Documents.items
    assert o.doc.Content.Text == "external text"


@pytest.mark.parametrize("case", ["unsaved_path", "missing_file", "outside_allowed", "autosave", "protected"])
def test_reload_preflight_never_closes(reload_office, monkeypatch, case):
    o = reload_office
    if case == "unsaved_path":
        o.doc.Path = ""
    elif case == "missing_file":
        (o.tmp / "reload.docx").unlink()
    elif case == "outside_allowed":
        monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, allowed_dirs=(str(o.tmp / "allowed"),)))
    elif case == "autosave":
        o.doc.AutoSaveOn = True
    else:
        o.doc.ProtectionType = 3
    with pytest.raises(ToolError):
        reload(o, discard_changes=True)
    assert o.word.Documents.items == [o.old_doc] and not o.open_calls


def test_reload_reopen_failure_drops_history_and_reports_manual_recovery(reload_office):
    o = reload_office
    o.word_write("previous agent edit")

    def fail(*args):
        raise pywintypes.com_error(-2147352567, "cannot open", None, None)

    o.word.Documents.Open = fail
    with pytest.raises(ToolError, match="interrupted after closing.*manually") as exc:
        reload(o)
    assert o.old_doc.FullName in str(exc.value)
    assert not undo.STACKS and not o.word.Documents.items and o.word.AutomationSecurity == 1


def test_reload_close_refusal_keeps_history(reload_office):
    o = reload_office
    o.word_write("previous agent edit")
    before = dict(undo.STACKS)

    def fail(*args):
        raise pywintypes.com_error(-2147352567, "cannot close", None, None)

    o.doc.Close = fail
    with pytest.raises(ToolError, match="cannot close"):
        reload(o)
    assert undo.STACKS == before and not o.open_calls


def test_reload_requires_explicit_document_and_is_not_readonly_tool(reload_office, monkeypatch):
    o = reload_office
    with pytest.raises(ToolError, match="required"):
        o.call("word_reload_document", document="")
    info = registry.CATALOG["word_reload_document"]
    assert info.kind == "destructive" and not info.read_actions
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, mode="readonly"))
    registered = []
    monkeypatch.setattr(registry, "mcp", NS(add_tool=lambda *a, **kw: registered.append(kw)))
    registry.office_tool(info.group, info.kind)(word_core.word_reload_document)
    assert not registered  # test fresh registration, rather than a wrapper retained by the fake client
    assert not registry.CATALOG["word_reload_document"].registered
    assert o.word.Documents.items == [o.old_doc] and not o.open_calls
