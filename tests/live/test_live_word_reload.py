"""Reload acceptance on this run's marked temporary files only. Never run without owner approval."""

from pathlib import Path

import pytest

from tests.live.mcpclient import ToolFailed
from tests.live.office_cleanup import DialogWatch
from tests.live.test_live_office_operations import srv as isolated_server
from tests.live.test_live_stage1 import in_word

srv = isolated_server
pytestmark = pytest.mark.live


@pytest.fixture
def reload_doc(srv, doc, tmp):
    path = str(Path(tmp) / "reload.docx")
    srv.call("word_insert_text", document=doc, text="Original disk text")
    srv.call("word_save_as", document=doc, path=path)
    return path


def placement(path):
    return in_word(path, lambda d: (int(d.ActiveWindow.WindowState), bool(d.ActiveWindow.Visible),
                                   float(d.ActiveWindow.Left), float(d.ActiveWindow.Top),
                                   float(d.ActiveWindow.Width), float(d.ActiveWindow.Height)))


@pytest.mark.parametrize("read_only", [False, True])
def test_probe_reload_reverts_to_saved_file(srv, reload_doc, read_only):
    # Word locks an open file against writes and replacement (seen live on 2026-10-30), so the reload is checked as a
    # revert to the saved file, keeping read-only mode and window placement.
    path = reload_doc
    if read_only:
        srv.call("word_close_document", document=path)
        srv.call("word_open_document", path=path, read_only=True)
    before = placement(path)
    in_word(path, lambda d: d.Content.InsertAfter("Unsaved in-memory text"))
    with DialogWatch() as dialogs:
        result = srv.call("word_reload_document", document=path, discard_changes=True)
    assert not dialogs.seen, dialogs.seen
    assert result["path"] == path and result["read_only"] == read_only and result["links"]
    assert "history" in result["undo"]
    text = srv.call("word_read_document", document=path)["text"]
    assert "Original disk text" in text and "Unsaved in-memory text" not in text
    assert placement(path) == pytest.approx(before)


def test_probe_reload_unsaved_refusal_and_explicit_discard(srv, reload_doc):
    path = reload_doc
    srv.call("word_insert_text", document=path, text="Unsaved local edit")
    before = srv.call("word_read_document", document=path)["text"]
    with DialogWatch() as dialogs:
        with pytest.raises(ToolFailed, match="unsaved changes.*discard_changes=true"):
            srv.call("word_reload_document", document=path)
        assert srv.call("word_read_document", document=path)["text"] == before
        assert not in_word(path, lambda d: bool(d.Saved))
        srv.call("word_reload_document", document=path, discard_changes=True)
    assert not dialogs.seen, dialogs.seen
    text = srv.call("word_read_document", document=path)["text"]
    assert "Original disk text" in text and "Unsaved local edit" not in text
