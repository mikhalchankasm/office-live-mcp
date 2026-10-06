"""Written-only Word review and generation probes on marked temporary files; never run without explicit approval."""

from pathlib import Path

import pytest

from tests.live.test_live_office_operations import srv as isolated_server

from tests.live.test_live_stage1 import in_word, word_state

srv = isolated_server

pytestmark = pytest.mark.live


@pytest.fixture
def review_doc(srv, doc, tmp):
    path = str(Path(tmp) / "review.docx")
    srv.call("word_insert_text", document=doc, text="Before target after\nSecond numbered paragraph", style="List Number")
    srv.call("word_save_as", document=doc, path=path)
    return path


def test_probe_tracked_fragment_style_numbering_selection_and_undo(srv, review_doc):
    path = review_doc
    def state(doc):
        selection = doc.Application.Selection.Range
        return (bool(doc.TrackRevisions), int(doc.Paragraphs.Count),
                [(str(doc.Paragraphs(i).Range.Style.NameLocal), str(doc.Paragraphs(i).Range.ListFormat.ListString))
                 for i in range(1, int(doc.Paragraphs.Count) + 1)], int(selection.Start), int(selection.End))
    before = in_word(path, state)
    fingerprint = word_state(path)[0]
    args = {"document": path, "paragraph": 1, "expected_text": "target", "new_text": "replacement"}
    preview = srv.call("word_edit_preview", **args)
    assert not preview["blocked"] and word_state(path)[0] == fingerprint
    result = srv.call("word_edit_apply", **args, fingerprint=preview["fingerprint"])
    assert result["revisions_added"] > 0 and result["undo"] == "available" and not result["warnings"]
    assert in_word(path, state) == before
    assert "replacement" in in_word(path, lambda doc: doc.Paragraphs(1).Range.Text)
    srv.call("office_undo", file=path)
    assert word_state(path)[0] == fingerprint


def test_probe_comment_threads_reply_key_date_and_context(srv, review_doc):
    path = review_doc
    srv.call("word_manage_comments", document=path, action="add", find_text="target", text="Please clarify")
    first = srv.call("word_manage_comments", document=path, limit=50)["comments"][0]
    srv.call("word_manage_comments", document=path, action="reply", index=first["index"], text="Agreed", thread_key=first["thread_key"])
    result = srv.call("word_manage_comments", document=path, limit=50)
    root = next(c for c in result["comments"] if c["index"] == first["index"])
    assert root["replies"] and root["replies"][0]["text"] == "Agreed"
    assert root["thread_key"] == first["thread_key"] and root["context_before"] and root["location"]["paragraph"] == 1
    if "Date" not in result["unavailable_properties"]:
        assert root["date"]


def test_probe_generation_preview_run_split_placeholders_and_no_files(srv, review_doc, wb, tmp):
    path = review_doc
    def prepare(doc):
        doc.Content.Text = "{{Code}} {{Name}}\r"
        doc.Range(0, 4).Font.Bold = True  # deliberately split {{Code}} across runs when saved
        doc.Sections(1).Headers(1).Range.Text = "{{Name}}"
        doc.Save()
    in_word(path, prepare)
    book = str(Path(tmp) / "preview.xlsx")
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["Code", "Name"], ["007", "Act"], ["0008", "act"]], value_mode="text")
    srv.call("excel_save_as", workbook=wb, path=book)
    # Saved + text: any edit clears Saved; the full XML fingerprint drifts with Word's background work in long runs.
    untouched = lambda: in_word(path, lambda doc: (bool(doc.Saved), doc.Content.Text))  # noqa: E731
    before = untouched()
    out = Path(tmp) / "planned_only"
    result = srv.call("bridge_excel_to_word_preview", workbook=book, sheet="Data", cells="A1:B3", template=path,
                      output_dir=str(out), filename_column="Name", export_pdf=True)
    assert result["ready"] and result["placeholders_analyzed"] and result["matched"] == ["Code", "Name"]
    assert result["sample"][0]["values"]["Code"] == "007" and len(result["renamed"]) == 1
    assert not out.exists() and untouched() == before and before[0] is True
