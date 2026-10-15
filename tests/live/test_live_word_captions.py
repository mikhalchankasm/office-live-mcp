"""Written-only probes on marked temporary documents; never run without the owner's live window."""

from pathlib import Path

import pytest

from tests.live.mcpclient import ToolFailed
from tests.live.office_cleanup import DialogWatch
from tests.live.test_live_office_operations import srv as isolated_server
from tests.live.test_live_stage1 import in_word, word_state

srv = isolated_server
pytestmark = pytest.mark.live


@pytest.fixture
def captions_doc(srv, doc, tmp):
    path = str(Path(tmp) / "captions.docx")
    srv.call("word_insert_text", document=doc, text="See here\nEquation target")
    for i in range(2):
        srv.call("word_create_table", document=doc, data=[[str(i)]])
    srv.call("word_save_as", document=doc, path=path)
    in_word(path, lambda d: d.Range(0, 0).Select())
    return path


def test_caption_table_reference_selection_and_undo(srv, captions_doc):
    path = captions_doc
    def selection(d):
        r = d.Application.Selection.Range
        return int(r.Start), int(r.End)
    before, selected = word_state(path)[0], in_word(path, selection)
    local = in_word(path, lambda d: str(d.Application.CaptionLabels(-2).Name))  # wdCaptionTable
    with DialogWatch() as watch:
        result = srv.call("word_captions", document=path, action="add", target="table", table_index=2, text="Сводная по регионам")
        assert result["undo"] == "available"
        inventory = srv.call("word_captions", document=path)
        cap = inventory["captions"][0]
        assert (cap["label_id"], cap["label"], cap["number"]) == (-2, local, "1")
        assert cap["objects"] == [{"target": "table", "index": 2}] and inventory["tables_without_caption"] == [1]
        after_caption = word_state(path)[0]
        ref = srv.call("word_captions", document=path, action="cross_reference", caption=cap["caption"],
                       at="after_text", find_text="See here")
        assert ref["text"] == local + " 1" and ref["undo"] == "available"
        assert in_word(path, lambda d: any("\\h" in str(f.Code.Text) for f in d.Content.Fields if int(f.Type) == 3))  # wdFieldRef
        assert in_word(path, selection) == selected
        srv.call("office_undo", file=path)
        assert word_state(path)[0] == after_caption
        srv.call("office_undo", file=path)
        assert word_state(path)[0] == before  # word_state uses word_fingerprint(styles=False)
    assert not watch.seen


def test_caption_picture_equation_duplicate_and_number(srv, captions_doc, tmp):
    from PIL import Image
    path = captions_doc
    png = str(Path(tmp) / "caption.png")
    Image.new("RGB", (16, 16), "blue").save(png)
    srv.call("word_insert_image", document=path, path=png)
    srv.call("word_captions", document=path, action="add", target="picture", picture_index=1, text="Picture")
    before = word_state(path)[0]
    with pytest.raises(ToolFailed, match="adjacent"):
        srv.call("word_captions", document=path, action="add", target="picture", picture_index=1)
    assert word_state(path)[0] == before
    srv.call("word_captions", document=path, action="cross_reference", label="figure", caption=1, reference="number", paragraph=1)
    assert in_word(path, lambda d: any(str(f.Result.Text) == "1" for f in d.Content.Fields if int(f.Type) == 3))  # wdFieldRef
    srv.call("office_undo", file=path)
    assert word_state(path)[0] == before
    srv.call("word_captions", document=path, action="add", target="paragraph", paragraph=1, label="equation", text="Equation")
    result = srv.call("word_captions", document=path)
    assert any(c["label_id"] == -3 and c["orphan"] for c in result["captions"])  # wdCaptionEquation
    assert not result["pictures_without_caption"]


@pytest.mark.parametrize("reference", ["text", "page"])
def test_caption_reference_kinds_preserve_unrelated_field(srv, captions_doc, reference):
    path = captions_doc
    def sentinel(d):
        r = d.Range(int(d.Content.End) - 1, int(d.Content.End) - 1)
        field = d.Fields.Add(r, 33)  # wdFieldPage
        field.Result.Text = "stale sentinel"
    in_word(path, sentinel)
    srv.call("word_captions", document=path, action="add", target="table", table_index=2, text="Title")
    cap = srv.call("word_captions", document=path)["captions"][0]
    result = srv.call("word_captions", document=path, action="cross_reference", caption=1, reference=reference, paragraph=1)
    assert result["text"] == (cap["text"] if reference == "text" else "1")
    assert in_word(path, lambda d: [str(f.Result.Text) for f in d.Content.Fields if int(f.Type) == 33]) == ["stale sentinel"]


@pytest.mark.parametrize("caption", ["heading", "bookmark"])
def test_heading_and_bookmark_reference_undo(srv, captions_doc, caption):
    path = captions_doc
    def setup(d):
        d.Paragraphs(1).Style = -2  # wdStyleHeading1
        d.Bookmarks.Add("CaptionProbe", d.Paragraphs(1).Range)
    in_word(path, setup)
    before = word_state(path)[0]
    result = srv.call("word_captions", document=path, action="cross_reference", caption=caption,
                      heading=1, bookmark="CaptionProbe", paragraph=2)
    assert "See here" in result["text"] and result["label_id"] is None
    srv.call("office_undo", file=path)
    assert word_state(path)[0] == before
