"""Written-only watermark probes on marked temporary documents; never run without the owner's live window."""

from pathlib import Path

import pytest

from tests.live.mcpclient import ToolFailed
from tests.live.test_live_office_operations import srv as isolated_server
from tests.live.test_live_stage1 import in_word, word_state

srv = isolated_server
pytestmark = pytest.mark.live


@pytest.fixture
def watermark_doc(srv, doc, tmp):
    path = str(Path(tmp) / "watermark.docx")
    srv.call("word_save_as", document=doc, path=path)
    return path


def test_probe_watermark_status_leaves_empty_headers_untouched(srv, watermark_doc):
    path = watermark_doc
    before = word_state(path)
    assert not srv.call("word_watermark", document=path)["present"]
    assert srv.call("word_watermark", document=path, action="remove")["changed"] == []
    assert word_state(path) == before


def test_probe_watermark_linked_headers_shapes_and_undo(srv, watermark_doc):
    path = watermark_doc
    srv.call("word_insert_text", document=path, text="First section")
    srv.call("word_insert_break", document=path, kind="section_next_page")
    srv.call("word_insert_text", document=path, text="Second section")
    def prepare(doc):
        for i in (1, 2):
            doc.Sections(i).PageSetup.DifferentFirstPageHeaderFooter = True
            doc.Sections(i).PageSetup.OddAndEvenPagesHeaderFooter = True
        for typ in (1, 2, 3):
            hf = doc.Sections(1).Headers(typ)
            hf.Range.Text = "Header stays"
            shape = hf.Shapes.AddTextEffect(0, "Logo stays", "Calibri", 12, 0, 0, 0, 0, hf.Range)
            shape.Name = "WatermarkProbeLogo" + str(typ)
            doc.Sections(2).Headers(typ).LinkToPrevious = True
    in_word(path, prepare)
    before = word_state(path)[0]
    result = srv.call("word_watermark", document=path, action="add", text="ЧЕРНОВИК")
    assert len(result["changed"]) == 3 and result["undo"] == "available"
    status = srv.call("word_watermark", document=path)
    assert len(status["watermarks"]) == 3
    assert all(m["text"] == "ЧЕРНОВИК" and m["sections"] == [1, 2] for m in status["watermarks"])
    added = word_state(path)[0]
    for args in ({"action": "add", "text": "Again"}, {"action": "remove", "sections": [2]}):
        with pytest.raises(ToolFailed):
            srv.call("word_watermark", document=path, **args)
        assert word_state(path)[0] == added
    srv.call("word_watermark", document=path, action="remove")
    assert word_state(path)[0] == before
    srv.call("office_undo", file=path)
    assert word_state(path)[0] == added
    srv.call("office_undo", file=path)
    assert word_state(path)[0] == before


def test_probe_watermark_add_undo_removes_new_headers(srv, watermark_doc):
    path = watermark_doc
    from office_live import undo

    before = in_word(path, undo.watermark_fingerprint)
    before_parts = in_word(path, undo._word_body_xml)
    srv.call("word_watermark", document=path, action="add", text="Confidential", layout="horizontal")
    try:
        restored = srv.call("office_undo", file=path)
    except ToolFailed:
        import os
        folder = os.environ.get("OFFICE_LIVE_DIAG_DIR")
        if folder:  # evidence for the failed verification: parts before the add and after the undo attempt
            for label, parts in (("before", before_parts), ("after", in_word(path, undo._word_body_xml))):
                for name, body in parts:
                    target = Path(folder) / "watermark" / label / name.strip("/").replace("/", "_")
                    target.parent.mkdir(parents=True, exist_ok=True)
                    target.write_text(body, encoding="utf-8")
                (Path(folder) / "watermark" / f"{label}-counts.txt").write_text(repr(in_word(path, lambda d: (
                    d.Content.Text, int(d.Paragraphs.Count), int(d.Shapes.Count), int(d.ProtectionType),
                    bool(d.TrackRevisions), int(d.Comments.Count)))), encoding="utf-8")
        raise
    assert not restored.get("warnings"), restored
    # Word keeps a header once created; empty headers are invisible and excluded from this comparison.
    assert in_word(path, undo.watermark_fingerprint) == before
    assert not srv.call("word_watermark", document=path)["present"]
