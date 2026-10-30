"""word_find across stories and Russian wildcard counts (#55, #56) on marked temporary documents."""

import time
from pathlib import Path

import pytest

from tests.live.test_live_office_operations import srv as isolated_server
from tests.live.test_live_stage1 import in_word

srv = isolated_server
pytestmark = pytest.mark.live


@pytest.fixture
def find_doc(srv, doc, tmp):
    path = str(Path(tmp) / "find.docx")
    srv.call("word_save_as", document=doc, path=path)
    return path


def test_probe_find_all_stories_is_bounded(srv, find_doc):
    path = find_doc
    srv.call("word_insert_text", document=path, text="\n".join(f"Row {i}: 1660.00" for i in range(1, 41)))
    srv.call("word_insert_break", document=path, kind="section_next_page")
    srv.call("word_insert_text", document=path, text="Landscape 1660.00")
    srv.call("word_insert_break", document=path, kind="section_next_page")
    srv.call("word_insert_text", document=path, text="Last 1660.00")

    def prepare(doc):
        doc.Sections(2).PageSetup.Orientation = 1  # wdOrientLandscape
        for i in (1, 2, 3):
            doc.Sections(i).Footers(1).LinkToPrevious = False
            doc.Sections(i).Footers(1).Range.Text = f"Footer {i} 1660.00"
        doc.Sections(1).Headers(1).Range.Text = "Header 1660.00"
        for i in (2, 3):
            doc.Sections(i).Headers(1).LinkToPrevious = False
            doc.Sections(i).Headers(1).Range.Text = f"Header {i}"
        box = doc.Shapes.AddTextbox(1, 100, 100, 200, 30, doc.Paragraphs(1).Range)
        box.TextFrame.TextRange.Text = "Box 1660.00"
        table = doc.Tables.Add(doc.Paragraphs(5).Range, 2, 2)
        table.Cell(1, 1).Range.Text = "Cell 1660.00"

    in_word(path, prepare)
    main = srv.call("word_find", document=path, query="1660.00", max_results=500)
    started = time.monotonic()
    every = srv.call("word_find", document=path, query="1660.00", scope="all", max_results=500)
    assert time.monotonic() - started < 30
    assert "partial" not in every
    stories = [m["story"] for m in every["matches"] if m["story"] != "main"]
    assert sorted(stories) == ["primary_footer"] * 3 + ["primary_header", "textbox"], stories
    assert [m for m in every["matches"] if m["story"] == "main"] == main["matches"]


def test_probe_wildcard_counts_any_language(srv, find_doc):
    path = find_doc
    srv.call("word_insert_text", document=path, text="В1. Вопрос\nВ12. Вопрос\nВ123. Вопрос")
    found = srv.call("word_find", document=path, query="В[0-9]{1,2}.", wildcards=True)
    assert [m["text"] for m in found["matches"]] == ["В1.", "В12."]
    srv.call("word_replace_text", document=path, find="В([0-9]{1,2}).", replace="Q\\1.", wildcards=True)
    text = srv.call("word_read_document", document=path)["text"]
    assert "Q1. Вопрос" in text and "Q12. Вопрос" in text and "В123. Вопрос" in text


def test_probe_replace_25_pairs_one_undo(srv, find_doc):
    path = find_doc
    original = "\n".join(f"В{i}. Вопрос" for i in range(2, 27))
    srv.call("word_insert_text", document=path, text=original)
    before = srv.call("word_read_document", document=path)["text"]
    assert original in before
    pairs = [{"find": f"В{i}", "replace": f"В{i-1}"} for i in range(2, 27)]
    result = srv.call("word_replace_text", document=path, pairs=pairs, whole_word=True)
    assert result["replacements"] == 25 and result["undo"] == "available"
    assert result["pairs"] == [{**p, "count": 1} for p in pairs]
    after = srv.call("word_read_document", document=path)["text"]
    assert after == before.replace(original, "\n".join(f"В{i}. Вопрос" for i in range(1, 26)))
    undone = srv.call("office_undo", file=path)
    assert len(undone["undone"]) == 1
    assert srv.call("word_read_document", document=path)["text"] == before
