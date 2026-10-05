"""Opt-in navigation in OUR fixtures only. No ShellExecute or registry registration."""

import os
import subprocess
import sys
from pathlib import Path

import pytest

from office_live.links import build

pytestmark = pytest.mark.live
REPO = Path(__file__).resolve().parents[2]


def dispatch(uri, expected=0):
    result = subprocess.run([sys.executable, "-m", "office_live", "open-link", uri], cwd=REPO,
                            env={**os.environ, "OFFICE_LIVE_LINK_NO_DIALOG": "1"}, capture_output=True, timeout=40)
    assert result.returncode == expected, result.stderr.decode("utf-8", errors="replace")


def test_link_selects_own_workbook_range(srv, wb):
    before = srv.call("excel_read_range", workbook=wb, sheet="Data", cells="B2:C4")["values"]
    uri = srv.call("office_link", workbook=wb, sheet="Data", cells="B2:C4")["links"][0]["uri"]
    dispatch(uri)
    selected = srv.call("excel_get_selection")
    assert selected["workbook"] == wb and selected["sheet"] == "Data"
    assert selected["selection"].replace("$", "") == "B2:C4"
    assert srv.call("excel_read_range", workbook=wb, sheet="Data", cells="B2:C4")["values"] == before


def test_link_selects_own_document_paragraph(srv, doc):
    inserted = srv.call("word_insert_text", document=doc, text="Link test first\nLink test second\nLink test third")
    before = srv.call("word_read_document", document=doc)
    paragraph = inserted["paragraphs"][0] + 1
    uri = srv.call("office_link", document=doc, paragraph=paragraph)["links"][0]["uri"]
    dispatch(uri)
    selected = srv.call("word_get_selection", document=doc)
    assert selected["document"] == doc and "Link test second" in selected["text"]
    assert srv.call("word_read_document", document=doc) == before


@pytest.mark.parametrize("app", ["excel", "word"])
def test_closed_own_file_is_refused_without_opening(srv, wb, doc, tmp, app):
    # Save and close only our fixture; an existing file on disk must still never be opened by a link.
    target = Path(tmp) / ("closed.xlsx" if app == "excel" else "closed.docx")
    if app == "excel":
        srv.call("excel_save_as", workbook=wb, path=str(target))
        srv.call("excel_close_workbook", workbook=str(target), discard=True)
        uri = build(app, book=str(target), sheet="Data", range="A1")
        list_tool, collection = "excel_list_workbooks", "workbooks"
    else:
        srv.call("word_save_as", document=doc, path=str(target))
        srv.call("word_close_document", document=str(target), discard=True)
        uri = build(app, doc=str(target), paragraph=1)
        list_tool, collection = "word_list_documents", "documents"
    before = srv.call(list_tool)[collection]
    dispatch(uri, expected=1)
    after = srv.call(list_tool)[collection]
    assert {(item["name"], item["path"]) for item in after} == {(item["name"], item["path"]) for item in before}
    assert target.is_file()
