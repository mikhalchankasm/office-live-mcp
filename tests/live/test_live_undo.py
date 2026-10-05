"""Отмена на собственных черновиках. Не запускается без OFFICE_LIVE_LIVE_TESTS=1."""

import pytest

from tests.live.conftest import read

pytestmark = pytest.mark.live


def test_live_write_range_undo(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["007", "=1+1"]], value_mode="text")
    before = read(srv, wb, "A1:B1")
    result = srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[[8, 9]])
    assert result["undo"] == "available"
    srv.call("office_undo", file=wb)
    assert read(srv, wb, "A1:B1") == before


def test_live_format_range_undo(srv, wb):
    before = srv.call("excel_get_format", workbook=wb, sheet="Data", cells="A1")
    result = srv.call("excel_format_range", workbook=wb, sheet="Data", cells="A1", bold=True, fill_color="#FFFF00")
    assert result["undo"] == "available"
    srv.call("office_undo", file=wb)
    assert srv.call("excel_get_format", workbook=wb, sheet="Data", cells="A1") == before


def test_live_insert_rows_undo(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[[1], [2], [3]])
    srv.call("excel_insert_rows_columns", workbook=wb, sheet="Data", axis="rows", start=2, count=2)
    srv.call("office_undo", file=wb)
    assert read(srv, wb, "A1:A3") == [[1], [2], [3]]


def test_live_delete_sheet_undo(srv, wb):
    srv.call("excel_add_worksheet", workbook=wb, name="Other")
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values="007", value_mode="text")
    srv.call("excel_delete_sheet", workbook=wb, sheet="Data", confirm=True)
    result = srv.call("office_undo", file=wb)
    assert "#REF!" in result["warning"] and read(srv, wb, "A1") == [["007"]]


def test_live_word_insert_text_undo(srv, doc):
    before = srv.call("word_read_document", document=doc)
    result = srv.call("word_insert_text", document=doc, text="Undo this paragraph.", position="end")
    assert result["undo"] == "available"
    srv.call("office_undo", file=doc)
    assert srv.call("word_read_document", document=doc) == before
