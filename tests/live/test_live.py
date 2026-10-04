"""Живые тесты: реальные Excel и Word через настоящий MCP-протокол (stdio).

Запуск:  set OFFICE_LIVE_LIVE_TESTS=1 && .venv\\Scripts\\python -m pytest tests/live -q
Тесты создают СВОИ черновые книги/документы и закрывают только их (discard); приложения не закрываются, кроме Word,
если его запустили сами тесты и в нём не осталось документов.
"""

import base64
import os
import tempfile

import pytest

from tests.live.mcpclient import Stdio, ToolFailed
from tests.live.office_cleanup import quit_word_if_idle

pytestmark = pytest.mark.live


@pytest.fixture(scope="session")
def srv():
    c = Stdio()
    yield c
    c.close()
    quit_word_if_idle()


def _is_ours(name, original, path):
    """Наш файл: исходное имя черновика или файл во временной папке тестов (имя меняется после Save As)."""
    return name == original or (path and "ol_pytest_" in path)


@pytest.fixture
def wb(srv):
    name = srv.call("excel_new_workbook", sheets=["Data"])["workbook"]
    yield name
    for w in srv.call("excel_list_workbooks")["workbooks"]:
        if _is_ours(w["name"], name, w["path"]):
            srv.call("excel_close_workbook", workbook=w["path"] + "\\" + w["name"] if w["path"] else w["name"], discard=True)


@pytest.fixture
def doc(srv):
    name = srv.call("word_new_document")["document"]
    yield name
    for d in srv.call("word_list_documents")["documents"]:
        if _is_ours(d["name"], name, d["path"]):
            srv.call("word_close_document", document=d["path"] + "\\" + d["name"] if d["path"] else d["name"], discard=True)


@pytest.fixture
def tmp():
    d = tempfile.mkdtemp(prefix="ol_pytest_")
    yield d
    import shutil

    shutil.rmtree(d, ignore_errors=True)


def read(srv, wb, cells, sheet="Data", **kw):
    return srv.call("excel_read_range", workbook=wb, sheet=sheet, cells=cells, **kw)["values"]


# ================================================================== Excel


def test_write_read_and_text_mode(srv, wb):
    r = srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["id", "val"], ["007", "12%"], ["x", 5]], value_mode="text")
    assert r["written_shape"] == [3, 2] and r["overwritten_nonempty_cells"] == 0
    assert read(srv, wb, "A1:B3") == [["id", "val"], ["007", "12%"], ["x", 5]]  # строки остались текстом, число — числом
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="D1", values=[["12%", "2026-10-04", "=1+1"]])  # режим auto
    row = read(srv, wb, "D1:F1")[0]
    assert row[0] == pytest.approx(0.12) and row[1] == "2026-10-04" and row[2] == 2


def test_single_cell_expands_and_ragged_rows_are_padded(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="B2", values=[[1, 2, 3], [4]])
    assert read(srv, wb, "B2:D3") == [[1, 2, 3], [4, None, None]]
    with pytest.raises(ToolFailed, match="Pass just the top-left cell"):
        srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A10:B10", values=[[1, 2, 3]])
    with pytest.raises(ToolFailed, match="already contains"):
        srv.call("excel_write_range", workbook=wb, sheet="Data", cells="B2", values=[[9]], overwrite=False)


def test_formulas_fill_down_and_r1c1(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[[1, 10], [2, 20], [3, 30]])
    r = srv.call("excel_set_formula", workbook=wb, sheet="Data", cells="C1:C3", formula="=A1*B1")
    assert [v[0] for v in r["values"]] == [10, 40, 90]
    srv.call("excel_set_formula", workbook=wb, sheet="Data", cells="D1:D3", formula="=RC[-3]+RC[-2]", r1c1=True)
    assert [v[0] for v in read(srv, wb, "D1:D3")] == [11, 22, 33]
    bad = srv.call("excel_set_formula", workbook=wb, sheet="Data", cells="E1", formula="=1/0")
    assert bad["errors"] and bad["errors"][0]["error"] == "#DIV/0!"
    with pytest.raises(ToolFailed, match="parse the formula"):
        srv.call("excel_set_formula", workbook=wb, sheet="Data", cells="E2", formula="=SUM(")


def test_errors_are_actionable(srv, wb):
    with pytest.raises(ToolFailed, match="Invalid range"):
        srv.call("excel_read_range", workbook=wb, sheet="Data", cells="ZZZZ99999")
    with pytest.raises(ToolFailed, match="not found"):
        srv.call("excel_read_range", workbook=wb, sheet="NoSuchSheet", cells="A1")
    with pytest.raises(ToolFailed, match="unsaved changes"):
        srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[[1]])
        srv.call("excel_close_workbook", workbook=wb)
    with pytest.raises(ToolFailed, match="confirm=true"):
        srv.call("excel_delete_sheet", workbook=wb, sheet="Data")


def test_sort_filter_find_replace_dedupe(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["Name", "Qty"], ["b", 2], ["a", 5], ["c", 1], ["a", 5]])
    srv.call("excel_sort_range", workbook=wb, sheet="Data", cells="A1:B5", keys=[{"column": "Qty", "order": "desc"}])
    assert [r[1] for r in read(srv, wb, "A2:B5")] == [5, 5, 2, 1]
    r = srv.call("excel_filter_range", workbook=wb, sheet="Data", cells="A1:B5", filters=[{"column": "Qty", "criteria": ">1"}])
    assert r["visible_data_rows"] == 3
    srv.call("excel_filter_range", workbook=wb, sheet="Data", clear="remove")
    assert srv.call("excel_find", workbook=wb, sheet="Data", query="A", whole_cell=True)["total_matches"] == 2
    rep = srv.call("excel_replace", workbook=wb, sheet="Data", find="a", replace="z", whole_cell=True)
    assert rep["cells_changed"] == 2
    d = srv.call("excel_remove_duplicates", workbook=wb, sheet="Data", cells="A1:B5")
    assert d["duplicate_rows_removed"] == 1


def test_hide_rows_columns_by_number_and_condition(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["Task", "Status"], ["a", "Done"], ["b", "Open"], ["c", "Done"]])
    r = srv.call("excel_hide_rows_columns", workbook=wb, sheet="Data", where={"column": "Status", "operator": "equals", "value": "Done"})
    assert r["affected"] == ["2:2", "4:4"] and r["count"] == 2
    srv.call("excel_hide_rows_columns", workbook=wb, sheet="Data", lines="2-4", hidden=False)
    assert srv.call("excel_hide_rows_columns", workbook=wb, sheet="Data", axis="columns", lines="B")["affected"] == ["B:B"]
    with pytest.raises(ToolFailed, match="exactly one"):
        srv.call("excel_hide_rows_columns", workbook=wb, sheet="Data", lines="1", where={"column": "A"})


def test_number_formats_roundtrip_in_any_locale(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[[1234.5]])
    for fmt in ("0.00", "#,##0.00", "0.0%", "dd.mm.yyyy", '0.0 "kg"'):
        srv.call("excel_format_range", workbook=wb, sheet="Data", cells="A1", number_format=fmt)
        assert srv.call("excel_get_format", workbook=wb, sheet="Data", cells="A1")["number_format"] == fmt
    srv.call("excel_format_range", workbook=wb, sheet="Data", cells="A1", number_format="thousands_2")
    shown = srv.call("excel_read_range", workbook=wb, sheet="Data", cells="A1", display_text=True)["values"][0][0]
    assert "234" in shown and shown.replace(" ", " ").replace(" ", "")[:4] in ("1234", "1,23")
    with pytest.raises(ToolFailed, match="rejected number_format"):
        srv.call("excel_format_range", workbook=wb, sheet="Data", cells="A1", number_format="[[[bad")


def test_conditional_format_and_validation(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[[5, "late"], [50, "ok"], [500, "ok"]])
    srv.call("excel_conditional_format", workbook=wb, sheet="Data", cells="A1:A3", rule="cell_value", operator="less", value1=100, fill_color="#FFC7CE")
    srv.call("excel_conditional_format", workbook=wb, sheet="Data", cells="A1:B3", rule="formula", formula='=$B1="late"', fill_color="lightyellow", bold=True)
    rules = srv.call("excel_conditional_format", workbook=wb, sheet="Data", action="list")["rules"]
    assert any(r["rule"] == "cell_value" and r.get("fill") == "#FFC7CE" for r in rules)
    formula_rule = next(r for r in rules if r["rule"] == "formula")
    assert formula_rule["formula"].replace(" ", "") == '=$B1="late"' and formula_rule["applies_to"] == "A1:B3"  # английская запись
    srv.call("excel_data_validation", workbook=wb, sheet="Data", cells="B1:B3", items=["ok", "late", "n/a"])
    got = srv.call("excel_data_validation", workbook=wb, sheet="Data", cells="B1", action="get")
    assert got["has_validation"] and got["type"] == 3
    srv.call("excel_data_validation", workbook=wb, sheet="Data", cells="B1:B3", action="clear")
    assert not srv.call("excel_data_validation", workbook=wb, sheet="Data", cells="B1", action="get")["has_validation"]


def test_render_range_image_is_a_real_png(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["a", "b"], [1, 2]])
    srv.call("excel_format_range", workbook=wb, sheet="Data", cells="A1:B1", bold=True, fill_color="#1F4E78", font_color="white")
    res = srv.raw_call("excel_render_range_image", {"workbook": wb, "sheet": "Data", "cells": "A1:B2"})
    images = [c for c in res["content"] if c["type"] == "image"]
    assert images and base64.b64decode(images[0]["data"])[:8] == b"\x89PNG\r\n\x1a\n"


def test_pivot_filters_slicers_and_hiding(srv, wb):
    rows = [["Region", "Product", "Qty", "Revenue"], ["N", "A", 10, 100], ["N", "B", 5, 80], ["S", "A", 7, 70], ["S", "B", 12, 150], ["E", "A", 3, 30], ["W", "B", 8, 90]]
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=rows)
    p = srv.call("excel_create_pivot_table", workbook=wb, source="A1:D7", source_sheet="Data", rows=["Region"], values=[{"field": "Revenue", "function": "sum", "number_format": "#,##0.00"}], dest_sheet="P", name="PT")
    totals = {r[0]: r[1] for r in p["preview"][1:] if r[0]}
    assert totals["N"] == 180 and totals["S"] == 220
    only = srv.call("excel_pivot_filter", workbook=wb, field="Region", items=["N", "E"], mode="only")
    assert sorted(only["visible_items"]) == ["E", "N"]
    srv.call("excel_pivot_filter", workbook=wb, field="Region", items=["E"], mode="hide")
    assert srv.call("excel_pivot_info", workbook=wb)["fields"][0]["hidden_items"] == 3
    srv.call("excel_pivot_filter", workbook=wb, field="Region", mode="all")
    top = srv.call("excel_pivot_filter", workbook=wb, field="Region", action="top", operator="top_items", value1=2, data_field="Sum of Revenue")
    assert sorted(top["visible_items"]) == ["N", "S"]
    lab = srv.call("excel_pivot_filter", workbook=wb, field="Region", action="label", operator="begins_with", value1="S")
    assert lab["visible_items"] == ["S"]
    srv.call("excel_pivot_filter", workbook=wb, action="clear")
    sl = srv.call("excel_manage_slicers", workbook=wb, action="add", pivot="PT", field="Region", anchor_cell="H3", width=180, height=120)
    assert sl["slicers"][0]["width"] == pytest.approx(180)
    srv.call("excel_manage_slicers", workbook=wb, action="select", slicer="Region", items=["N", "S"])
    sel = {i["name"]: i["selected"] for i in srv.call("excel_manage_slicers", workbook=wb, action="list")["slicers"][0]["items"]}
    assert sel == {"E": False, "N": True, "S": True, "W": False}
    srv.call("excel_manage_slicers", workbook=wb, action="delete", slicer="Region")
    fld = srv.call("excel_pivot_fields", workbook=wb, action="add", field="Product", orientation="column")
    assert fld["ok"]
    changed = srv.call("excel_pivot_fields", workbook=wb, action="data_field", field="Sum of Revenue", function="average", number_format="0.0", new_caption="Avg Rev")["changed"]
    assert "function=average" in changed and "caption=Avg Rev" in changed and "number_format" in changed
    assert [d["caption"] for d in srv.call("excel_pivot_info", workbook=wb)["data_fields"]] == ["Avg Rev"]


def test_chart_table_profile_and_issues(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["k", "v"], ["a", 1], ["b", 2], ["c", 3]])
    srv.call("excel_create_chart", workbook=wb, sheet="Data", source="A1:B4", chart_type="column", title="T", name="Ch", anchor_cell="E2")
    assert srv.call("excel_manage_charts", workbook=wb, action="list")["charts"][0]["name"] == "Ch"
    res = srv.raw_call("excel_manage_charts", {"workbook": wb, "action": "export_image", "chart_name": "Ch"})
    assert any(c["type"] == "image" for c in res["content"])
    tbl = srv.call("excel_manage_tables", workbook=wb, action="create", sheet="Data", cells="A1:B4", name="T1")
    assert tbl["columns"] == ["k", "v"]
    prof = srv.call("excel_profile_range", workbook=wb, sheet="Data")
    assert prof["columns"][1]["numeric"]["sum"] == 6
    srv.call("excel_manage_tables", workbook=wb, action="to_range", name="T1")
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="D1", values=[["12"], [" pad "]], value_mode="text")
    srv.call("excel_set_formula", workbook=wb, sheet="Data", cells="F1:F4", formula="=B1*2")
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="F4", values=[[99]])
    issues = {i["type"] for i in srv.call("excel_find_issues", workbook=wb, sheet="Data")["issues"]}
    assert {"number_stored_as_text", "stray_spaces", "hardcoded_number_among_formulas"} <= issues


def test_save_as_guards_and_template_copy(srv, wb, tmp):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["x", 1]])
    path = os.path.join(tmp, "book.xlsx")
    assert srv.call("excel_save_as", workbook=wb, path=path)["ok"]
    with pytest.raises(ToolFailed, match="already exists"):
        srv.call("excel_save_as", workbook="book.xlsx", path=path)
    with pytest.raises(ToolFailed, match="extension"):
        srv.call("excel_save_as", workbook="book.xlsx", path=os.path.join(tmp, "evil.exe"))
    srv.call("excel_close_workbook", workbook="book.xlsx")
    copy = os.path.join(tmp, "copy.xlsx")
    created = srv.call("excel_create_from_template", template_path=path, new_path=copy)
    try:
        assert created["workbook"] == "copy.xlsx" and os.path.getsize(copy) > 0
        with pytest.raises(ToolFailed, match="extension"):
            srv.call("excel_create_from_template", template_path=path, new_path=os.path.join(tmp, "x.xlsm"))
    finally:
        srv.call("excel_close_workbook", workbook="copy.xlsx", discard=True)


def test_sample_workflow_on_synthetic_template(srv, wb, tmp):
    """Разобрать образец и сделать аналог: образец собираем сами (шапка с объединением, заморозка, формулы, правило)."""
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["Report title"], [None], ["Code", "Name", "Due", "Done%", "Next"]])
    srv.call("excel_merge_cells", workbook=wb, sheet="Data", cells="A1:E1")
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A4", values=[["1", "Section", None, 0, None], ["1.1", "Item", "2020-01-01", 10, None], ["1.2", "Item 2", "2099-01-01", 20, None]], value_mode="text")
    srv.call("excel_set_formula", workbook=wb, sheet="Data", cells="E4:E6", formula='=IF(C4="","",C4+1)')
    srv.call("excel_conditional_format", workbook=wb, sheet="Data", cells="A4:E100", rule="formula", formula='=LEN($A4)-LEN(SUBSTITUTE($A4,".",""))=0', fill_color="#A9D08E", bold=True)
    srv.call("excel_sheet_view", workbook=wb, sheet="Data", freeze_at="A4")
    sample = os.path.join(tmp, "sample.xlsx")
    srv.call("excel_save_as", workbook=wb, path=sample)
    before = os.path.getmtime(sample)
    insp = srv.call("office_inspect_file", path=sample, peek_rows=3)
    s0 = insp["sheets"][0]
    assert s0["merged_ranges"]["count"] == 1 and s0["freeze"]["ySplit"] == "3" and s0["conditional_formats"]["count"] == 1
    assert s0["formula_cells"] == 3 and insp["has_vba"] is False
    new = os.path.join(tmp, "analog.xlsx")
    name = srv.call("excel_create_from_template", template_path=sample, new_path=new)["workbook"]
    try:
        lay = srv.call("excel_describe_layout", workbook=name)
        assert lay["header_rows"] == 3 and lay["view"]["frozen"] and lay["data_rows"]["first"] == 4
        assert lay["title_rows"][0]["cells"][0]["text"] == "Report title"
        cols = {c["column"]: c for c in lay["columns"]}
        assert cols["A"]["header"] == "Code" and cols["E"]["formula_r1c1"]["cells"] == 3
        rule = lay["conditional_formats"][0]
        assert rule["fill"] == "#A9D08E" and rule["bold"] and "SUBSTITUTE" in rule["formula"]
        assert lay["row_kinds"]["hierarchy"]["levels"] == {"0": 1, "1": 2}
        srv.call("excel_clear_range", workbook=name, sheet="", cells="A4:E6", what="contents")
        srv.call("excel_write_range", workbook=name, sheet="", cells="A4", values=[["7", "New section", None, 0], ["7.1", "New item", "2021-05-05", 5], ["7.2", "Another", "2099-02-02", 6], ["7.3", "Last", "2099-03-03", 7]], value_mode="text")
        srv.call("excel_set_formula", workbook=name, sheet="", cells="E4:E7", formula=cols["E"]["formula_r1c1"]["formula"], r1c1=True)
        values = srv.call("excel_read_range", workbook=name, sheet="", cells="A4:E7")["values"]
        assert [r[0] for r in values] == ["7", "7.1", "7.2", "7.3"] and values[0][4] in (None, "")
        assert values[1][4] not in (None, "")  # формула пересчитана на каждой строке
    finally:
        srv.call("excel_close_workbook", workbook=name, discard=True)
    assert os.path.getmtime(sample) == before  # образец не тронут


# ================================================================== Word


def test_word_insert_structure_find_literal_replace(srv, doc):
    srv.call("word_insert_text", document=doc, text="Report", style="Title")
    srv.call("word_insert_text", document=doc, text="Intro", style="Heading 1")
    ins = srv.call("word_insert_text", document=doc, text="Revenue grew.\nRevenue fell.")
    assert ins["paragraphs"] == [3, 4]
    st = srv.call("word_get_structure", document=doc)
    assert st["paragraphs"] == 4 and st["headings"] == [{"paragraph": 2, "level": 1, "text": "Intro"}]
    assert srv.call("word_find", document=doc, query="revenue")["total_matches"] == 2
    r = srv.call("word_replace_text", document=doc, find="Revenue", replace="R ^ \\1 rev", match_case=True)
    assert r["replacements"] == 2
    text = srv.call("word_read_document", document=doc)["text"]
    assert "R ^ \\1 rev grew." in text  # ^ и \ остались буквальными
    srv.call("word_replace_text", document=doc, find="R ^ \\1 rev", replace="revenue")
    assert "revenue grew." in srv.call("word_read_document", document=doc)["text"]  # регистр замены НЕ подгоняется под найденное
    paras = srv.call("word_read_document", document=doc, mode="paragraphs")["paragraphs"]
    assert paras[1]["heading_level"] == 1
    chunk = srv.call("word_read_document", document=doc, max_chars=10)
    assert chunk["truncated"] and chunk["next_offset"] == 10


def test_word_insert_positions_and_delete(srv, doc):
    srv.call("word_insert_text", document=doc, text="one\ntwo\nthree")
    srv.call("word_insert_text", document=doc, text="after one", position="after_paragraph", paragraph=1)
    srv.call("word_insert_text", document=doc, text="before one", position="before_paragraph", paragraph=1)
    srv.call("word_insert_text", document=doc, text="TWO!", position="replace_paragraph", paragraph=4)
    texts = [p["text"] for p in srv.call("word_read_document", document=doc, mode="paragraphs")["paragraphs"]]
    assert texts == ["before one", "one", "after one", "TWO!", "three"]
    srv.call("word_delete_paragraphs", document=doc, start_paragraph=1, end_paragraph=2)
    assert srv.call("word_get_structure", document=doc)["paragraphs"] == 3
    with pytest.raises(ToolFailed, match="does not exist"):
        srv.call("word_insert_text", document=doc, text="x", position="after_paragraph", paragraph=99)


def test_word_tables_stay_separate_and_multiline_cells(srv, doc):
    srv.call("word_insert_text", document=doc, text="Intro")
    srv.call("word_create_table", document=doc, data=[["Item", "Qty"], ["Kiwi\n(green)", 2]])
    srv.call("word_create_table", document=doc, data=[["A", "B"], [1, 2]], position="after_paragraph", paragraph=1)  # вплотную к первой
    tables = srv.call("word_list_tables", document=doc)["tables"]
    assert [(t["rows"], t["columns"], t["nested"]) for t in tables] == [(2, 2, 0), (2, 2, 0)]
    assert srv.call("word_read_table", document=doc, table_index=2)["values"][1] == ["Kiwi\n(green)", "2"]
    srv.call("word_write_table", document=doc, table_index=1, values=[["x", "y"], ["z", "w"]], start_row=2)
    assert srv.call("word_read_table", document=doc, table_index=1)["rows"] == 3
    srv.call("word_modify_table", document=doc, table_index=1, action="merge_cells", row=1, col=1, to_row=1, to_col=2)
    assert srv.call("word_read_table", document=doc, table_index=1)["uniform"] is False
    srv.call("word_format_table", document=doc, table_index=2, row_from=2, fill_color="#E2EFDA", bold=True, borders="all")
    assert srv.call("word_modify_table", document=doc, table_index=2, action="delete_table")["tables_left"] == 1


def test_word_layout_headers_toc_comments_bookmarks_revisions(srv, doc, tmp):
    srv.call("word_insert_text", document=doc, text="Title", style="Heading 1")
    srv.call("word_insert_text", document=doc, text="Body text with a website.")
    srv.call("word_insert_break", document=doc, kind="page", position="end")
    srv.call("word_insert_text", document=doc, text="Next page")
    assert srv.call("word_manage_toc", document=doc, action="insert", position="start")["tocs"] == 1
    srv.call("word_headers_footers", document=doc, action="set", kind="footer", text="ACME", page_numbers="page_of_total", page_label="P. ", of_label="/")
    footer = srv.call("word_headers_footers", document=doc, action="get")["sections"][0]["footer_primary"]
    assert footer.startswith("ACME") and "P." in footer and "/" in footer
    srv.call("word_manage_comments", document=doc, action="add", find_text="website", text="check")
    srv.call("word_manage_comments", document=doc, action="reply", index=1, text="ok")
    assert srv.call("word_manage_comments", document=doc)["count"] == 2
    srv.call("word_manage_bookmarks", document=doc, action="add", name="Spot", find_text="Next page")
    srv.call("word_insert_text", document=doc, text="Filled", position="bookmark", bookmark="Spot", new_paragraph=False)
    assert srv.call("word_manage_bookmarks", document=doc)["bookmarks"][0]["text"] == "Filled"
    srv.call("word_track_changes", document=doc, action="enable")
    srv.call("word_replace_text", document=doc, find="Body", replace="Main")
    assert srv.call("word_track_changes", document=doc, action="list")["revisions"] == 2
    assert srv.call("word_track_changes", document=doc, action="accept_all")["revisions_left"] == 0
    with pytest.raises(ToolFailed, match="never been saved"):
        srv.call("word_export_pdf", document=doc, path=os.path.join(tmp, "a.pdf"))  # Word обнуляет имя несохранённого документа
    path = os.path.join(tmp, "doc.docx")
    srv.call("word_save_as", document=doc, path=path)
    pdf = srv.call("word_export_pdf", document="doc.docx", path=os.path.join(tmp, "a.pdf"))
    assert pdf["size_bytes"] > 1000
    with pytest.raises(ToolFailed, match="already exists"):
        srv.call("word_save_as", document="doc.docx", path=path)
    srv.call("word_close_document", document="doc.docx", discard=True)
    reopened = srv.call("word_open_document", path=path)
    try:
        assert reopened["document"] == "doc.docx"
        insp = srv.call("office_inspect_file", path=path)
        assert insp["kind"] == "word" and insp["comments"] >= 1
    finally:
        srv.call("word_close_document", document="doc.docx", discard=True)


# ================================================================== мосты


def test_bridge_word_table_to_excel_numbers_and_codes(srv, wb, doc):
    srv.call("word_create_table", document=doc, data=[["Item", "Price", "Share", "Code"], ["Pump", "1 234,50", "12%", "007"], ["Valve", "99,9", "8,5%", "0012"]])
    r = srv.call("bridge_word_table_to_excel", document=doc, table_index=1, workbook=wb, sheet="Data", decimal_separator=",")
    assert r["readback"][1] == ["Pump", 1234.5, 0.12, "007"] and r["readback"][2][3] == "0012"
    with pytest.raises(ToolFailed, match="already contains"):
        srv.call("bridge_word_table_to_excel", document=doc, table_index=1, workbook=wb, sheet="Data", decimal_separator=",")


def test_bridge_excel_to_word_table_picture_and_chart(srv, wb, doc):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["Region", "Revenue"], ["N", 1234.5], ["S", 70], ["E", 30000]])
    srv.call("excel_format_range", workbook=wb, sheet="Data", cells="B2:B4", number_format="#,##0.00")
    srv.call("excel_hide_rows_columns", workbook=wb, sheet="Data", lines="3")
    info = srv.call("bridge_excel_range_to_word_table", workbook=wb, sheet="Data", cells="A1:B4", document=doc)
    assert info["rows_copied"] == 3 and info["hidden_skipped"]["rows"] == 1
    cells = srv.call("word_read_table", document=doc, table_index=1)["values"]
    assert [r[0] for r in cells] == ["Region", "N", "E"] and "234" in cells[1][1]
    srv.call("excel_create_chart", workbook=wb, sheet="Data", source="A1:B4", name="C1", anchor_cell="E2")
    srv.call("bridge_excel_chart_to_word", workbook=wb, chart_name="C1", document=doc, caption="Fig. 1")
    srv.call("bridge_excel_range_to_word_picture", workbook=wb, sheet="Data", cells="A1:B4", document=doc)
    assert srv.call("word_get_structure", document=doc)["inline_images"] == 2


def test_bridge_mail_merge(srv, wb, tmp):
    tpl = srv.call("word_new_document", text="Dear {{Name}}, order {{Order}}. {{Nothing}}")["document"]
    tpl_path = os.path.join(tmp, "tpl.docx")
    srv.call("word_save_as", document=tpl, path=tpl_path)
    srv.call("word_close_document", document="tpl.docx", discard=True)
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["Name", "Order"], ["Ivanov", "A-1"], ["Petrov", "A-2"]], value_mode="text")
    r = srv.call("bridge_excel_to_word_documents", workbook=wb, sheet="Data", cells="A1:B3", template=tpl_path, output_dir=os.path.join(tmp, "out"), filename_column="Name")
    assert r["documents_created"] == 2 and r["unfilled_placeholders"] == ["Nothing"]
    assert sorted(os.listdir(os.path.join(tmp, "out"))) == ["Ivanov.docx", "Petrov.docx"]
    srv.call("word_open_document", path=os.path.join(tmp, "out", "Ivanov.docx"))
    try:
        assert "Dear Ivanov, order A-1." in srv.call("word_read_document", document="Ivanov.docx")["text"]
    finally:
        srv.call("word_close_document", document="Ivanov.docx", discard=True)


# ================================================================== режимы сервера (без Office)


def test_readonly_server_exposes_no_writing_tools():
    c = Stdio(env={"OFFICE_LIVE_MODE": "readonly"})
    try:
        names = {t["name"] for t in c.list_tools()}
        assert "excel_read_range" in names and "word_read_document" in names
        assert not names & {"excel_write_range", "excel_save", "word_insert_text", "excel_delete_sheet", "excel_create_from_template"}
    finally:
        c.close()


def test_toolset_selection_and_eval_flag():
    c = Stdio(env={"OFFICE_LIVE_TOOLSETS": "excel_core", "OFFICE_LIVE_ALLOW_EVAL": "1"})
    try:
        names = {t["name"] for t in c.list_tools()}
        assert "excel_write_range" in names and "word_insert_text" not in names and "office_run_python" in names
    finally:
        c.close()


def test_strict_target_requires_explicit_names():
    c = Stdio(env={"OFFICE_LIVE_STRICT_TARGET": "1"})
    try:
        with pytest.raises(ToolFailed, match="name is required"):
            c.call("excel_write_range", workbook="", sheet="", cells="A1", values=[[1]])
        with pytest.raises(ToolFailed, match="name is required"):
            c.call("word_insert_text", document="", text="x")
    finally:
        c.close()


def test_audit_log_records_writes_not_reads(tmp):
    import json

    log = os.path.join(tmp, "audit.jsonl")
    c = Stdio(env={"OFFICE_LIVE_AUDIT_LOG": log})
    name = c.call("excel_new_workbook", sheets=["S"])["workbook"]
    try:
        c.call("excel_write_range", workbook=name, sheet="S", cells="A1", values=[[1]])
        c.call("excel_read_range", workbook=name, sheet="S", cells="A1")
        with pytest.raises(ToolFailed):
            c.call("excel_delete_sheet", workbook=name, sheet="S")  # без confirm — отказ тоже попадает в журнал
    finally:
        c.call("excel_close_workbook", workbook=name, discard=True)
        c.close()
    rows = [json.loads(line) for line in open(log, encoding="utf-8")]
    tools = [r["tool"] for r in rows]
    assert "excel_write_range" in tools and "excel_read_range" not in tools
    refused = next(r for r in rows if r["tool"] == "excel_delete_sheet")
    assert refused["ok"] is False and "confirm" in refused["error"]


# ================================================================== улучшения по замечаниям живого агента


def test_render_works_when_another_sheet_is_active(srv, wb):
    """Регрессия: CopyPicture рисует только видимое — раньше снимок неактивного листа выходил пустым."""
    srv.call("excel_add_worksheet", workbook=wb, name="Other", activate=True)
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["Region", "Revenue"], ["N", 1234], ["S", 99]])
    srv.call("excel_format_range", workbook=wb, sheet="Data", cells="A1:B1", bold=True, fill_color="#1F4E78", font_color="white")
    res = srv.raw_call("excel_render_range_image", {"workbook": wb, "sheet": "Data", "cells": "A1:B3"})
    png = base64.b64decode(next(c for c in res["content"] if c["type"] == "image")["data"])
    assert png[:8] == b"\x89PNG\r\n\x1a\n" and len(png) > 2500  # не пустая картинка
    # пользовательский вид восстановлен: активным остался лист Other
    assert srv.call("excel_list_workbooks")["active_sheet"] == "Other"


def test_write_range_reports_cell_types(srv, wb):
    r = srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["text", 5, "2026-10-04", "=1/0", True]])
    assert r["readback_types"] == {"text": 1, "number": 1, "date": 1, "error": 1, "bool": 1}


def test_find_issues_checks_adjacent_blocks_independently(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["Name", "Qty"], ["a", 1], ["b", 2]])
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="D1", values=[["Name", "Total"], ["a", 1], ["b", 2]])  # те же заголовки, другой блок
    types = {i["type"] for i in srv.call("excel_find_issues", workbook=wb, sheet="Data")["issues"]}
    assert "duplicate_header" not in types


def test_chart_axis_format_limits_and_colors(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["k", "v"], ["a", 1000.5], ["b", 2500.25], ["c", 4000]])
    srv.call("excel_create_chart", workbook=wb, sheet="Data", source="A1:B4", name="Ch", anchor_cell="E2", value_axis_number_format="#,##0", value_axis_min=0, value_axis_max=5000, series_colors=["#C00000"], data_label_number_format="#,##0")
    out = srv.call("excel_manage_charts", workbook=wb, action="update", chart_name="Ch", value_axis_number_format="0.0", data_labels=True, series_colors=["#1F4E78"])
    assert any(a.startswith("value_axis_number_format") for a in out["applied"]) and any(a.startswith("series_colors") for a in out["applied"])


def test_word_insert_next_to_tables_without_counting_paragraphs(srv, doc):
    srv.call("word_insert_text", document=doc, text="Intro")
    srv.call("word_create_table", document=doc, data=[["A", "B"], [1, 2]])
    srv.call("word_insert_text", document=doc, text="Tail")
    srv.call("word_insert_text", document=doc, text="Right after the table", position="after_table", table_index=1)
    srv.call("word_insert_text", document=doc, text="Right before the table", position="before_table", table_index=1, style="Heading 2")
    paras = srv.call("word_read_document", document=doc, mode="paragraphs", end_paragraph=30)["paragraphs"]
    texts = [p["text"] for p in paras if not p.get("in_table")]
    assert texts == ["Intro", "Right before the table", "Right after the table", "Tail", ""][: len(texts)] or texts[:4] == ["Intro", "Right before the table", "Right after the table", "Tail"]
    assert next(p for p in paras if p["text"] == "Right before the table").get("heading_level") == 2
    assert srv.call("word_list_tables", document=doc)["tables"][0]["nested"] == 0
    srv.call("word_create_table", document=doc, data=[["X"], ["y"]], position="after_table", table_index=1)
    assert [t["first_cell"] for t in srv.call("word_list_tables", document=doc)["tables"]] == ["A", "X"]
    with pytest.raises(ToolFailed, match="table_index"):
        srv.call("word_insert_text", document=doc, text="x", position="after_table", table_index=9)


def test_bridge_copies_cell_looks_including_conditional_formatting(srv, wb, doc):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["Region", "Revenue"], ["N", 40000], ["S", 90000]])
    srv.call("excel_format_range", workbook=wb, sheet="Data", cells="A1:B1", bold=True, fill_color="#1F4E78", font_color="white")
    srv.call("excel_conditional_format", workbook=wb, sheet="Data", cells="B2:B3", rule="cell_value", operator="less", value1=50000, fill_color="#FF9999")
    info = srv.call("bridge_excel_range_to_word_table", workbook=wb, sheet="Data", cells="A1:B3", document=doc)
    assert info["cells_with_copied_look"] >= 3
    plain = srv.call("bridge_excel_range_to_word_table", workbook=wb, sheet="Data", cells="A1:B3", document=doc, keep_formatting=False, position="after_table", table_index=1)
    assert plain["cells_with_copied_look"] == 0
    page = srv.raw_call("word_render_page_image", {"document": doc, "page": 1, "max_width_px": 800})
    png = base64.b64decode(next(c for c in page["content"] if c["type"] == "image")["data"])
    assert png[:8] == b"\x89PNG\r\n\x1a\n" and len(png) > 5000
    with pytest.raises(ToolFailed, match="does not exist"):
        srv.call("word_render_page_image", document=doc, page=99)


# ================================================================== замечания второго прогона агента (образец -> аналог)


def test_write_range_accepts_a_scalar_and_fills_the_range(srv, wb):
    r = srv.call("excel_write_range", workbook=wb, sheet="Data", cells="B2:C3", values=0)
    assert r["readback"] == [[0, 0], [0, 0]] and r["written_shape"] == [1, 1]


def test_find_issues_without_header_row_skips_header_checks(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A5", values=[["x", "x"], ["y", None]])  # диапазон начинается с данных
    with_hdr = {i["type"] for i in srv.call("excel_find_issues", workbook=wb, sheet="Data", cells="A5:B6")["issues"]}
    no_hdr = {i["type"] for i in srv.call("excel_find_issues", workbook=wb, sheet="Data", cells="A5:B6", has_header=False)["issues"]}
    assert "duplicate_header" in with_hdr and "duplicate_header" not in no_hdr


def test_layout_reports_hidden_columns_that_drive_conditional_formatting(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["Code", "Name", "Flag"], ["1", "a", " "], ["1.1", "b", None], ["1.2", "c", " "]], value_mode="text")
    srv.call("excel_conditional_format", workbook=wb, sheet="Data", cells="A2:C4", rule="formula", formula='=AND($C2="",$A2<>"")', fill_color="#FFC000")
    srv.call("excel_hide_rows_columns", workbook=wb, sheet="Data", axis="columns", lines="C")
    lay = srv.call("excel_describe_layout", workbook=wb, sheet="Data", header_rows=1)
    rule = lay["conditional_formats"][0]
    assert [d["column"] for d in rule["depends_on"]] == ["C", "A"] or {d["column"] for d in rule["depends_on"]} == {"A", "C"}
    assert next(d for d in rule["depends_on"] if d["column"] == "C")["hidden"] is True
    assert lay["hidden_columns_drive_formatting"]["columns"] == ["C"]


def test_render_survives_a_minimized_workbook_window():
    """Регрессия живого прогона: свёрнутое окно книги давало пустой снимок. Окно сворачиваем через office_run_python."""
    c = Stdio(env={"OFFICE_LIVE_ALLOW_EVAL": "1"})
    name = c.call("excel_new_workbook", sheets=["Data"])["workbook"]
    try:
        c.call("excel_write_range", workbook=name, sheet="Data", cells="A1", values=[["Region", "Revenue"], ["North", 1234], ["South", 99]])
        c.call("excel_format_range", workbook=name, sheet="Data", cells="A1:B1", bold=True, fill_color="#1F4E78", font_color="white")
        c.call("office_run_python", workbook=name, code="wb.Windows(1).WindowState = -4140")  # xlMinimized
        res = c.raw_call("excel_render_range_image", {"workbook": name, "sheet": "Data", "cells": "A1:B3"})
        png = base64.b64decode(next(i for i in res["content"] if i["type"] == "image")["data"])
        assert len(png) > 2000
        state = c.call("office_run_python", workbook=name, code="result = wb.Windows(1).WindowState")["result"]
        assert state == -4140  # состояние окна возвращено как было
    finally:
        c.call("excel_close_workbook", workbook=name, discard=True)
        c.close()
