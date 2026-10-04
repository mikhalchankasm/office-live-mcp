"""Живые регрессии аудита: по одному тесту на исправленный дефект. Работают только с собственными черновыми книгами/документами.

Запуск:  set OFFICE_LIVE_LIVE_TESTS=1 && .venv\\Scripts\\python -m pytest tests/live -q
"""

import json
import os

import pytest

from tests.live.conftest import read
from tests.live.mcpclient import Stdio, ToolFailed

pytestmark = pytest.mark.live


def _probe(code, workbook="", document=""):
    """Читает свойства через аварийный люк (отдельный сервер с OFFICE_LIVE_ALLOW_EVAL=1)."""
    c = Stdio(env={"OFFICE_LIVE_ALLOW_EVAL": "1"})
    try:
        return c.call("office_run_python", workbook=workbook, document=document, code=code)["result"]
    finally:
        c.close()


def _close(srv, *paths):
    for p in paths:
        try:
            srv.call("excel_close_workbook", workbook=p, discard=True)
        except ToolFailed:
            pass


# ================================================================== скрытые изменения и побочные эффекты


def test_helper_operations_leave_no_scratch_objects_and_never_hide_a_change(srv, wb, tmp):
    """Перевод формул не делает книгу «грязной»; снимок оставляет книгу изменённой (флаг не подменяется), но без лишних объектов."""
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["n", "v"], ["a", 1], ["b", 12]])
    srv.call("excel_conditional_format", workbook=wb, sheet="Data", cells="A2:B3", rule="formula", formula='=AND($B2>10,$A2<>"")', fill_color="#FFC000")
    path = os.path.join(tmp, "clean.xlsx")
    srv.call("excel_save_as", workbook=wb, path=path)
    before = srv.call("excel_workbook_info", workbook=path)
    assert before["saved"] is True
    books_before = sorted(w["path"] + w["name"] for w in srv.call("excel_list_workbooks")["workbooks"])
    srv.call("excel_conditional_format", workbook=path, sheet="Data", action="list")  # английский текст формулы — через временную книгу
    srv.call("excel_describe_layout", workbook=path, sheet="Data")
    assert srv.call("excel_workbook_info", workbook=path)["saved"] is True, "formula translation must not make the workbook dirty"
    srv.raw_call("excel_render_range_image", {"workbook": path, "sheet": "Data", "cells": "A1:B3"})
    after = srv.call("excel_workbook_info", workbook=path)
    assert after["saved"] is False  # временная диаграмма пометила книгу изменённой; сервер этого не маскирует
    assert srv.call("excel_manage_charts", workbook=path, action="list")["charts"] == []  # а сама диаграмма удалена
    assert after["sheets"] == before["sheets"]
    assert sorted(w["path"] + w["name"] for w in srv.call("excel_list_workbooks")["workbooks"]) == books_before  # временная книга закрыта


def test_readonly_server_exposes_listing_actions_but_refuses_changes(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["x", "y"], [1, 2]])
    srv.call("excel_create_chart", workbook=wb, sheet="Data", source="A1:B2", name="Chart1")
    ro = Stdio(env={"OFFICE_LIVE_MODE": "readonly"})
    try:
        names = {t["name"] for t in ro.list_tools()}
        assert {"excel_manage_names", "excel_manage_charts", "word_manage_comments"} <= names  # многоактные: доступны на чтение
        assert not {"excel_new_workbook", "word_new_document", "excel_sheet_view", "excel_write_range"} & names
        assert "names" in ro.call("excel_manage_names", workbook=wb, action="list")
        assert [c["name"] for c in ro.call("excel_manage_charts", workbook=wb, action="list")["charts"]] == ["Chart1"]
        with pytest.raises(ToolFailed, match="Read-only mode"):
            ro.call("excel_manage_names", workbook=wb, action="add", name="X", sheet="Data", cells="A1")
        with pytest.raises(ToolFailed, match="Read-only mode"):
            ro.call("excel_manage_charts", workbook=wb, action="delete", chart_name="Chart1")
        with pytest.raises(ToolFailed, match="read-only"):
            ro.call("excel_manage_charts", workbook=wb, action="export_image", chart_name="Chart1", export_path=os.path.join(os.environ.get("TEMP", "."), "ol_pytest_never.png"))
    finally:
        ro.close()
    assert [c["name"] for c in srv.call("excel_manage_charts", workbook=wb, action="list")["charts"]] == ["Chart1"]  # диаграмма цела


def test_allowed_dirs_apply_to_already_open_documents(srv, tmp):
    inside, outside = os.path.join(tmp, "inside"), os.path.join(tmp, "outside")
    os.makedirs(inside)
    os.makedirs(outside)
    a_path, b_path = os.path.join(outside, "a.xlsx"), os.path.join(inside, "b.xlsx")
    a = srv.call("excel_new_workbook", sheets=["S"])["workbook"]
    srv.call("excel_save_as", workbook=a, path=a_path)
    b = srv.call("excel_new_workbook", sheets=["S"])["workbook"]
    srv.call("excel_save_as", workbook=b, path=b_path)
    guarded = Stdio(env={"OFFICE_LIVE_ALLOWED_DIRS": inside})
    try:
        names = [w["name"] for w in guarded.call("excel_list_workbooks")["workbooks"]]
        assert "b.xlsx" in names and "a.xlsx" not in names  # чужое даже не видно в списке
        with pytest.raises(ToolFailed, match="not found"):
            guarded.call("excel_read_range", workbook="a.xlsx", sheet="S", cells="A1")
        with pytest.raises(ToolFailed, match="not found"):
            guarded.call("excel_save", workbook="a.xlsx")
        guarded.call("excel_write_range", workbook="b.xlsx", sheet="S", cells="A1", values=[[1]])
        assert guarded.call("excel_save", workbook="b.xlsx")["saved"] is True
    finally:
        guarded.close()
        _close(srv, a_path, b_path)


def test_selection_and_active_names_of_a_forbidden_workbook_are_not_revealed(srv, tmp):
    inside, outside = os.path.join(tmp, "inside"), os.path.join(tmp, "outside")
    os.makedirs(inside)
    os.makedirs(outside)
    secret_path, ok_path = os.path.join(outside, "secret.xlsx"), os.path.join(inside, "ok.xlsx")
    ok = srv.call("excel_new_workbook", sheets=["S"])["workbook"]
    srv.call("excel_save_as", workbook=ok, path=ok_path)
    secret = srv.call("excel_new_workbook", sheets=["Hidden"])["workbook"]
    srv.call("excel_write_range", workbook=secret, sheet="Hidden", cells="A1", values=[["PRIVATE"]])
    srv.call("excel_save_as", workbook=secret, path=secret_path)
    srv.call("excel_select_range", workbook="secret.xlsx", sheet="Hidden", cells="A1")  # секретная книга стала активной
    guarded = Stdio(env={"OFFICE_LIVE_ALLOWED_DIRS": inside})
    try:
        sel = guarded.call("excel_get_selection")
        assert sel["workbook"] is None and "PRIVATE" not in json.dumps(sel, ensure_ascii=False)
        listing = guarded.call("excel_list_workbooks")
        assert listing["active_workbook"] is None and "secret" not in json.dumps(listing, ensure_ascii=False).lower() and "Hidden" not in json.dumps(listing)
    finally:
        guarded.close()
        _close(srv, secret_path, ok_path)


def test_save_and_close_respect_the_write_path_checks(srv, wb, tmp):
    path = os.path.join(tmp, "p.xlsx")
    srv.call("excel_save_as", workbook=wb, path=path)
    srv.call("excel_write_range", workbook="p.xlsx", sheet="Data", cells="A1", values=[[1]])
    guarded = Stdio(env={"OFFICE_LIVE_ALLOWED_DIRS": os.path.join(tmp, "elsewhere")})
    try:
        for tool, args in (("excel_save", {"workbook": "p.xlsx"}), ("excel_close_workbook", {"workbook": "p.xlsx", "save": True})):
            with pytest.raises(ToolFailed, match="not found|no open workbooks|outside the allowed"):
                guarded.call(tool, **args)
    finally:
        guarded.close()
    assert srv.call("excel_workbook_info", workbook="p.xlsx")["saved"] is False  # правка осталась несохранённой, файл не тронут


# ================================================================== адресация листов и замена


def test_replace_treats_wildcards_literally_and_reports_actual_changes(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["a*b"], ["axxb"], ["what?"], ["~tilde"], ["1.0"]], value_mode="text")
    r = srv.call("excel_replace", workbook=wb, sheet="Data", find="*", replace="#")
    assert r["cells_matching"] == 1 and r["cells_changed"] == 1
    r = srv.call("excel_replace", workbook=wb, sheet="Data", find="?", replace="!")
    assert r["cells_changed"] == 1
    r = srv.call("excel_replace", workbook=wb, sheet="Data", find="~", replace="-")
    assert r["cells_changed"] == 1
    assert [row[0] for row in read(srv, wb, "A1:A5")] == ["a#b", "axxb", "what!", "-tilde", "1.0"]
    none = srv.call("excel_replace", workbook=wb, sheet="Data", find="zzz", replace="y")
    assert none["cells_changed"] == 0 and "warning" not in none


def test_hide_rows_uses_the_sheet_from_the_reference(srv, wb):
    srv.call("excel_add_worksheet", workbook=wb, name="Other")
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["S"], ["x"], ["x"]])
    srv.call("excel_write_range", workbook=wb, sheet="Other", cells="A1", values=[["S"], ["x"], ["y"]])
    r = srv.call("excel_hide_rows_columns", workbook=wb, where={"column": "S", "operator": "equals", "value": "x"}, cells="Other!A1:A3")
    assert r["sheet"] == "Other" and r["count"] == 1
    hidden = _probe("result = {n: [bool(wb.Worksheets(n).Rows(i).Hidden) for i in (1, 2, 3)] for n in ('Data', 'Other')}", workbook=wb)
    assert hidden == {"Data": [False, False, False], "Other": [False, True, False]}  # скрыта строка листа Other, а не активного Data
    with pytest.raises(ToolFailed, match="Conflicting sheets"):
        srv.call("excel_hide_rows_columns", workbook=wb, sheet="Data", where={"column": "S", "operator": "equals", "value": "x"}, cells="Other!A1:A3")
    with pytest.raises(ToolFailed, match="Conflicting sheets"):
        srv.call("excel_filter_range", workbook=wb, sheet="Data", cells="Other!A1:A3", filters=[{"column": "S", "criteria": "x"}])
    f = srv.call("excel_filter_range", workbook=wb, sheet="", cells="Other!A1:A3", filters=[{"column": "S", "criteria": "=y"}])
    assert f["sheet"] == "Other" and f["visible_data_rows"] == 1


def test_find_stops_at_the_cap_and_says_so(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["hit"]] * 30)
    capped = srv.call("excel_find", workbook=wb, sheet="Data", query="hit", max_results=5)
    assert capped["returned"] == 5 and capped["truncated"] is True and capped["total_matches"] >= 6
    full = srv.call("excel_find", workbook=wb, sheet="Data", query="hit", max_results=5, count_all=True)
    assert full["total_matches"] == 30 and full["truncated"] is False


def test_narrow_table_header_is_not_mistaken_for_a_title(srv, wb):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["Item", "Qty"], ["a", 1], ["b", 2], ["c", 3]])
    lay = srv.call("excel_describe_layout", workbook=wb, sheet="Data")
    assert lay["title_rows"] == [] and [c["header"] for c in lay["columns"]] == ["Item", "Qty"]
    assert lay["truncated"]["columns"] is False


# ================================================================== сводные


def test_pivot_data_field_edit_keeps_unrelated_settings(srv, wb):
    rows = [["Region", "Revenue"], ["N", 100], ["N", 50], ["S", 70]]
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=rows)
    srv.call("excel_create_pivot_table", workbook=wb, source="A1:B4", source_sheet="Data", rows=["Region"],
             values=[{"field": "Revenue", "function": "sum"}, {"field": "Revenue", "function": "average", "caption": "Avg Revenue"}], dest_sheet="P", name="PT")
    res = srv.call("excel_pivot_fields", workbook=wb, action="data_field", field="Avg Revenue", number_format="0.00")
    assert res["changed"] == ["number_format"]  # функция НЕ сброшена на сумму
    info = srv.call("excel_pivot_info", workbook=wb)
    assert [d["caption"] for d in info["data_fields"]] == ["Sum of Revenue", "Avg Revenue"]
    res = srv.call("excel_pivot_fields", workbook=wb, action="data_field", field="Avg Revenue", function="max", new_caption="Max Revenue")
    assert "function=max" in res["changed"]
    info = srv.call("excel_pivot_info", workbook=wb)
    assert [d["caption"] for d in info["data_fields"]] == ["Sum of Revenue", "Max Revenue"]  # изменено именно выбранное поле, не первое с тем же источником


# ================================================================== Word


def _table_probe(code, document):
    return _probe(code, document=document)


def test_word_table_block_formatting_stays_inside_the_block(srv, doc):
    srv.call("word_create_table", document=doc, data=[["a", "b", "c"], ["d", "e", "f"], ["g", "h", "i"]], style="", header_row=False)
    srv.call("word_format_table", document=doc, table_index=1, row_from=2, row_to=3, col_from=2, col_to=3, bold=True, font_color="#C00000")
    bold = _table_probe("t = doc.Tables(1)\nresult = [[bool(t.Cell(r, c).Range.Font.Bold) for c in (1, 2, 3)] for r in (1, 2, 3)]", doc)
    assert bold == [[False, False, False], [False, True, True], [False, True, True]]  # раньше захватывало и (2,3)->(3,1)


def test_word_table_borders_outline_does_not_create_inside_lines(srv, doc):
    srv.call("word_create_table", document=doc, data=[["a", "b", "c"], ["d", "e", "f"], ["g", "h", "i"]], style="", header_row=False)
    srv.call("word_format_table", document=doc, table_index=1, borders="outline")
    probe = _table_probe("t = doc.Tables(1)\nresult = {'top': t.Borders(-1).LineStyle, 'inside_h': t.Borders(-5).LineStyle, 'inside_v': t.Borders(-6).LineStyle}", doc)
    assert probe["top"] == 1 and probe["inside_h"] == 0 and probe["inside_v"] == 0
    srv.call("word_format_table", document=doc, table_index=1, row_from=2, row_to=3, col_from=2, col_to=3, borders="outline", border_color="#FF0000")
    cells = _table_probe("t = doc.Tables(1)\nresult = {'in_top': t.Cell(2, 2).Borders(-1).LineStyle, 'in_right': t.Cell(3, 3).Borders(-4).LineStyle, 'inner': t.Cell(2, 2).Borders(-4).LineStyle, 'outside_cell_left': t.Cell(3, 1).Borders(-2).LineStyle}", doc)
    assert cells["in_top"] == 1 and cells["in_right"] == 1 and cells["inner"] == 0 and cells["outside_cell_left"] == 1  # (3,1) левая — внешняя граница таблицы


def test_word_write_table_checks_width_before_adding_rows(srv, doc):
    srv.call("word_create_table", document=doc, data=[["a", "b"], ["c", "d"]])
    with pytest.raises(ToolFailed, match="columns wide"):
        srv.call("word_write_table", document=doc, table_index=1, values=[["1", "2", "3"]] * 4, start_row=1)
    assert srv.call("word_read_table", document=doc, table_index=1)["rows"] == 2  # строки не добавлены


def test_word_create_table_bad_arguments_leave_the_document_untouched(srv, doc):
    before = srv.call("word_get_structure", document=doc)
    for bad in ({"alignment": "sideways"}, {"header_fill": "not-a-color"}, {"column_widths_cm": [1, 2, 3, 4]}):
        with pytest.raises(ToolFailed):
            srv.call("word_create_table", document=doc, data=[["a", "b"], ["c", "d"]], **bad)
    assert len(srv.call("word_get_structure", document=doc)["tables"]) == len(before["tables"]) == 0


def test_headers_footers_validate_before_changing_and_respect_section_links(srv, doc):
    srv.call("word_headers_footers", document=doc, action="set", kind="footer", text="KEEP")
    with pytest.raises(ToolFailed, match="page_numbers"):
        srv.call("word_headers_footers", document=doc, action="set", kind="footer", text="LOST?", page_numbers="bogus")
    assert srv.call("word_headers_footers", document=doc)["sections"][0]["footer_primary"] == "KEEP"  # раньше содержимое стиралось до проверки
    srv.call("word_insert_break", document=doc, kind="section_next_page", position="end")
    srv.call("word_insert_text", document=doc, text="second section")
    with pytest.raises(ToolFailed, match="linked to the previous section"):
        srv.call("word_headers_footers", document=doc, action="set", kind="footer", section=2, text="TWO")
    assert srv.call("word_headers_footers", document=doc, section=1)["sections"][0]["footer_primary"] == "KEEP"
    srv.call("word_headers_footers", document=doc, action="set", kind="footer", section=2, text="TWO", break_link=True)
    got = {s["section"]: s.get("footer_primary") for s in srv.call("word_headers_footers", document=doc, section=0)["sections"]}
    assert got == {1: "KEEP", 2: "TWO"}


def test_update_fields_reaches_headers_and_footers(srv, doc):
    srv.call("word_insert_text", document=doc, text="Body")
    srv.call("word_headers_footers", document=doc, action="set", kind="footer", text="p.", page_numbers="page")
    r = srv.call("word_update_fields", document=doc)
    assert r["ok"] is True and r["fields_updated"] >= 1 and "9" in r["fields_by_story"]  # 9 = wdPrimaryFooterStory


def test_mail_merge_preflight_unique_names_and_no_silent_overwrite(srv, wb, tmp):
    tpl = srv.call("word_new_document", text="Dear {{Name}}")["document"]
    tpl_path = os.path.join(tmp, "tpl.docx")
    srv.call("word_save_as", document=tpl, path=tpl_path)
    srv.call("word_close_document", document="tpl.docx", discard=True)
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["Name"], ["Ivanov"], ["Ivanov"], [""]], value_mode="text")
    out = os.path.join(tmp, "out")
    r = srv.call("bridge_excel_to_word_documents", workbook=wb, sheet="Data", cells="A1:A4", template=tpl_path, output_dir=out, filename_column="Name", export_pdf=True)
    assert r["documents_created"] == 2 and r["renamed_for_uniqueness"][0]["used"].startswith("Ivanov_") and r["skipped_blank_rows"] == 1
    produced = sorted(os.listdir(out))
    assert len(produced) == 4 and produced.count("Ivanov.docx") == 1 and sum(f.endswith(".pdf") for f in produced) == 2
    with pytest.raises(ToolFailed, match="already exist"):
        srv.call("bridge_excel_to_word_documents", workbook=wb, sheet="Data", cells="A1:A4", template=tpl_path, output_dir=out, filename_column="Name", export_pdf=True)
    assert sorted(os.listdir(out)) == produced  # отказ — до создания чего-либо
    os.remove(os.path.join(out, "Ivanov.pdf"))  # свободен .docx, но занят другой .pdf -> всё равно отказ (раньше PDF затирался молча)
    with pytest.raises(ToolFailed, match="already exist"):
        srv.call("bridge_excel_to_word_documents", workbook=wb, sheet="Data", cells="A1:A4", template=tpl_path, output_dir=out, filename_column="Name", export_pdf=True)


# ================================================================== прочее


def test_export_chart_image_refuses_to_replace_a_file(srv, wb, tmp):
    srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["x", "y"], [1, 2], [2, 3]])
    srv.call("excel_create_chart", workbook=wb, sheet="Data", source="A1:B3", name="Chart1")
    dest = os.path.join(tmp, "chart.png")
    srv.raw_call("excel_manage_charts", {"workbook": wb, "action": "export_image", "chart_name": "Chart1", "export_path": dest})
    assert os.path.getsize(dest) > 500
    stamp = os.path.getmtime(dest)
    res = srv.raw_call("excel_manage_charts", {"workbook": wb, "action": "export_image", "chart_name": "Chart1", "export_path": dest})
    assert res.get("isError") and "already exists" in res["content"][0]["text"] and os.path.getmtime(dest) == stamp
    ok = srv.raw_call("excel_manage_charts", {"workbook": wb, "action": "export_image", "chart_name": "Chart1", "export_path": dest, "overwrite": True})
    assert not ok.get("isError")


def test_strict_target_needs_exact_names_and_sheet(srv, wb):
    strict = Stdio(env={"OFFICE_LIVE_STRICT_TARGET": "1"})
    try:
        with pytest.raises(ToolFailed, match="name is required"):
            strict.call("excel_write_range", workbook="   ", sheet="Data", cells="A1", values=[[1]])
        with pytest.raises(ToolFailed, match="not found"):
            strict.call("excel_write_range", workbook=wb[:3], sheet="Data", cells="A1", values=[[1]])  # подстрока не годится
        with pytest.raises(ToolFailed, match="Sheet name is required"):
            strict.call("excel_write_range", workbook=wb, sheet="", cells="A1", values=[[1]])
        strict.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[[1]])
    finally:
        strict.close()
