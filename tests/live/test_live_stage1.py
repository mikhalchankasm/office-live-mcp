"""Этап 1 на собственных временных файлах. Только написано; требуется OFFICE_LIVE_LIVE_TESTS=1."""

import contextlib
from pathlib import Path
import uuid

import pytest

from tests.live.conftest import RUN_TAG, close_own_documents, close_own_workbooks, read
from tests.live.mcpclient import ToolFailed

pytestmark = pytest.mark.live


@pytest.fixture
def stage_book(srv, wb, tmp):
    path = str(Path(tmp) / "stage1.xlsx")
    srv.call("excel_save_as", workbook=wb, path=path)
    return path


@pytest.fixture
def stage_doc(srv, doc, tmp):
    path = str(Path(tmp) / "stage1.docx")
    srv.call("word_save_as", document=doc, path=path)
    return path


def in_excel(path, fn):
    from office_live import com, xl_common

    assert RUN_TAG in path

    def invoke():
        app, book = xl_common.pick_workbook(path)
        return fn(app, book.Worksheets("Data"))

    return com.run_com(invoke)


def in_word(path, fn):
    from office_live import com, wd_common

    assert RUN_TAG in path

    def invoke():
        _, doc = wd_common.pick_document(path)
        return fn(doc)

    return com.run_com(invoke)


def word_state(path):
    from office_live.undo import word_fingerprint

    return in_word(path, lambda doc: (word_fingerprint(doc), bool(doc.Saved), int(doc.Revisions.Count)))


@contextlib.contextmanager
def extra_file(srv, tmp, kind):
    marker = f"{RUN_TAG}_{uuid.uuid4().hex[:8]}"
    if kind == "excel":
        name = srv.call("excel_new_workbook", sheets=["Data"])["workbook"]
        srv.call("excel_manage_names", workbook=name, action="add", name=marker, formula="=1")
        cleanup = close_own_workbooks
        path = str(Path(tmp) / (marker + ".xlsx"))
        save_args = {"workbook": name}
    else:
        name = srv.call("word_new_document", text="Revised source.")["document"]
        srv.call("word_document_properties", document=name, action="set", properties={"Comments": marker})
        cleanup = close_own_documents
        path = str(Path(tmp) / (marker + ".docx"))
        save_args = {"document": name}
    try:
        srv.call(kind + "_save_as", path=path, **save_args)
        yield path
    finally:
        cleanup(srv, name, marker)


@contextlib.contextmanager
def compared_result(srv, **args):
    result = srv.call("word_compare_documents", **args)
    name = result["document"]
    marker = f"{RUN_TAG}_{uuid.uuid4().hex[:8]}"
    srv.call("word_document_properties", document=name, action="set", properties={"Comments": marker})
    try:
        yield result
    finally:
        close_own_documents(srv, name, marker)


@pytest.mark.parametrize("mode", ["values", "formulas", "formats"])
def test_live_compare_ranges_modes(srv, stage_book, mode):
    book = stage_book
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=[[1, 1], ["text", "text"]])
    if mode == "formulas":
        srv.call("excel_set_formula", workbook=book, sheet="Data", cells="A3:B3", formula=[["=A1+1", "=B1+1"]])
    elif mode == "formats":
        srv.call("excel_format_range", workbook=book, sheet="Data", cells="B2", bold=True)
    result = srv.call("excel_compare_ranges", workbook_a=book, sheet_a="Data", cells_a="A1:A3", cells_b="B1:B3", compare=mode)
    assert result["summary"]["compared_cells"] == 3
    assert len(result["differences"]) == (0 if mode == "values" else 1)


def test_live_compare_ranges_key_two_workbooks(srv, stage_book, tmp):
    with extra_file(srv, tmp, "excel") as other:
        srv.call("excel_write_range", workbook=stage_book, sheet="Data", cells="A1", values=[["ID", "Value"], [1, "a"], [2, "b"]])
        srv.call("excel_write_range", workbook=other, sheet="Data", cells="A1", values=[["ID", "Value"], [3, "new"], [2, "changed"]])
        result = srv.call("excel_compare_ranges", workbook_a=stage_book, workbook_b=other, sheet_a="Data", cells_a="A:B", match="key", key_column="ID")
        assert result["summary"]["rows_only_in_a"] == result["summary"]["rows_only_in_b"] == 1
        assert result["differences"][0]["key"] == 2


def test_live_clean_text_preview_types_formats_and_undo(srv, stage_book):
    book = stage_book
    values = [[" 00123 "], [" =x "], [" 12,5 "], [" a\u00a0  b\nline "], ["1,234"], ["+79991234567"]]
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=values, value_mode="text")
    srv.call("excel_format_range", workbook=book, sheet="Data", cells="A3", number_format="text")
    srv.call("excel_set_formula", workbook=book, sheet="Data", cells="A7", formula="=1+1")
    before = read(srv, book, "A1:A7")
    fmt = srv.call("excel_get_format", workbook=book, sheet="Data", cells="A3")
    args = dict(workbook=book, sheet="Data", cells="A1:A7", operations=["trim", "clean", "text_to_number"])
    history = srv.call("office_undo", file=book, action="history")
    preview = srv.call("excel_clean_text", **args)
    assert preview["preview"] and preview["skipped_formulas"] == 1
    assert read(srv, book, "A1:A7") == before and srv.call("office_undo", file=book, action="history") == history
    result = srv.call("excel_clean_text", **args, preview=False)
    assert result["undo"] == "available" and result["format_changed"] == ["A3"]
    assert read(srv, book, "A1:A3") == [["00123"], ["=x"], [12.5]]
    assert in_excel(book, lambda app, ws: bool(ws.Range("A2").HasFormula)) is False
    srv.call("office_undo", file=book)
    assert read(srv, book, "A1:A7") == before
    assert srv.call("excel_get_format", workbook=book, sheet="Data", cells="A3") == fmt


def test_live_clean_text_merge_and_conflict(srv, stage_book):
    book = stage_book
    srv.call("excel_merge_cells", workbook=book, sheet="Data", cells="A1:B1")
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=" x ", value_mode="text")
    result = srv.call("excel_clean_text", workbook=book, sheet="Data", cells="A1:B2", operations=["trim"], preview=False)
    assert result["skipped_merged"] == 1 and read(srv, book, "A1") == [["x"]]
    in_excel(book, lambda app, ws: setattr(ws.Range("B2"), "Value2", "user edit"))
    with pytest.raises(ToolFailed, match="later edits"):
        srv.call("office_undo", file=book)
    srv.call("office_undo", file=book, force=True)
    assert read(srv, book, "A1") == [[" x "]]


def test_live_clean_text_protected_sheet_refusal(srv, stage_book):
    book = stage_book
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=" x ", value_mode="text")
    in_excel(book, lambda app, ws: ws.Protect())
    try:
        with pytest.raises(ToolFailed, match="protected"):
            srv.call("excel_clean_text", workbook=book, sheet="Data", cells="A1", operations=["trim"], preview=False)
        assert read(srv, book, "A1") == [[" x "]]
    finally:
        in_excel(book, lambda app, ws: ws.Unprotect())


def test_live_clean_text_pivot_refusal(srv, stage_book):
    book = stage_book
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=[["Group", "Amount"], [" a ", 1], [" b ", 2]])
    srv.call("excel_create_pivot_table", workbook=book, source="A1:B3", source_sheet="Data", dest_sheet="Data", dest_cell="D1",
             name="Stage1Pivot", rows=["Group"], values=[{"field": "Amount", "function": "sum"}])
    before = read(srv, book, "D1:E10")
    with pytest.raises(ToolFailed, match="Stage1Pivot"):
        srv.call("excel_clean_text", workbook=book, sheet="Data", cells="D1:E10", operations=["trim"], preview=False)
    assert read(srv, book, "D1:E10") == before


@pytest.mark.parametrize("closed", [False, True])
def test_live_compare_documents_sources_saved_and_temporary_cleanup(srv, stage_doc, tmp, closed):
    srv.call("word_insert_text", document=stage_doc, text="Original source.")
    with extra_file(srv, tmp, "word") as other:
        before = word_state(stage_doc)
        other_before = word_state(other)
        if closed:
            srv.call("word_close_document", document=other, discard=True)
        with compared_result(srv, original=stage_doc, revised=other) as result:
            assert result["revision_count"] > 0 and result["saved"] is False
            assert word_state(stage_doc) == before
            if not closed:
                assert word_state(other) == other_before
            else:
                assert not any(d["name"] == Path(other).name for d in srv.call("word_list_documents")["documents"])


def test_live_compare_documents_existing_revisions_warning(srv, stage_doc, tmp):
    with extra_file(srv, tmp, "word") as other:
        srv.call("word_track_changes", document=stage_doc, action="enable")
        srv.call("word_insert_text", document=stage_doc, text="Tracked source.")
        before = word_state(stage_doc)
        with compared_result(srv, original=stage_doc, revised=other) as result:
            assert result["warnings"] and word_state(stage_doc) == before


@pytest.mark.parametrize("header,desc", [(True, False), (True, True), (False, False)])
def test_live_sort_table_three_keys_and_undo(srv, stage_doc, header, desc):
    doc = stage_doc
    body = [["b", "1 234,5", "05.10.2026"], ["a", "12,5", "01.01.2025"], ["a", "12,5", "01.01.2024"]]
    data = ([["Name", "Number", "Date"]] if header else []) + body
    index = srv.call("word_create_table", document=doc, data=data, header_row=header)["table"]
    in_word(doc, lambda d: setattr(d.Tables(index).Range, "LanguageID", 1049))
    before = srv.call("word_read_table", document=doc, table_index=index)["values"]
    result = srv.call("word_sort_table", document=doc, table_index=index, header=header,
                      keys=[{"column": 1, "order": "desc" if desc else "asc"}, {"column": 2, "type": "number"}, {"column": 3, "type": "date"}])
    expected = [body[0], body[2], body[1]] if desc else [body[2], body[1], body[0]]
    assert result["values"][int(header):] == expected and result["sorted_rows"] == 3
    srv.call("office_undo", file=doc)
    assert srv.call("word_read_table", document=doc, table_index=index)["values"] == before


@pytest.mark.parametrize("problem", ["merged", "nested", "tracking", "number", "date", "column"])
def test_live_sort_table_refusals_no_edits(srv, stage_doc, problem):
    doc = stage_doc
    index = srv.call("word_create_table", document=doc, data=[["H", "V"], ["b", "bad"], ["a", "also bad"]])["table"]
    keys = [{"column": 1}]
    if problem == "merged":
        srv.call("word_modify_table", document=doc, table_index=index, action="merge_cells", row=2, col=1, to_row=2, to_col=2)
    elif problem == "nested":
        def nest(d):
            rng = d.Tables(index).Cell(2, 1).Range
            rng.Collapse(1)
            d.Tables.Add(rng, 1, 1)
        in_word(doc, nest)
    elif problem == "tracking":
        srv.call("word_track_changes", document=doc, action="enable")
    elif problem in {"number", "date"}:
        keys = [{"column": 2, "type": problem}]
    else:
        keys = [{"column": 3}]
    before = word_state(doc)
    with pytest.raises(ToolFailed):
        srv.call("word_sort_table", document=doc, table_index=index, keys=keys)
    assert word_state(doc) == before


def test_live_sort_table_user_edit_conflict_and_force(srv, stage_doc):
    doc = stage_doc
    index = srv.call("word_create_table", document=doc, data=[["H"], ["b"], ["a"]])["table"]
    before = srv.call("word_read_table", document=doc, table_index=index)["values"]
    srv.call("word_sort_table", document=doc, table_index=index, keys=[{"column": 1}])
    in_word(doc, lambda d: setattr(d.Tables(index).Cell(2, 1).Range, "Text", "user edit"))
    with pytest.raises(ToolFailed, match="later edits"):
        srv.call("office_undo", file=doc)
    srv.call("office_undo", file=doc, force=True)
    assert srv.call("word_read_table", document=doc, table_index=index)["values"] == before
