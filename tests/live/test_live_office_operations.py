"""Office operations Office 16/Russian-locale probes. Written only; NEVER run without the owner's approval.

Every COM target belongs to this run's marked temporary directory. No user document is opened or closed.
"""

from pathlib import Path

import pytest

from tests.live.conftest import read
from tests.live.mcpclient import Stdio, ToolFailed
from tests.live.test_live_stage1 import extra_file, in_excel, in_word, word_state

pytestmark = pytest.mark.live


@pytest.fixture
def srv(tmp):
    client = Stdio(env={"OFFICE_LIVE_MODE": "full", "OFFICE_LIVE_JOURNAL_DIR": str(Path(tmp) / "journal"),
                        "OFFICE_LIVE_AUDIT_LOG": str(Path(tmp) / "audit.jsonl"), "OFFICE_LIVE_ALLOWED_DIRS": tmp})
    yield client
    client.close()


@pytest.fixture
def operations_book(srv, wb, tmp):
    path = str(Path(tmp) / "operations.xlsx")
    srv.call("excel_save_as", workbook=wb, path=path)
    return path


@pytest.fixture
def operations_doc(srv, doc, tmp):
    path = str(Path(tmp) / "operations.docx")
    srv.call("word_save_as", document=doc, path=path)
    return path


def goal(srv, book, **kw):
    return srv.call("excel_goal_seek", **{"workbook": book, "sheet": "Data", "set_cell": "B1", "changing_cell": "A1", "target_value": 100, **kw})


def test_probe_goal_seek_bool_no_dialog_manual_settings_and_undo(srv, operations_book):
    book = operations_book
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=[[2]])
    srv.call("excel_set_formula", workbook=book, sheet="Data", cells="B1", formula="=A1*2")
    def setup(app, ws):
        other = ws.Parent.Worksheets.Add(None, ws)
        other.Name = "InactiveTargetProbe"
        other.Activate()
        old = (app.MaxIterations, app.MaxChange, app.Calculation)
        app.Calculation = -4135
        return old
    before = in_excel(book, setup)
    try:
        result = goal(srv, book, max_iterations=50, max_change=0.00001)
        assert result["converged"] is True and result["manual_calculation"]
        assert result["dependency_verified"]
        assert in_excel(book, lambda app, ws: app.ActiveSheet.Name) == "InactiveTargetProbe"
        assert result["new_value"] == pytest.approx(50)
        assert result["achieved_value"] == pytest.approx(100)
        assert in_excel(book, lambda app, ws: (app.MaxIterations, app.MaxChange)) == before[:2]
        srv.call("office_undo", file=book)
        assert read(srv, book, "A1")[0][0] == 2
    finally:
        in_excel(book, lambda app, ws: setattr(app, "Calculation", before[2]))


def test_probe_goal_seek_nonconvergence_restores_or_keeps_attempt(srv, operations_book):
    book = operations_book
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=[[2]])
    srv.call("excel_set_formula", workbook=book, sheet="Data", cells="B1", formula="=A1^2")
    in_excel(book, lambda app, ws: ws.Activate())
    result = goal(srv, book, target_value=-1, max_iterations=5)
    assert not result["converged"] and not result["kept"] and read(srv, book, "A1")[0][0] == 2
    assert "undo" not in result
    result = goal(srv, book, target_value=-1, max_iterations=5, keep_if_not_converged=True)
    assert not result["converged"] and result["kept"] and result["undo"] == "available"
    srv.call("office_undo", file=book)
    assert read(srv, book, "A1")[0][0] == 2


def test_probe_goal_seek_cross_sheet_api_is_observational_only(srv, operations_book):
    """Record cross-sheet COM behavior; the public tool continues to refuse it regardless of this probe."""
    def probe(app, ws):
        other = ws.Parent.Worksheets.Add(None, ws)
        other.Name = "Другой"
        ws.Range("A1").Value2 = 2
        other.Range("B1").Formula = "=Data!A1*2"
        other.Activate()
        try:
            try:
                result = other.Range("B1").GoalSeek(100, ws.Range("A1"))
                assert isinstance(result, bool)
                return {"accepted": True, "converged": result, "value": ws.Range("A1").Value2}
            except Exception as exc:
                return {"accepted": False, "error": str(exc)}
        finally:
            ws.Range("A1").Value2 = 2
    observed = in_excel(operations_book, probe)
    print("Cross-sheet GoalSeek:", observed)
    with pytest.raises(ToolFailed):
        goal(srv, operations_book, set_cell="Другой!B1", changing_cell="Data!A1")


@pytest.mark.parametrize("sort,below", [(False, True), (True, True), (True, False)])
def test_probe_subtotal_row_anchor_and_undo_preserve_below_right(srv, operations_book, sort, below):
    book = operations_book
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=[["Key", "Value"], ["b", 2], ["a", 3], ["b", 4]])
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="D2", values=[["right"]])
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A10", values=[["below", "data", None, "below/right"]])
    srv.call("excel_format_range", workbook=book, sheet="Data", cells="B3", number_format="0.00")
    before = in_excel(book, lambda app, ws: ws.Range("A1:F15").Formula)
    # Late-bound NumberFormat is read in the user's locale ('0,00' in Russian Excel): compare before/after the same way.
    format_before = in_excel(book, lambda app, ws: ws.Range("B3").NumberFormat)
    result = srv.call("excel_subtotals", workbook=book, sheet="Data", cells="A1:B4", group_by="Key", columns=["Value"], sort_first=sort, summary_below=below)
    assert result["groups"] == (2 if sort else 3)
    assert result["inserted_rows"] == result["groups"] + 1
    assert result["range"] == f"A1:B{4 + result['inserted_rows']}"
    srv.call("office_undo", file=book)
    assert in_excel(book, lambda app, ws: ws.Range("A1:F15").Formula) == before
    assert in_excel(book, lambda app, ws: ws.Range("B3").NumberFormat) == format_before
    assert in_excel(book, lambda app, ws: [int(ws.Rows(r).OutlineLevel) for r in range(1, 5)]) == [1] * 4


@pytest.mark.parametrize("location", ["N2:N3", "B5:D5"])
def test_probe_sparkline_quoted_russian_source_orientation_and_add_undo(srv, operations_book, location):
    book = operations_book
    def setup(app, ws):
        source = ws.Parent.Worksheets.Add(None, ws)
        source.Name = "Данные 'тест' 1"  # Excel forbids an apostrophe at the start or end of a sheet name
        source.Range("B2:D3").Value2 = ((1, 2, 3), (3, 2, 1))
        ws.Range(location).Value2 = 7
    in_excel(book, setup)
    result = srv.call("excel_sparklines", workbook=book, sheet="Data", action="add", data="'Данные ''тест'' 1'!B2:D3",
                      location=location, type="line", color="#CC0000", markers=["high", "negative", "all"])
    assert result["cells_with_content"] == (2 if location.startswith("N") else 3)
    listed = srv.call("excel_sparklines", workbook=book, sheet="Data", action="list")["groups"]
    assert len(listed) == 1 and listed[0]["location"] == location and "Данные" in listed[0]["source"]
    srv.call("office_undo", file=book)
    assert srv.call("excel_sparklines", workbook=book, sheet="Data", action="list")["groups"] == []
    assert all(v == 7 for row in read(srv, book, location) for v in row)


def test_probe_sparkline_clear_properties_theme_date_axis_roundtrip(srv, operations_book):
    from office_live.excel_format import spark_state
    book = operations_book
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=[[1, 2, 3], [3, 2, 1]])
    srv.call("excel_sparklines", workbook=book, sheet="Data", action="add", data="A1:C2", location="D1:D2", markers=["low"])
    def setup(app, ws):
        ws.Range("E1:G1").Value2 = ((44927, 44928, 44929),)
        g = ws.Range("D1:D2").SparklineGroups.Item(1)
        g.DateRange = "Data!E1:G1"
        g.Axes.Horizontal.Axis.Visible = True
        g.SeriesColor.ThemeColor = 4  # SeriesColor is a FormatColor itself (live probe)
        g.SeriesColor.TintAndShade = 0.3
        g.LineWeight = 2
        g.DisplayHidden = True
        return spark_state(g, full=True)
    before = in_excel(book, setup)
    result = srv.call("excel_sparklines", workbook=book, sheet="Data", action="clear", location="D1:D2")
    assert result["undo"] == "available", result
    srv.call("office_undo", file=book)
    after = in_excel(book, lambda app, ws: spark_state(ws.Range("D1:D2").SparklineGroups.Item(1), full=True))
    assert after == before


def test_probe_word_password_barrier_wrong_password_and_passwordless_reprotect(srv, operations_doc, tmp):
    target, secret = operations_doc, "Stage3-temporary-password"
    password_active = True
    try:
        result = srv.call("word_restrict_editing", document=target, action="protect", password=secret)
        assert "Password operation" in result["undo"]
        with pytest.raises(ToolFailed) as failure:
            srv.call("word_restrict_editing", document=target, action="unprotect", password="wrong-operations")
        assert "wrong-operations" not in str(failure.value) and "Неверный пароль" in str(failure.value)
        result = srv.call("word_restrict_editing", document=target, action="unprotect", password=secret)
        password_active = False
        assert "Password operation" in result["undo"]
        assert secret not in (Path(tmp) / "audit.jsonl").read_text(encoding="utf-8")
        srv.call("word_restrict_editing", document=target, action="protect", mode="comments")
        srv.call("word_restrict_editing", document=target, action="unprotect")
        srv.call("office_undo", file=target)
        assert srv.call("word_restrict_editing", document=target)["mode"] == "comments"
    finally:
        if in_word(target, lambda doc: int(doc.ProtectionType)) != -1:
            in_word(target, lambda doc: doc.Unprotect(secret if password_active else ""))


def test_probe_insert_file_bookmark_saved_disk_version_selection_one_undo(srv, operations_doc, tmp):
    target = operations_doc
    with extra_file(srv, tmp, "word") as source:
        in_word(source, lambda doc: doc.Bookmarks.Add("piece", doc.Range(0, 7)))
        srv.call("word_save", document=source)
        srv.call("word_insert_text", document=source, text="unsaved tail")
        selection = in_word(target, lambda doc: (doc.Application.ActiveWindow.Hwnd, doc.Application.Selection.Start, doc.Application.Selection.End))
        before = word_state(target)[0]
        result = srv.call("word_insert_document", document=target, file_path=source, source_bookmark="piece", position="start")
        assert result["characters_inserted"] > 0 and any("на диске" in warning for warning in result["warnings"])
        assert in_word(target, lambda doc: (doc.Application.ActiveWindow.Hwnd, doc.Application.Selection.Start, doc.Application.Selection.End)) == selection
        undone = srv.call("office_undo", file=target)
        assert undone["native_steps"] == 1 and word_state(target)[0] == before


@pytest.mark.parametrize("mode", ["read_only", "comments", "tracked_changes", "forms"])
def test_probe_word_protection_inverse_guard_comments_tracking(srv, operations_doc, mode):
    target = operations_doc
    try:
        result = srv.call("word_restrict_editing", document=target, action="protect", mode=mode)
        assert result["mode"] == mode and result["undo"] == "available"
        if mode == "tracked_changes":
            changed = srv.call("word_insert_text", document=target, text="tracked")
            assert any("tracked_changes" in warning for warning in changed["warnings"])
            srv.call("office_undo", file=target)
        else:
            with pytest.raises(ToolFailed):
                srv.call("word_insert_text", document=target, text="refused")
        if mode == "comments":
            srv.call("word_manage_comments", document=target, action="add", paragraph=1, text="allowed comment")
            srv.call("office_undo", file=target)
        srv.call("office_undo", file=target)
        assert srv.call("word_restrict_editing", document=target)["mode"] == "none"
    finally:
        if in_word(target, lambda doc: int(doc.ProtectionType)) != -1:
            srv.call("word_restrict_editing", document=target, action="unprotect")


def test_probe_word_undo_record_does_not_undo_protection(srv, operations_doc):
    def probe(doc):
        record = doc.Application.UndoRecord
        try:
            record.StartCustomRecord("Office operations protection probe")
            try:
                doc.Protect(3, True, "", False, False)
            finally:
                record.EndCustomRecord()
            doc.Undo(1)
            assert int(doc.ProtectionType) == 3
        finally:
            if int(doc.ProtectionType) != -1:
                doc.Unprotect("")
    in_word(operations_doc, probe)


@pytest.mark.parametrize("separator", ["tab", "comma", "semicolon", "|"])
def test_probe_word_table_text_positional_separator_and_one_undo(srv, operations_doc, separator):
    target = operations_doc
    literal = {"tab": "\t", "comma": ",", "semicolon": ";"}.get(separator, separator)
    srv.call("word_insert_text", document=target, text=f"A{literal}B\n1{literal}2\n3", position="start")
    before = word_state(target)[0]
    result = srv.call("word_table_text", document=target, action="text_to_table", start_paragraph=1, end_paragraph=3, separator=separator)
    assert result["columns"] == 2 and result["uneven_rows"] == 1
    assert in_word(target, lambda doc: bool(doc.Tables(result["table"]).Rows(1).HeadingFormat))
    undone = srv.call("office_undo", file=target)
    # Live probe: Word splits ConvertToTable + HeadingFormat into two native steps; office_undo collects both.
    assert word_state(target)[0] == before and 1 <= undone["native_steps"] <= 3 and not undone.get("warnings")
    srv.call("word_table_text", document=target, action="text_to_table", start_paragraph=1, end_paragraph=3, separator=separator)
    before_table = word_state(target)[0]
    srv.call("word_table_text", document=target, action="table_to_text", table=1, separator=separator)
    undone = srv.call("office_undo", file=target)
    assert word_state(target)[0] == before_table and 1 <= undone["native_steps"] <= 3 and not undone.get("warnings")


def test_probe_word_table_to_paragraphs_and_merged_refusal(srv, operations_doc):
    target = operations_doc
    srv.call("word_create_table", document=target, data=[["a", "b"], ["c", "d"]])
    before = word_state(target)[0]
    result = srv.call("word_table_text", document=target, action="table_to_text", separator="paragraph")
    assert result["end_paragraph"] >= result["start_paragraph"] + 3
    srv.call("office_undo", file=target)
    assert word_state(target)[0] == before
    in_word(target, lambda doc: doc.Tables(1).Cell(1, 1).Merge(doc.Tables(1).Cell(1, 2)))
    before = word_state(target)[0]
    with pytest.raises(ToolFailed):
        srv.call("word_table_text", document=target, action="table_to_text")
    assert word_state(target)[0] == before


def test_probe_word_nested_table_refused_without_change(srv, operations_doc):
    target = operations_doc
    srv.call("word_create_table", document=target, data=[["a", "b"], ["c", "d"]])
    def nest(doc):
        rng = doc.Tables(1).Cell(2, 1).Range
        rng.End = int(rng.End) - 1
        doc.Tables.Add(rng, 1, 1)
    in_word(target, nest)
    before = word_state(target)[0]
    with pytest.raises(ToolFailed):
        srv.call("word_table_text", document=target, action="table_to_text")
    assert word_state(target)[0] == before


def test_probe_insert_plain_text_conversion_no_dialog(srv, operations_doc, tmp):
    path = Path(tmp) / "unicode.txt"
    path.write_text("Текст без диалога преобразования.\n", encoding="utf-8")
    result = srv.call("word_insert_document", document=operations_doc, file_path=str(path))
    assert result["characters_inserted"] > 0
    assert srv.call("office_undo", file=operations_doc)["native_steps"] == 1


@pytest.mark.parametrize("formula,verified", [("=C1*2", True), ("=InputValue*2", True), ('=INDIRECT("A1")*2', False)])
def test_probe_goal_seek_indirect_named_and_dynamic_dependency(srv, operations_book, formula, verified):
    book = operations_book
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=[[2]])
    srv.call("excel_set_formula", workbook=book, sheet="Data", cells="C1", formula="=A1")
    srv.call("excel_manage_names", workbook=book, action="add", name="InputValue", formula="=Data!$A$1")
    srv.call("excel_set_formula", workbook=book, sheet="Data", cells="B1", formula=formula)
    result = goal(srv, book)
    assert result["converged"] and result["dependency_verified"] is verified
    srv.call("office_undo", file=book)
    assert read(srv, book, "A1")[0][0] == 2
