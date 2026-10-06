import copy
from types import SimpleNamespace as NS
from urllib.parse import quote

import pytest

from office_live import com, config, journal, links, navigation, protocol, registry, undo
from office_live.errors import ToolError
from tests.history_fakes import Collection, Excel, fake_office  # noqa: F401


@pytest.mark.parametrize("app,params", [
    ("excel", {"book": r"C:\Проект & отчёт\Оценка.xlsx", "sheet": "Лист '1'!", "range": "$B$2:$C$5"}),
    ("excel", {"book": "Book1"}), ("excel", {"book": "Book1", "sheet": "Data"}),
    ("excel", {"book": "Book1", "sheet": "Data", "range": "A:C"}),
    ("word", {"doc": "Document1", "paragraph": "2"}),
    ("word", {"doc": "Document1", "paragraphs": "1-1000000"}),
    ("word", {"doc": "Document1", "table": "1"}), ("word", {"doc": "Document1", "bookmark": "Отчёт_1"}),
    ("word", {"doc": "Document1"}), ("word", {"doc": "Document1", "span": "0-20"}),
])
def test_uri_roundtrip(app, params):
    uri = links.build(app, **params)
    assert uri.isascii() and len(uri) <= 2048
    assert links.parse(uri) == (app, params)


BAD = [
    "javascript:alert(1)", "https://excel/open?book=Book1", "file:///C:/test.xlsx",
    "OFFICELIVE://excel/open?book=Book1", "officelive://excel:80/open?book=Book1",
    "officelive://user@excel/open?book=Book1", "officelive://excel/../open?book=Book1",
    "officelive://excel/open/?book=Book1", "officelive://excel/open?book=Book1#fragment",
    "officelive://excel/open?book=Book1&book=Book2", "officelive://excel/open?book=Book1&exec=calc",
    "officelive://excel/open?%62ook=Book1", "officelive://excel/open?book=Book1&doc=Draft",
    "officelive://excel/open?book=Book1&", "officelive://excel/open?book=Book1&sheet=",
    "officelive://excel/open?book=Book1&range=A1", "officelive://excel/open?book=Book+1",
    "officelive://excel/open?book=Book%FF", "officelive://excel/open?book=Book%", "officelive://excel/open?book=Book%0G",
    "officelive://excel/open?book=" + "x" * 2048,
    *["officelive://excel/open?book=" + quote(v, safe="") for v in (
        r"..\book.xlsx", r"C:\a\..\b.xlsx", "..", r"\\host\share\book.xlsx", "//host/share/book.xlsx",
        r"\\?\C:\book.xlsx", r"C:book.xlsx", "/a/book.xlsx", "https://host/book.xlsx", "javascript:calc",
        r"C:\a.xlsx:stream", 'Book" -m os', "Book\x00", "Book\n", "Book\t", "Book\x7f", "Book\u202e1",
        r"%TEMP%\book.xlsx", "%2e%2e", r"C:\a.\b.xlsx", " book.xlsx", "book.xlsx ",
    )],
    *["officelive://excel/open?book=Book1&sheet=Data&range=" + quote(v, safe="") for v in (
        "A0", "XFE1", "A1048577", "A1,B2", "Data!A1", "[other.xlsx]Data!A1", "=cmd()", "NamedRange",
        "A1:B0", "$$A1", "A1$$", "A2:A1", "B1:A2", "1:0", "A", "A01", "A١", "1", "A1 ",
    )],
    *["officelive://word/open?doc=Draft&" + s for s in (
        "paragraph=0", "paragraph=-1", "paragraph=01", "paragraph=1000001", "paragraph=1.0", "paragraph=1e3",
        "paragraphs=3-2", "paragraphs=1-2-3", "table=0", "paragraph=1&table=1", "bookmark=a%20b",
        "bookmark=" + "a" * 41, "bookmark=1bad", "paragraph=1&paragraph=2",
    )],
]


@pytest.mark.parametrize("uri", BAD)
def test_malicious_uri_rejected_before_com(uri, monkeypatch):
    monkeypatch.setattr(com, "run_com", lambda *a, **k: pytest.fail("COM must not be entered"))
    monkeypatch.setenv("OFFICE_LIVE_LINK_NO_DIALOG", "1")
    with pytest.raises(ToolError):
        links.parse(uri)
    assert protocol.open_link([uri]) == 1


def test_handler_errors_use_short_dialog_without_traceback(monkeypatch, capsys):
    messages = []
    monkeypatch.setattr(protocol, "error_dialog", messages.append)
    assert protocol.open_link(["bad"]) == 1
    assert 0 < len(messages[0]) < 400
    assert "Traceback" not in capsys.readouterr().err


@pytest.fixture
def navigate(office, monkeypatch):
    raised = []
    monkeypatch.setenv("OFFICE_LIVE_LINK_NO_DIALOG", "1")
    monkeypatch.setattr(protocol, "foreground", raised.append)
    return lambda uri: protocol.open_link([uri]), raised


def test_office_link_is_read_and_does_not_select(office):
    out = office.call("office_link", workbook=office.wb.Name, sheet="Data", cells="B2:C4")
    assert links.parse(out["links"][0]["uri"])[1]["range"] == "B2:C4"
    assert out["links_truncated"] is False and registry.CATALOG["office_link"].kind == "read"
    assert office.excel.Selection is None and not undo.STACKS
    assert not list((office.tmp / "journal").glob("*"))


def test_excel_link_navigates_without_content_history_or_journal_changes(office, navigate):
    run, raised = navigate
    office.ws.Range("B2:C4").Value = 42
    before = copy.deepcopy((office.ws.data, office.ws.formats, office.wb.Saved))
    uri = office.call("office_link", workbook=office.wb.Name, sheet="Data", cells="B2:C4")["links"][0]["uri"]
    assert run(uri) == 0
    assert office.excel.Selection.Address == "B2:C4"
    assert (office.excel.ActiveWindow.ScrollRow, office.excel.ActiveWindow.ScrollColumn) == (2, 2)
    assert raised == [42] and office.excel.EnableEvents
    assert (office.ws.data, office.ws.formats, office.wb.Saved) == before
    assert not undo.STACKS and not list((office.tmp / "journal").glob("*"))


@pytest.mark.parametrize("params", [{"book": "Book"}, {"book": "Book1", "sheet": "Dat"},
                                      {"book": "Book1", "sheet": "Missing"}, {"book": r"C:\closed.xlsx"}])
def test_exact_target_missing_file_or_sheet_refuses_without_selection(office, navigate, params):
    run, raised = navigate
    assert run(links.build("excel", **params)) == 1
    assert office.excel.Selection is None and not raised
    assert office.excel.Workbooks.Count == 1


def test_all_instances_and_ambiguous_unsaved_name(office, navigate, monkeypatch):
    other = Excel()
    book = other.Workbooks.Add()
    monkeypatch.setattr(com, "apps", lambda kind, launch=False: [office.excel, other])
    run, _ = navigate
    assert run(links.build("excel", book=book.Name, sheet="Data", range="A1")) == 1
    book.FullName, book.Path = r"C:\second\Book1.xlsx", r"C:\second"
    assert run(links.build("excel", book=book.FullName, sheet="Data", range="C3")) == 0
    assert other.Selection.Address == "C3" and office.excel.Selection is None


def test_allowed_folders_and_unc_cannot_be_bypassed_by_short_name(office, navigate, monkeypatch):
    run, _ = navigate
    office.wb.FullName, office.wb.Path = r"C:\outside\Book1.xlsx", r"C:\outside"
    monkeypatch.setattr(config, "SETTINGS", config.load({"OFFICE_LIVE_ALLOWED_DIRS": r"C:\allowed"}))
    assert run(links.build("excel", book=office.wb.Name)) == 1
    monkeypatch.setattr(config, "SETTINGS", config.load({}))
    office.wb.FullName, office.wb.Path = r"\\server\share\Book1.xlsx", r"\\server\share"
    assert run(links.build("excel", book=office.wb.Name)) == 1
    assert office.excel.Selection is None


@pytest.fixture
def word_location(office):
    doc, app = office.doc, office.word
    selected, scrolled = [], []

    def make_range(start, end):
        return NS(Start=start, End=end, Select=lambda: selected.append((start, end)))

    doc.Range = make_range
    doc.Content.Start, doc.Content.End = 0, 12
    doc.Activate = lambda: setattr(app, "ActiveDocument", doc)
    doc.Paragraphs = Collection([NS(Range=make_range(0, 4)), NS(Range=make_range(4, 8)), NS(Range=make_range(8, 12))])
    doc.Tables = Collection([NS(Range=make_range(4, 12))])
    doc.Bookmarks = Collection([NS(Name="mark", Range=make_range(4, 8))])
    doc.Bookmarks.Exists = lambda name: name == "mark"
    app.ActiveWindow = NS(Hwnd=84, ScrollIntoView=lambda rng, top: scrolled.append((rng.Start, rng.End, top)))
    return selected, scrolled


@pytest.mark.parametrize("place,expected", [({"paragraph": "2"}, (4, 8)), ({"paragraphs": "2-3"}, (4, 12)),
                                           ({"table": "1"}, (4, 12)), ({"bookmark": "mark"}, (4, 8)), ({"span": "5-7"}, (5, 7))])
def test_word_links_select_and_scroll(office, navigate, word_location, place, expected):
    selected, scrolled = word_location
    run, raised = navigate
    before = office.doc.Content.Text, office.doc.Saved
    assert run(links.build("word", doc=office.doc.Name, **place)) == 0
    assert selected == [expected] and scrolled == [(*expected, True)] and raised == [84]
    assert (office.doc.Content.Text, office.doc.Saved) == before
    assert not office.word.starts and not undo.STACKS


@pytest.mark.parametrize("place", [{"paragraph": "4"}, {"table": "2"}, {"bookmark": "missing"}, {"span": "0-13"}])
def test_word_missing_location_does_not_activate(office, navigate, word_location, place):
    assert navigate[0](links.build("word", doc=office.doc.Name, **place)) == 1
    assert word_location == ([], []) and not navigate[1]


def test_write_links_snapshot_expansion_journal_and_undo_do_not_capture_log(office):
    office.call("office_journal", workbook=office.wb.Name, action="enable_sheet")
    out = office.call("excel_write_range", workbook=office.wb.Name, sheet="Data", cells="B2", values=[[1, 2], [3, 4]])
    link = out["links"][0]
    assert links.parse(link["uri"])[1]["range"] == "B2:C3"
    log = office.wb.Worksheets("Лог")
    assert log.Hyperlinks.Count == 1
    assert log.Hyperlinks(1).SubAddress == "'Data'!B2:C3" and log.Hyperlinks(1).Address == ""
    text = journal.journal_path(office.wb.Name).read_text(encoding="utf-8")
    assert links.markdown(link) in text
    entry = next(iter(undo.STACKS.values()))[-1]
    assert entry.areas == [{"sheet": "Data", "address": "B2:C3"}]
    office.call("office_undo", file=office.wb.Name)
    assert office.ws.Range("B2").Value is None and log.Hyperlinks.Count >= 1


def test_internal_hyperlinks_escape_apostrophes_and_skip_deleted_sheets(office):
    office.ws.Name = "O'Brien"
    office.call("office_journal", workbook=office.wb.Name, action="enable_sheet")
    out = office.call("excel_write_range", workbook=office.wb.Name, sheet=office.ws.Name, cells="A1", values=[[1]])
    log = office.wb.Worksheets("Лог")
    assert log.Hyperlinks(1).SubAddress == "'O''Brien'!A1"
    office.ws.Name = "Renamed"
    journal.append_sheet(office.excel, office.wb, "deleted", {}, links=out["links"])
    assert log.Hyperlinks.Count == 1


def test_read_links_limit_dedup_and_compare_both_books(office):
    office.ws.Range("A1:A25").Value = "match"
    out = office.call("excel_find", workbook=office.wb.Name, query="match")
    assert len(out["links"]) == 20 and out["links_truncated"]
    other = office.excel.Workbooks.Add()
    other.ActiveSheet.Range("B1").Value = "different"
    out = office.call("excel_compare_ranges", workbook_a=office.wb.Name, sheet_a="Data", cells_a="A1", workbook_b=other.Name, sheet_b="Data", cells_b="B1")
    assert {links.parse(link["uri"])[1]["book"] for link in out["links"]} == {office.wb.Name, other.Name}


def test_trace_links_and_no_write_links_for_preview(office):
    office.ws.Range("B1").Formula = "=A1"
    office.ws.direct[("precedents", "B1")] = ["A1"]
    out = office.call("excel_trace_formula", workbook=office.wb.Name, sheet="Data", cell="B1")
    assert {links.parse(link["uri"])[1]["range"] for link in out["links"]} == {"A1", "B1"}
    office.ws.Range("A2").Value = " text "
    out = office.call("excel_clean_text", workbook=office.wb.Name, sheet="Data", cells="A2", operations=["trim"])
    assert "links" not in out


@pytest.mark.parametrize("tool,result", [
    ("word_find", {"matches": [{"paragraph": 2}, {"story": "header"}]}),
    ("word_get_structure", {"headings": [{"paragraph": 2}]}),
    ("word_insert_text", {"paragraphs": [2, 3]}),
    ("word_write_table", {"table": 1}),
    ("word_manage_bookmarks", {"bookmark": "mark"}),
])
def test_word_result_locations_are_checked(office, word_location, tool, result):
    recording = NS(tool=tool, kind="read" if tool in {"word_find", "word_get_structure"} else "write", args={},
                   targets={("document", office.doc.Name, 84): (office.word, office.doc)}, pending={})
    navigation.result_links(recording, result)
    assert len(result["links"]) == 1 and not result["links_truncated"]
    assert word_location == ([], [])


def test_unknown_deleted_and_unsafe_places_produce_no_invented_links(office, word_location):
    for tool, data in (("word_delete_paragraphs", {"deleted_paragraphs": [1, 2]}), ("word_insert_text", {"ok": True}),
                       ("word_write_table", {"table": 999})):
        recording = NS(tool=tool, kind="write", args={}, targets={("document", "doc", 1): (office.word, office.doc)}, pending={})
        navigation.result_links(recording, data)
        assert "links" not in data


def test_markdown_label_cannot_inject_a_second_link():
    link = {"label": "[click](javascript:evil) <tag>", "uri": links.build("word", doc="Draft")}
    assert links.markdown(link).startswith(r"[\[click\]\(javascript:evil\) \<tag\>]")


def test_write_links_with_undo_disabled_use_actual_result(office, monkeypatch):
    from dataclasses import replace

    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo=False))
    out = office.call("excel_write_range", workbook=office.wb.Name, sheet="Data", cells="B2", values=[[1, 2], [3, 4]])
    assert links.parse(out["links"][0]["uri"])[1]["range"] == "B2:C3"
    other = office.excel.Workbooks.Add()
    out = office.call("excel_copy_range", workbook=office.wb.Name, sheet="Data", source="B2:C3",
                      dest_workbook=other.Name, dest_sheet="Data", dest_cell="E4", what="values")
    assert [links.parse(link["uri"])[1] for link in out["links"]] == [{"book": other.Name, "sheet": "Data", "range": "E4:F5"}]
    assert not undo.STACKS


def test_office_link_readonly_registration_and_reject_mixed_arguments(office, monkeypatch):
    from dataclasses import replace

    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, mode="readonly"))
    assert office.call("office_link", workbook=office.wb.Name)["links"]
    assert registry.CATALOG["office_link"].registered
    for arguments in ({}, {"workbook": office.wb.Name, "document": office.doc.Name},
                      {"workbook": office.wb.Name, "paragraph": 1}, {"document": office.doc.Name, "end_paragraph": 2},
                      {"document": office.doc.Name, "paragraph": 1, "table": 1}):
        with pytest.raises(ToolError):
            office.call("office_link", **arguments)


def test_issues_links_use_reported_cells_and_columns(office):
    recording = NS(tool="excel_find_issues", kind="read", args={}, targets={("workbook", office.wb.Name, 42): (office.excel, office.wb)}, pending={})
    result = {"sheet": "Data", "issues": [{"examples": [{"cell": "B2"}, {"cell": "C"}, {"cell": "B2"}]}]}
    navigation.result_links(recording, result)
    assert [links.parse(link["uri"])[1]["range"] for link in result["links"]] == ["B2", "C:C"]


def test_journal_links_stay_in_their_own_document(office):
    first = navigation.item("excel", office.wb, sheet="Data", range="A1")
    other = office.excel.Workbooks.Add()
    second = navigation.item("excel", other, sheet="Data", range="B2")
    journal.append([f"workbook:{office.wb.Name}", f"workbook:{other.Name}"], "copy", {}, links=[first, second])
    one = journal.journal_path(office.wb.Name).read_text(encoding="utf-8")
    two = journal.journal_path(other.Name).read_text(encoding="utf-8")
    assert first["uri"] in one and second["uri"] not in one
    assert second["uri"] in two and first["uri"] not in two


def test_link_metadata_failure_keeps_successful_write_and_undo(office, monkeypatch):
    def fail(*args):
        raise RuntimeError("fixture: target closed during link metadata read")

    monkeypatch.setattr(navigation, "_result_links", fail)
    out = office.call("excel_write_range", workbook=office.wb.Name, sheet="Data", cells="A1", values=[[42]])
    assert out["ok"] and out["undo"] == "available" and "links" not in out
    assert office.ws.Range("A1").Value == 42
    office.call("office_undo", file=office.wb.Name)
    assert office.ws.Range("A1").Value is None


def test_disabled_service_log_hyperlinks_are_never_snapshotted(office):
    office.call("office_journal", workbook=office.wb.Name, action="enable_sheet")
    office.call("excel_write_range", workbook=office.wb.Name, sheet="Data", cells="A1", values=[[42]])
    log = office.wb.Worksheets("Лог")
    assert log.Hyperlinks.Count == 1
    office.call("office_journal", workbook=office.wb.Name, action="disable_sheet")
    out = office.call("excel_write_range", workbook=office.wb.Name, sheet="Лог", cells="Z1", values=[[1]])
    assert "service log sheet" in out["undo"]
    entry = next(iter(undo.STACKS.values()))[-1]
    assert not entry.areas and not entry.backup_sheets
