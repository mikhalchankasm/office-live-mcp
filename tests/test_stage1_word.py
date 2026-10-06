"""Сравнение документов и сортировка таблиц без запущенного Word."""

import copy
from dataclasses import replace
from types import SimpleNamespace as NS

import pytest
import pywintypes

from office_live import com, config, undo, word_tables
from office_live.errors import AppBusyError, ToolError
from tests.history_fakes import Collection, Document, Word, fake_office  # noqa: F401


def error():
    return pywintypes.com_error(-2147352567, "simulated COM failure", None, None)


def add_doc(app, name, path=None):
    doc = Document(app)
    doc.Name, doc.FullName, doc.Path = name, str(path) if path else name, str(path.parent) if path else ""
    app.Documents.items.append(doc)
    return doc


@pytest.fixture
def comparison(office):
    o = office
    revised = add_doc(o.word, "Revised.docx")
    revised.Content.Text = "revised"
    calls, opened, activated = [], [], []
    o.word.ScreenUpdating = True
    o.word.ActiveWindow = NS(Activate=lambda: activated.append(o.word.ScreenUpdating))

    def compare(*args):
        calls.append(args)
        assert len(args) == 16 and args[2] == 2 and args[-1] is True
        result = add_doc(o.word, "Compared.docx")
        result.Saved = False
        result.Revisions = Collection([NS(Type=code, Author="Tester", Range=NS(Text="x" * 250)) for code in [1, 2, 3, 14, 15, 16, 17, 18] * 5])
        return result

    def open_file(*args):
        assert len(args) == 15
        # Visible=True: скрытое открытие меняет активный документ пользователя; окно прячется сразу после открытия
        assert args[1:4] == (False, True, False) and args[11] is True and args[14] is True
        assert all(args[i] == "__office_live_no_password__" for i in (4, 5, 7, 8))
        assert o.word.AutomationSecurity == 3 and o.word.ScreenUpdating is False
        from pathlib import Path

        path = Path(args[0])
        doc = add_doc(o.word, path.name, path)
        doc.ReadOnly = True
        doc.Windows = Collection([NS(Visible=True)])
        opened.append(doc)
        return doc

    o.word.CompareDocuments = compare
    o.word.Documents.Open = open_file
    return NS(o=o, revised=revised, calls=calls, opened=opened, activated=activated)


def compare(c, **kw):
    return c.o.call("word_compare_documents", **{"original": c.o.doc.Name, "revised": c.revised.Name, **kw})


def test_compare_open_documents_source_fingerprints_saved_and_revision_report(comparison):
    c = comparison
    before = [(undo.word_fingerprint(d), d.Saved) for d in (c.o.doc, c.revised)]
    result = compare(c, compare_formatting=False, compare_case=False, granularity="character", author="Custom")
    assert c.calls[0][0:4] == (c.o.doc, c.revised, 2, 0)
    assert c.calls[0][4:14] == (False, False, True, True, True, True, True, True, True, True)
    assert c.calls[0][-2:] == ("Custom", True)
    assert result["document"] == "Compared.docx" and result["saved"] is False
    assert result["counts"] == {"insertion": 10, "deletion": 10, "format": 5, "move": 10, "other": 5}
    assert len(result["revisions"]) == 30 and result["truncated"]
    assert result["revisions"][0]["text_truncated"] and len(result["revisions"][0]["text"]) == 200
    assert [(undo.word_fingerprint(d), d.Saved) for d in (c.o.doc, c.revised)] == before
    assert not undo.STACKS and not c.o.word.starts


@pytest.mark.parametrize("already_open", [False, True])
def test_compare_allowed_file_path_open_or_temporary(comparison, already_open):
    c = comparison
    path = c.o.tmp / "Revised.docx"
    path.write_text("fake doc")
    if already_open:
        c.revised.Path, c.revised.FullName = str(path.parent), str(path)
    else:
        c.o.word.Documents.items.remove(c.revised)
    result = compare(c, revised=str(path))
    assert result["document"] == "Compared.docx"
    assert len(c.opened) == (0 if already_open else 1)
    assert all(d not in c.o.word.Documents.items for d in c.opened)
    assert all(d.Windows(1).Visible is False for d in c.opened)
    assert c.activated == ([] if already_open else [False])  # окно пользователя вернули до включения перерисовки
    assert c.o.word.AutomationSecurity == 1 and c.o.word.ScreenUpdating is True


def test_compare_temporary_copy_closed_when_its_window_cannot_be_hidden(comparison):
    c = comparison
    path = c.o.tmp / "Revised.docx"
    path.write_text("fake doc")
    c.o.word.Documents.items.remove(c.revised)
    original_open = c.o.word.Documents.Open

    class Unhideable:
        Visible = True

        def __setattr__(self, name, value):
            raise error()

    def open_file(*args):
        doc = original_open(*args)
        doc.Windows = Collection([Unhideable()])
        return doc

    c.o.word.Documents.Open = open_file
    with pytest.raises(ToolError):
        compare(c, revised=str(path))
    assert c.opened[0] not in c.o.word.Documents.items and not c.calls
    assert c.o.word.ScreenUpdating is True and c.o.word.AutomationSecurity == 1


def test_compare_two_closed_paths_use_one_instance(comparison):
    c = comparison
    paths = [c.o.tmp / name for name in ("original.docx", "new.docx")]
    for path in paths:
        path.write_text("fake doc")
    compare(c, original=str(paths[0]), revised=str(paths[1]))
    assert len(c.opened) == 2 and all(d.Application is c.o.word for d in c.opened)
    assert all(d not in c.o.word.Documents.items for d in c.opened)


def test_compare_same_closed_path_is_opened_and_closed_once(comparison):
    c = comparison
    path = c.o.tmp / "same.docx"
    path.write_text("fake")
    compare(c, original=str(path), revised=str(path))
    assert len(c.opened) == 1 and c.calls[0][0] is c.calls[0][1]
    assert c.opened[0] not in c.o.word.Documents.items


def test_compare_cleanup_attempts_every_temporary_even_if_one_close_fails(comparison):
    c = comparison
    first, second = c.o.tmp / "first.docx", c.o.tmp / "second.docx"
    first.write_text("fake")
    second.write_text("fake")
    original_open = c.o.word.Documents.Open

    def open_file(*args):
        doc = original_open(*args)
        if args[0] == str(second):
            doc.Close = lambda *a: (_ for _ in ()).throw(error())
        return doc

    c.o.word.Documents.Open = open_file
    with pytest.raises(ToolError, match="Compared.docx.*Could not close"):
        compare(c, original=str(first), revised=str(second))
    assert c.opened[0] not in c.o.word.Documents.items and c.opened[1] in c.o.word.Documents.items


def test_compare_existing_revisions_warning_without_accepting_sources(comparison):
    c = comparison
    c.o.doc.Revisions = Collection([NS(Type=1)])
    assert "treated as accepted" in compare(c)["warnings"][0]
    assert c.o.doc.Revisions.Count == 1


@pytest.mark.parametrize("which", ["original", "revised"])
def test_compare_disallowed_paths_before_any_open(comparison, monkeypatch, which):
    c = comparison
    allowed = c.o.tmp / "allowed"
    allowed.mkdir()
    good = allowed / "good.docx"
    good.write_text("fake")
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, allowed_dirs=(allowed,)))
    kw = {"original": str(good), "revised": str(good)}
    kw[which] = str(c.o.tmp / "outside.docx")
    with pytest.raises(ToolError, match="outside the allowed"):
        compare(c, **kw)
    assert not c.calls and not c.opened


@pytest.mark.parametrize("kw,message", [({"granularity": "line"}, "granularity"), ({"original": " "}, "required"),
                                       ({"revised": " "}, "required"), ({"author": " "}, "author"), ({"revised": "missing.docx"}, "not found")])
def test_compare_input_refusals(comparison, kw, message):
    with pytest.raises(ToolError, match=message):
        compare(comparison, **kw)
    assert not comparison.calls and not comparison.opened


def test_compare_missing_second_file_before_opening_first(comparison):
    c = comparison
    first = c.o.tmp / "first.docx"
    first.write_text("fake")
    with pytest.raises(ToolError, match="File not found"):
        compare(c, original=str(first), revised=str(c.o.tmp / "missing.docx"))
    assert not c.calls and not c.opened


def test_compare_different_instances_and_autosave_refused(comparison, monkeypatch):
    c = comparison
    second = Word()
    second.ActiveDocument.Name = second.ActiveDocument.FullName = "Other.docx"
    monkeypatch.setattr(com, "apps", lambda *a, **kw: [c.o.word, second])
    with pytest.raises(ToolError, match="same Word instance"):
        compare(c, revised="Other.docx")
    c.revised.AutoSaveOn = True
    with pytest.raises(ToolError, match="AutoSave"):
        compare(c)
    assert not c.calls and not c.opened


@pytest.mark.parametrize("failure", ["compare", "second_open", "macros", "report"])
def test_compare_failures_close_only_owned_temporary_docs(comparison, failure):
    c = comparison
    paths = [c.o.tmp / name for name in ("first.docx", "second.docx")]
    for path in paths:
        path.write_text("fake")
    original_open = c.o.word.Documents.Open
    original_compare = c.o.word.CompareDocuments

    def open_file(*args):
        if failure == "second_open" and c.opened:
            raise error()
        if failure == "macros":
            raise error()
        return original_open(*args)

    def compare_file(*args):
        if failure == "compare":
            raise error()
        result = original_compare(*args)
        if failure == "report":
            result.Revisions = NS(Count="invalid")
        return result

    c.o.word.Documents.Open = open_file
    c.o.word.CompareDocuments = compare_file
    with pytest.raises(ToolError):
        compare(c, original=str(paths[0]), revised=str(paths[1]))
    assert c.o.doc in c.o.word.Documents.items and c.revised in c.o.word.Documents.items
    assert all(d not in c.o.word.Documents.items for d in c.opened)
    assert c.o.word.AutomationSecurity == 1


class Table:
    def __init__(self, data):
        self.data = copy.deepcopy(data)
        self.Uniform, self.NestingLevel = True, 1
        self.Rows, self.Columns = NS(Count=len(data)), NS(Count=len(data[0]))
        self.Range = NS(Cells=Collection([None] * (len(data) * len(data[0]))), LanguageID=1049)
        self.Tables = Collection()
        self.calls, self.fail = [], False

    def Cell(self, row, col, /):
        return NS(Range=NS(Text=self.data[row - 1][col - 1] + "\r\x07"))

    def Sort(self, *args):
        assert len(args) == 17 and args[11:16] == (False,) * 5
        self.calls.append(args)
        start = int(args[0])
        body = self.data[start:]
        for i in (7, 4, 1):
            if args[i] == "":
                continue
            col, kind, order = int(args[i]) - 1, args[i + 1], args[i + 2]

            def key(row, col=col, kind=kind):
                value = row[col]
                if kind:
                    return word_tables._sort_value(value, "number" if kind == 1 else "date", args[16])
                return value if args[10] else value.casefold()

            body.sort(key=key, reverse=bool(order))
        self.data[start:] = body
        if self.fail:
            raise error()


def table(o, data=None):
    tbl = Table(data or [["Name", "Value", "Date"], ["b", "1 234,5", "2026-10-05"], ["a", "12,5", "2025-01-01"], ["a", "12,5", "2024-01-01"]])
    o.doc.Tables = Collection([tbl])
    return tbl


def sort(o, **kw):
    return o.call("word_sort_table", **{"document": o.doc.Name, "table_index": 1, "keys": [{"column": 1}], **kw})


@pytest.mark.parametrize("header,desc", [(True, False), (True, True), (False, False), (False, True)])
def test_sort_one_key_header_desc_and_undo(office, header, desc):
    o = office
    tbl = table(o, [["z"], ["b"], ["c"], ["a"]])
    before = copy.deepcopy(tbl.data)
    result = sort(o, header=header, keys=[{"column": 1, "order": "desc" if desc else "asc"}])
    assert result["sorted_rows"] == 4 - int(header) and result["undo"] == "available"
    assert result["values"] == (([["z"]] if header else []) + sorted(before[int(header):], reverse=desc))
    assert tbl.calls[0][:11] == (header, "1", 0, int(desc), "", 0, 0, "", 0, 0, False)
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.Tables(1).data == before


def test_sort_three_keys_numbers_dates_case_and_report(office):
    o = office
    tbl = table(o)
    result = sort(o, keys=[{"column": 1}, {"column": 2, "type": "number", "order": "desc"}, {"column": 3, "type": "date"}], case_sensitive=True)
    assert result["values"][1:] == [["a", "12,5", "2024-01-01"], ["a", "12,5", "2025-01-01"], ["b", "1 234,5", "2026-10-05"]]
    assert tbl.calls[0] == (True, "1", 0, 0, "2", 1, 1, "3", 2, 0, True, False, False, False, False, False, 1049)


def test_sort_returns_twenty_rows(office):
    table(office, [["H"]] + [[str(x)] for x in range(30, 0, -1)])
    result = sort(office, keys=[{"column": 1, "type": "number"}])
    assert result["sorted_rows"] == 30 and len(result["values"]) == 20 and result["truncated"]


@pytest.mark.parametrize("keys", [[], [{}] * 4, [None], [{"column": 0}], [{"column": True}], [{"column": "1"}],
                                  [{"column": 4}], [{"column": 1}, {"column": 1}], [{"column": 1, "order": "bad"}], [{"column": 1, "type": "bad"}]])
def test_sort_bad_keys_no_record_or_edits(office, keys):
    tbl = table(office)
    before = copy.deepcopy(tbl.data)
    with pytest.raises(ToolError):
        sort(office, keys=keys)
    assert tbl.data == before and not tbl.calls and not office.word.starts and not undo.STACKS


@pytest.mark.parametrize("problem,message", [("absent", "no tables"), ("index", "does not exist"), ("index_zero", "table_index"),
                                           ("one_row", "no data"), ("nonuniform", "nonuniform"), ("merged", "Merged"), ("nested", "Nested"),
                                           ("inner", "Nested"), ("tracking", "Track Changes"), ("protected", "protected"),
                                           ("number", "Unparseable"), ("date", "Unparseable"), ("autosave", "AutoSave"),
                                           ("busy_record", "another custom")])
def test_sort_state_refusals_before_mutation(office, monkeypatch, problem, message):
    o = office
    tbl = table(o)
    kw = {}
    if problem == "absent":
        o.doc.Tables = Collection()
    elif problem in {"index", "index_zero"}:
        kw["table_index"] = 2 if problem == "index" else 0
    elif problem == "one_row":
        tbl = table(o, [["header"]])
    elif problem == "nonuniform":
        tbl.Uniform = False
    elif problem == "merged":
        tbl.Range.Cells.items.pop()
    elif problem == "nested":
        tbl.Tables.items.append(NS())
    elif problem == "inner":
        tbl.NestingLevel = 2
    elif problem == "tracking":
        o.doc.TrackRevisions = True
    elif problem == "protected":
        o.doc.ProtectionType = 1
    elif problem in {"number", "date"}:
        tbl.data[1][1] = "not a " + problem
        kw["keys"] = [{"column": 2, "type": problem}]
    elif problem == "autosave":
        o.doc.AutoSaveOn = True
    elif problem == "undo_off":
        monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo=False))
    elif problem == "busy_record":
        o.word.UndoRecord.IsRecordingCustomRecord = True
    before = copy.deepcopy(tbl.data)
    with pytest.raises(ToolError, match=message):
        sort(o, **kw)
    assert tbl.data == before and not tbl.calls and not o.word.starts and not undo.STACKS


def test_sort_later_user_edit_blocks_undo_and_force_restores(office):
    o = office
    table(o)
    original = undo.word_fingerprint(o.doc)
    sort(o)
    o.doc.save_state()
    o.doc.Tables(1).data[1][0] = "user"
    with pytest.raises(ToolError, match="later edits"):
        o.call("office_undo", file=o.doc.Name)
    result = o.call("office_undo", file=o.doc.Name, force=True)
    assert result["native_steps"] == 2 and undo.word_fingerprint(o.doc) == original


def test_sort_partial_failure_retains_native_undo(office):
    o = office
    tbl = table(o)
    before = copy.deepcopy(tbl.data)
    tbl.fail = True
    with pytest.raises(ToolError, match="applied=3.*office_undo will restore"):
        sort(o)
    o.call("office_undo", file=o.doc.Name)
    assert o.doc.Tables(1).data == before


def test_sort_busy_after_write_keeps_native_record_with_force_guard(office, monkeypatch):
    o = office
    tbl = table(o)
    before = copy.deepcopy(tbl.data)
    busy = False
    original_sort, original_cell, fingerprint = Table.Sort, Table.Cell, undo.word_fingerprint

    def fail(self, *args):
        nonlocal busy
        original_sort(self, *args)
        busy = True
        raise AppBusyError("Office busy")

    def read_cell(self, *args):
        if busy:
            raise AppBusyError("Office busy")
        return original_cell(self, *args)

    def stamp(doc, styles=True):
        if busy:
            raise AppBusyError("Office busy")
        return fingerprint(doc, styles=styles)

    monkeypatch.setattr(Table, "Sort", fail)
    monkeypatch.setattr(Table, "Cell", read_cell)
    monkeypatch.setattr(undo, "word_fingerprint", stamp)
    with pytest.raises(ToolError, match="applied=unknown.*force=true"):
        sort(o)
    busy = False
    with pytest.raises(ToolError, match="force=true"):
        o.call("office_undo", file=o.doc.Name)
    o.call("office_undo", file=o.doc.Name, force=True)
    assert o.doc.Tables(1).data == before


@pytest.mark.parametrize("value,kind,language,valid", [("1 234,5", "number", 1049, True), ("1,234.5", "number", 1033, True),
                                                   ("2026-10-05", "date", 1049, True), ("31.02.2026", "date", 1049, False),
                                                   ("05.10.26", "date", 1049, False), ("", "number", 1049, False),
                                                   ("abc", "number", 1049, False), ("1\n2", "number", 1049, False)])
def test_native_sort_parser_no_office(value, kind, language, valid):
    assert (word_tables._sort_value(value, kind, language) is not None) is valid


def test_sort_works_without_undo_when_owner_disabled_it(office, monkeypatch):
    o = office
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, undo=False))
    tbl = table(o, [["h"], ["b"], ["a"]])
    assert sort(o)["values"] == [["h"], ["a"], ["b"]]
    assert tbl.calls and not undo.STACKS


def test_compare_continues_when_previous_window_cannot_be_reactivated(comparison):
    c = comparison
    path = c.o.tmp / "Revised.docx"
    path.write_text("fake doc")
    c.o.word.Documents.items.remove(c.revised)

    def refuse():
        raise error()

    c.o.word.ActiveWindow = NS(Activate=refuse)
    result = compare(c, revised=str(path))
    assert result["document"] == "Compared.docx" and c.calls
    assert all(d not in c.o.word.Documents.items for d in c.opened) and c.o.word.ScreenUpdating is True
