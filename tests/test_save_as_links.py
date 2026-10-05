import os

import pytest

from office_live import config, links, navigation, protocol
from office_live.errors import ToolError
from tests.history_fakes import fake_office  # noqa: F401


@pytest.fixture
def save(office, monkeypatch):
    office.word.DisplayAlerts = -1

    def rename(obj, path, fmt):
        obj.FullName, obj.Path, obj.Name = path, os.path.dirname(path), os.path.basename(path)

    monkeypatch.setattr(office.wb, "SaveAs", lambda path, fmt: rename(office.wb, path, fmt), raising=False)
    monkeypatch.setattr(office.doc, "SaveAs2", lambda path, fmt: rename(office.doc, path, fmt), raising=False)
    return lambda app, name, path: office.call(app + "_save_as", **{"workbook" if app == "excel" else "document": name, "path": str(path)})


@pytest.mark.parametrize("app,field", [("excel", "book"), ("word", "doc")])
def test_save_as_new_links_and_session_alias_chains(office, save, app, field):
    obj = office.wb if app == "excel" else office.doc
    old = obj.Name
    first = office.tmp / ("first.xlsx" if app == "excel" else "first.docx")
    second = office.tmp / ("second.xlsx" if app == "excel" else "second.docx")
    result = save(app, old, first)
    assert links.parse(result["links"][0]["uri"])[1][field] == str(first)
    assert navigation.resolve(app, {field: old})[1] is obj
    result = save(app, obj.Name, second)
    assert links.parse(result["links"][0]["uri"])[1][field] == str(second)
    for alias in (old, str(first), first.name):
        assert navigation.resolve(app, {field: alias})[1] is obj
    assert office.call("office_link", **{"workbook" if app == "excel" else "document": old})["links"] == result["links"]


@pytest.mark.parametrize("app", ["excel", "word"])
def test_failed_save_does_not_record_alias(office, save, monkeypatch, app):
    obj = office.wb if app == "excel" else office.doc

    def fail(*args):
        raise ToolError("save refused")

    monkeypatch.setattr(obj, "SaveAs" if app == "excel" else "SaveAs2", fail)
    with pytest.raises(ToolError, match="save refused"):
        save(app, obj.Name, office.tmp / ("failed.xlsx" if app == "excel" else "failed.docx"))
    assert not navigation.RENAMES


def test_old_exact_name_wins_over_alias(office, save):
    old = office.wb.Name
    save("excel", old, office.tmp / "renamed.xlsx")
    new = office.excel.Workbooks.Add()
    new.Name = new.FullName = old
    assert navigation.resolve("excel", {"book": old})[1] is new


def test_alias_never_bypasses_ambiguity_or_allowed_folders(office, save, monkeypatch):
    old = office.wb.Name
    save("excel", old, office.tmp / "renamed.xlsx")
    second = office.excel.Workbooks.Add()
    third = office.excel.Workbooks.Add()
    second.Name = third.Name = old
    with pytest.raises(ToolError, match="several"):
        navigation.resolve("excel", {"book": old})
    office.excel.Workbooks.items = [office.wb]
    monkeypatch.setattr(config, "SETTINGS", config.load({"OFFICE_LIVE_ALLOWED_DIRS": str(office.tmp / "allowed")}))
    with pytest.raises(ToolError):
        navigation.resolve("excel", {"book": old})


def test_conflicting_alias_refused_and_cycle_collapsed(office):
    navigation.remember_rename("excel", ("A", "A"), r"C:\B.xlsx")
    navigation.remember_rename("excel", ("A", "A"), r"C:\C.xlsx")
    assert navigation.renamed_target("excel", "A") is None
    navigation.RENAMES.clear()
    navigation.remember_rename("excel", ("A.xlsx", r"C:\A.xlsx"), r"C:\B.xlsx")
    navigation.remember_rename("excel", ("B.xlsx", r"C:\B.xlsx"), r"C:\A.xlsx")
    assert navigation.renamed_target("excel", "A.xlsx") == r"c:\a.xlsx"


@pytest.mark.parametrize("app,field", [("excel", "book"), ("word", "doc")])
def test_separate_handler_without_alias_has_clear_error_and_no_focus(office, save, app, field, monkeypatch):
    obj = office.wb if app == "excel" else office.doc
    old = obj.Name
    save(app, old, office.tmp / ("renamed.xlsx" if app == "excel" else "renamed.docx"))
    monkeypatch.setattr(navigation, "RENAMES", {})  # another process has no session memory
    messages = []
    monkeypatch.setattr(protocol, "error_dialog", messages.append)
    monkeypatch.setattr(protocol, "foreground", lambda _: pytest.fail("Missing target must not focus anything"))
    monkeypatch.setattr(protocol, "allowed_settings", lambda: config.SETTINGS)
    assert protocol.open_link([links.build(app, **{field: old})]) == 1
    assert "renamed or closed" in messages[0]
    assert office.excel.Workbooks.Count == 1 and office.word.Documents.Count == 1


def test_closed_alias_does_not_open_a_file(office, save):
    old = office.wb.Name
    save("excel", old, office.tmp / "renamed.xlsx")
    office.wb.Close(False)
    with pytest.raises(ToolError, match="renamed or closed"):
        navigation.resolve("excel", {"book": old})
    assert office.excel.Workbooks.Count == 0


def test_restore_uses_successful_save_as_alias(office, save, monkeypatch):
    from office_live import window
    from tests.window_fakes import FakeWin32

    desktop = FakeWin32()
    monkeypatch.setattr(window, "Win32", lambda: desktop)
    office.call("office_window", action="arrange", include_chat=False)
    save("excel", office.wb.Name, office.tmp / "renamed.xlsx")
    assert office.call("office_window", action="restore")["ok"]
