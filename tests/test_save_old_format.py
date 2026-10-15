"""Save/close contract on fake COM objects; no Office required."""

from types import SimpleNamespace

import pytest
import pywintypes

from office_live import excel_core as core
from office_live.errors import ToolError


class FlagObject:
    def __setattr__(self, name, value):
        if name in {"DisplayAlerts", "CheckCompatibility", "Saved"} and "events" in self.__dict__:
            self.events.append((name, value))
        super().__setattr__(name, value)


class Workbook(FlagObject):
    def __init__(self, app, path):
        self.app = app
        self.FullName, self.Path, self.Name = str(path), str(path.parent), path.name
        self.FileFormat, self.ReadOnly, self.Saved = 56, False, False
        self._compatibility = True
        self.restore_error = self.save_error = None
        self.dirty_on_restore = False
        self.saved_after_save = True
        self.events = app.events

    @property
    def CheckCompatibility(self):
        return self._compatibility

    @CheckCompatibility.setter
    def CheckCompatibility(self, value):
        if value and self.restore_error:
            raise self.restore_error
        self._compatibility = value
        if value and self.dirty_on_restore:
            self.Saved = False

    def Save(self):
        self.events.append(("Save", self.app.DisplayAlerts, self.CheckCompatibility))
        if self.save_error:
            raise self.save_error
        self.Saved = self.saved_after_save

    def Close(self, save):
        self.events.append(("Close", save))


@pytest.fixture
def book(tmp_path, monkeypatch):
    app = FlagObject()
    app.DisplayAlerts = True
    app.events = []
    wb = Workbook(app, tmp_path / "old.xls")
    monkeypatch.setattr(core, "pick_workbook", lambda *a, **kw: (app, wb))
    monkeypatch.setattr(core, "_excel_running", lambda: False)
    return SimpleNamespace(app=app, wb=wb, events=app.events)


@pytest.mark.parametrize("file_format", [56, 43, 39, 33, 29, 16, 17, 18, -4143])
def test_legacy_save_suppresses_prompts_and_restores_flags(book, file_format):
    book.wb.FileFormat = file_format
    result = core.excel_save(book.wb.Name)
    assert ("Save", False, False) in book.events
    assert book.app.DisplayAlerts is True and book.wb.CheckCompatibility is True
    assert result["saved"] is True and result["compatibility_check_skipped"] is True
    assert all(text in result["note"] for text in ("Excel 97-2003", "lost", "excel_save_as", ".xlsx", ".xlsm"))
    assert "without" not in result["note"].split(".xlsx")[0]  # no "lossless via .xlsx" promise
    # CheckCompatibility alone left a second modal question during Save (live probe), so alerts are off too.
    assert book.events[-2:] == [("CheckCompatibility", True), ("DisplayAlerts", True)]
    assert "default answer" in result["note"]


@pytest.mark.parametrize("restore_fails", [False, True])
def test_save_error_is_preserved_and_restoration_is_independent(book, restore_fails):
    error = pywintypes.com_error(-2147352567, "original save failure", None, None)
    book.wb.save_error = error
    if restore_fails:
        book.wb.restore_error = pywintypes.com_error(-2147352567, "restore failure", None, None)
    with pytest.raises(pywintypes.com_error) as exc:
        core.excel_save(book.wb.Name)
    assert exc.value is error
    assert book.app.DisplayAlerts is True
    assert ("CheckCompatibility", True) in book.events
    if not restore_fails:
        assert book.wb.CheckCompatibility is True
    assert ("Saved", True) not in book.events


def test_compatibility_restore_error_does_not_prevent_alerts_restore(book):
    book.wb.restore_error = pywintypes.com_error(-2147352567, "restore failure", None, None)
    assert core.excel_save(book.wb.Name)["saved"] is True
    assert book.app.DisplayAlerts is True


@pytest.mark.parametrize("saved_after", [True, False])
def test_only_compatibility_flag_dirtying_is_cleared(book, saved_after):
    book.wb.dirty_on_restore = True
    book.wb.saved_after_save = saved_after
    result = core.excel_save(book.wb.Name)
    assert result["saved"] is saved_after
    assert book.wb.Saved is saved_after


@pytest.mark.parametrize("file_format,extension", [(51, "xlsx"), (52, "xlsm"), (50, "xlsb"), (6, "csv"), (51, "xls")])
def test_nonlegacy_save_never_writes_prompt_flags(book, file_format, extension):
    book.wb.FileFormat = file_format
    book.wb.FullName = str(book.wb.FullName).replace(".xls", "." + extension)
    result = core.excel_save(book.wb.Name)
    assert result == {"ok": True, "workbook": book.wb.Name, "path": book.wb.FullName, "saved": True}
    assert book.events == [("Save", True, True), ("Saved", True)]


@pytest.mark.parametrize("extension,skipped", [("XLS", True), ("xlt", True), ("xla", True), ("xlsx", False)])
def test_missing_file_format_falls_back_to_extension(book, extension, skipped):
    del book.wb.FileFormat
    book.wb.FullName = book.wb.FullName.replace(".xls", "." + extension)
    assert core._save_workbook(book.app, book.wb) is skipped


def test_com_file_format_failure_falls_back_to_extension(book):
    class UnreadableFormat(Workbook):
        @property
        def FileFormat(self):
            raise pywintypes.com_error(-2147352567, "unavailable", None, None)

    book.wb.__class__ = UnreadableFormat
    assert core._save_workbook(book.app, book.wb) is True


@pytest.mark.parametrize("alerts,compatibility", [(False, True), (True, False), (False, False)])
def test_original_false_flags_are_preserved(book, alerts, compatibility):
    book.app.DisplayAlerts, book.wb.CheckCompatibility = alerts, compatibility
    core.excel_save(book.wb.Name)
    assert book.app.DisplayAlerts is alerts and book.wb.CheckCompatibility is compatibility


def test_close_saves_under_same_flags_then_closes_without_save(book):
    book.wb.dirty_on_restore = True
    result = core.excel_close_workbook(book.wb.Name, save=True)
    assert ("Save", False, False) in book.events
    assert book.events[-1] == ("Close", False)
    assert book.app.DisplayAlerts is True and book.wb.CheckCompatibility is True and book.wb.Saved is True
    assert result["saved_before_close"] is True and result["compatibility_check_skipped"] is True
    assert result["note"] == core._COMPATIBILITY_NOTE


def test_close_save_failure_leaves_book_open(book):
    book.wb.save_error = pywintypes.com_error(-2147352567, "save failure", None, None)
    with pytest.raises(pywintypes.com_error):
        core.excel_close_workbook(book.wb.Name, save=True)
    assert not any(event[0] == "Close" for event in book.events)
    assert book.app.DisplayAlerts is True and book.wb.CheckCompatibility is True


def test_close_readonly_legacy_book_refuses_before_suppressing_alerts(book):
    book.wb.ReadOnly = True
    with pytest.raises(ToolError, match="read-only"):
        core.excel_close_workbook(book.wb.Name, save=True)
    assert book.events == []


def test_close_legacy_book_refuses_if_save_did_not_mark_it_saved(book):
    book.wb.saved_after_save = False
    with pytest.raises(ToolError, match="left open"):
        core.excel_close_workbook(book.wb.Name, save=True)
    assert ("Save", False, False) in book.events
    assert not any(event[0] == "Close" for event in book.events)
    assert book.app.DisplayAlerts is True and book.wb.CheckCompatibility is True


def test_close_modern_book_preserves_response_and_prompt_flags(book):
    book.wb.FileFormat = 51
    assert core.excel_close_workbook(book.wb.Name, save=True) == {
        "ok": True, "closed": book.wb.Name, "saved_before_close": True, "open_workbooks": [],
    }
    assert book.events == [("Save", True, True), ("Saved", True), ("Close", False)]

