"""Регрессии аудита: выбор цели, зона доступа, AutoSave, события, режим readonly по действиям, журнал. Office не нужен."""

import json
from dataclasses import replace
from pathlib import Path

import pytest

from office_live import com, config, registry, safety, wd_common, xl_common
from office_live.errors import ToolError


class Coll:
    def __init__(self, items):
        self.items = items

    @property
    def Count(self):  # noqa: N802 — имя как у COM
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]


class Book:
    def __init__(self, name, path="", autosave=False, sheets=("Data", "Other")):
        self.Name = name
        self.Path = path
        self.FullName = (path + "\\" + name) if path else name
        self.AutoSaveOn = autosave
        self.Worksheets = self.Sheets = Coll([Sheet(n) for n in sheets])
        self.ActiveSheet = self.Worksheets(1)


class Sheet:
    def __init__(self, name):
        self.Name = name


class App:
    def __init__(self, books, doc_mode=False):
        self.EnableEvents = True
        coll = Coll(books)
        if doc_mode:
            self.Documents = coll
            self.ActiveDocument = books[0]
        else:
            self.Workbooks = coll
            self.ActiveWorkbook = books[0]


def with_settings(monkeypatch, **env):
    s = config.load(env)
    monkeypatch.setattr(config, "SETTINGS", s)
    return s


@pytest.fixture
def excel(monkeypatch):
    books = [Book("Report.xlsx", r"D:\Docs"), Book("Report2.xlsx", r"D:\Docs"), Book("Budget.xlsx", r"C:\Secret")]
    app = App(books)
    monkeypatch.setattr(com, "apps", lambda kind, launch=False: [app])
    monkeypatch.setitem(com._CTX, "cleanup", [])
    return app, books


def as_write(monkeypatch):
    monkeypatch.setitem(com._CTX, "kind", "write")


# ------------------------------------------------------------------ строгий режим и имена


def test_whitespace_name_counts_as_missing_in_strict_write(excel, monkeypatch):
    with_settings(monkeypatch, OFFICE_LIVE_STRICT_TARGET="1")
    as_write(monkeypatch)
    for blank in ("", "   ", "\t"):
        with pytest.raises(ToolError, match="name is required"):
            xl_common.pick_workbook(blank)


def test_strict_write_demands_exact_name_not_substring(excel, monkeypatch):
    with_settings(monkeypatch, OFFICE_LIVE_STRICT_TARGET="1")
    as_write(monkeypatch)
    _, wb = xl_common.pick_workbook("Report.xlsx")
    assert wb.Name == "Report.xlsx"
    with pytest.raises(ToolError, match="not found"):
        xl_common.pick_workbook("Budget")  # подстрока допустима только вне строгого режима
    _, wb = xl_common.pick_workbook(r"d:\docs\report2.xlsx")  # полный путь без учёта регистра
    assert wb.Name == "Report2.xlsx"


def test_strict_write_demands_sheet_name(excel, monkeypatch):
    _, books = excel
    with_settings(monkeypatch, OFFICE_LIVE_STRICT_TARGET="1")
    as_write(monkeypatch)
    with pytest.raises(ToolError, match="Sheet name is required"):
        xl_common.pick_sheet(books[0], "")
    with pytest.raises(ToolError, match="not found"):
        xl_common.pick_sheet(books[0], "Dat")
    assert xl_common.pick_sheet(books[0], "data").Name == "Data"


def test_relaxed_mode_still_allows_unique_substring_and_active(excel, monkeypatch):
    with_settings(monkeypatch)
    as_write(monkeypatch)
    assert xl_common.pick_workbook("Budget")[1].Name == "Budget.xlsx"
    assert xl_common.pick_workbook("")[1].Name == "Report.xlsx"  # активная
    with pytest.raises(ToolError, match="ambiguous"):
        xl_common.pick_workbook("Report")


def test_reads_are_not_affected_by_strict_mode(excel, monkeypatch):
    with_settings(monkeypatch, OFFICE_LIVE_STRICT_TARGET="1")
    monkeypatch.setitem(com._CTX, "kind", "read")
    assert xl_common.pick_workbook("")[1].Name == "Report.xlsx"


# ------------------------------------------------------------------ зона доступа для УЖЕ открытых документов


def test_doc_allowed_semantics(monkeypatch, tmp_path):
    inside = tmp_path / "ok"
    inside.mkdir()
    s = config.load({"OFFICE_LIVE_ALLOWED_DIRS": str(inside)})
    assert safety.doc_allowed(str(inside / "a.xlsx"), s)
    assert not safety.doc_allowed(str(tmp_path / "a.xlsx"), s)
    assert safety.doc_allowed("", s)  # несохранённый документ без пути — наш
    assert safety.doc_allowed(str(tmp_path / "anything.xlsx"), config.load({}))  # без ограничений — всё


def test_open_documents_outside_allowed_dirs_are_invisible(excel, monkeypatch):
    with_settings(monkeypatch, OFFICE_LIVE_ALLOWED_DIRS=r"D:\Docs")
    monkeypatch.setitem(com._CTX, "kind", "read")
    assert [wb.Name for _, wb in xl_common.all_workbooks()] == ["Report.xlsx", "Report2.xlsx"]
    with pytest.raises(ToolError, match="not found"):
        xl_common.pick_workbook("Budget.xlsx")


def test_active_forbidden_workbook_is_refused(monkeypatch):
    app = App([Book("Budget.xlsx", r"C:\Secret"), Book("Ok.xlsx", r"D:\Docs")])
    monkeypatch.setattr(com, "apps", lambda kind, launch=False: [app])
    with_settings(monkeypatch, OFFICE_LIVE_ALLOWED_DIRS=r"D:\Docs")
    with pytest.raises(ToolError, match="active workbook is outside"):
        xl_common.pick_workbook("")


def test_word_pickers_follow_the_same_rules(monkeypatch):
    docs = [Book("a.docx", r"D:\Docs"), Book("b.docx", r"C:\Secret")]
    app = App(docs, doc_mode=True)
    monkeypatch.setattr(com, "apps", lambda kind, launch=False: [app])
    with_settings(monkeypatch, OFFICE_LIVE_ALLOWED_DIRS=r"D:\Docs", OFFICE_LIVE_STRICT_TARGET="1")
    as_write(monkeypatch)
    assert [d.Name for _, d in wd_common.all_documents()] == ["a.docx"]
    with pytest.raises(ToolError, match="not found"):
        wd_common.pick_document("b.docx")
    with pytest.raises(ToolError, match="not found"):
        wd_common.pick_document("a")  # подстрока в строгом режиме
    with pytest.raises(ToolError, match="name is required"):
        wd_common.pick_document("  ")
    assert wd_common.pick_document("a.docx")[1].Name == "a.docx"


# ------------------------------------------------------------------ AutoSave и события


def test_autosave_blocks_writes_but_not_reads_or_saving(monkeypatch):
    app = App([Book("Cloud.xlsx", r"C:\OneDrive", autosave=True)])
    monkeypatch.setattr(com, "apps", lambda kind, launch=False: [app])
    with_settings(monkeypatch)
    as_write(monkeypatch)
    with pytest.raises(ToolError, match="AutoSave is ON"):
        xl_common.pick_workbook("Cloud.xlsx")
    assert xl_common.pick_workbook("Cloud.xlsx", allow_autosave=True)[1].Name == "Cloud.xlsx"  # save/close/save_as
    monkeypatch.setitem(com._CTX, "kind", "read")
    assert xl_common.pick_workbook("Cloud.xlsx")[1].Name == "Cloud.xlsx"
    with_settings(monkeypatch, OFFICE_LIVE_AUTOSAVE="allow")
    as_write(monkeypatch)
    assert xl_common.pick_workbook("Cloud.xlsx")[1].Name == "Cloud.xlsx"


def test_autosave_guard_for_word(monkeypatch):
    app = App([Book("Cloud.docx", r"C:\OneDrive", autosave=True)], doc_mode=True)
    monkeypatch.setattr(com, "apps", lambda kind, launch=False: [app])
    with_settings(monkeypatch)
    as_write(monkeypatch)
    with pytest.raises(ToolError, match="AutoSave is ON"):
        wd_common.pick_document("Cloud.docx")
    assert wd_common.pick_document("Cloud.docx", allow_autosave=True)[1].Name == "Cloud.docx"


def test_events_are_switched_off_for_writes_and_restored(excel, monkeypatch):
    app, _ = excel
    with_settings(monkeypatch)
    as_write(monkeypatch)
    xl_common.pick_workbook("Report.xlsx")
    assert app.EnableEvents is False  # чужие макросы не срабатывают от наших правок
    for undo in com._CTX["cleanup"]:
        undo()
    assert app.EnableEvents is True


def test_events_are_left_alone_for_reads_and_when_opted_in(excel, monkeypatch):
    app, _ = excel
    with_settings(monkeypatch)
    monkeypatch.setitem(com._CTX, "kind", "read")
    xl_common.pick_workbook("Report.xlsx")
    assert app.EnableEvents is True
    with_settings(monkeypatch, OFFICE_LIVE_ENABLE_EVENTS="1")
    as_write(monkeypatch)
    xl_common.pick_workbook("Report.xlsx")
    assert app.EnableEvents is True


def test_run_com_runs_cleanups_even_when_the_tool_fails():
    seen = []

    def tool():
        com.add_cleanup(lambda: seen.append("restored"))
        com.note_target("workbook:X")
        raise KeyError("boom")

    meta: dict = {}
    with pytest.raises(ToolError):
        com.run_com(tool, meta=meta)
    assert seen == ["restored"] and meta["targets"] == ["workbook:X"]


# ------------------------------------------------------------------ конфигурация


def test_autosave_setting_is_validated():
    assert config.load({}).autosave == "block"
    assert config.load({"OFFICE_LIVE_AUTOSAVE": "ALLOW"}).autosave == "allow"
    with pytest.raises(config.ConfigError):
        config.load({"OFFICE_LIVE_AUTOSAVE": "maybe"})


# ------------------------------------------------------------------ readonly по действиям и журнал


class StubServer:
    def __init__(self):
        self.tools = {}

    def add_tool(self, fn, name=None, **kw):
        self.tools[name] = (fn, kw)


@pytest.fixture
def stub_registry(monkeypatch):
    stub = StubServer()
    monkeypatch.setattr(registry, "mcp", stub)
    yield stub
    for name in list(registry.CATALOG):
        if name.startswith("zz_"):
            registry.CATALOG.pop(name)


def test_multi_action_tool_is_available_in_readonly_only_for_read_actions(monkeypatch, stub_registry):
    with_settings(monkeypatch, OFFICE_LIVE_MODE="readonly")

    @registry.office_tool("excel_core", "write", read_actions=("list", "get"))
    def zz_manage(action: str = "list", name: str = ""):
        """Manage things (test double used by the unit tests)."""
        return {"action": action, "name": name}

    wrapped, _ = stub_registry.tools["zz_manage"]
    assert wrapped("list")["action"] == "list" and wrapped(action="get")["action"] == "get"
    assert wrapped()["action"] == "list"  # значение по умолчанию тоже «чтение»
    for bad in ("delete", "add", "export"):
        with pytest.raises(ToolError, match="Read-only mode"):
            wrapped(action=bad)


def test_pure_writer_is_not_registered_in_readonly(monkeypatch, stub_registry):
    with_settings(monkeypatch, OFFICE_LIVE_MODE="readonly")

    @registry.office_tool("excel_core", "write")
    def zz_writer(x: int = 1):
        """Writes something (test double used by the unit tests)."""
        return x

    assert "zz_writer" not in stub_registry.tools and registry.CATALOG["zz_writer"].registered is False


def test_toolset_restriction_still_hides_read_actions(monkeypatch, stub_registry):
    with_settings(monkeypatch, OFFICE_LIVE_MODE="readonly", OFFICE_LIVE_TOOLSETS="word_core")

    @registry.office_tool("excel_core", "write", read_actions=("list",))
    def zz_hidden(action: str = "list"):
        """Excel tool hidden by the toolset filter (test double used by the unit tests)."""
        return action

    assert "zz_hidden" not in stub_registry.tools


def test_audit_log_writes_start_then_end_with_targets_and_hides_long_text(monkeypatch, stub_registry, tmp_path):
    log = tmp_path / "audit.jsonl"
    s = config.load({"OFFICE_LIVE_AUDIT_LOG": str(log)})
    monkeypatch.setattr(config, "SETTINGS", s)

    @registry.office_tool("excel_core", "write")
    def zz_audited(text: str = "", fail: bool = False):
        """Audited writer (test double used by the unit tests)."""
        com.note_target("workbook:Book1")
        if fail:
            raise ToolError("nope")
        return "ok"

    wrapped, _ = stub_registry.tools["zz_audited"]
    secret = "CONFIDENTIAL " * 20
    assert wrapped(text=secret) == "ok"
    with pytest.raises(ToolError):
        wrapped(fail=True)
    rows = [json.loads(line) for line in log.read_text(encoding="utf-8").splitlines()]
    assert [(r["phase"], r.get("ok")) for r in rows] == [("start", None), ("end", True), ("start", None), ("end", False)]
    assert rows[1]["targets"] == ["workbook:Book1"] and "duration_ms" in rows[1]
    assert "CONFIDENTIAL" not in log.read_text(encoding="utf-8")  # содержимое документа в журнал не попадает
    s2 = replace(s, audit_content=True)
    monkeypatch.setattr(config, "SETTINGS", s2)
    wrapped(text="short text")
    assert "short text" in log.read_text(encoding="utf-8")


def test_audit_log_unwritable_warns_once_but_does_not_break_tools(monkeypatch, stub_registry, tmp_path, capsys):
    blocker = tmp_path / "file.txt"
    blocker.write_text("x")
    s = config.load({"OFFICE_LIVE_AUDIT_LOG": str(blocker / "sub" / "audit.jsonl")})  # родитель — файл: писать нельзя
    monkeypatch.setattr(config, "SETTINGS", s)
    monkeypatch.setattr(registry, "_audit_warned", False)

    @registry.office_tool("excel_core", "write")
    def zz_quiet():
        """Writer with a broken audit log (test double used by the unit tests)."""
        return 1

    wrapped, _ = stub_registry.tools["zz_quiet"]
    assert wrapped() == 1 and wrapped() == 1
    assert capsys.readouterr().err.count("audit log is not writable") == 1


def test_destructive_hint_override():
    ann = registry._annotations("write", "t", None, destructive=True)
    assert ann.destructive_hint is True and ann.read_only_hint is False
    assert registry._annotations("read", "t", None).destructive_hint is None
    assert Path  # keep import used


# ------------------------------------------------------------------ AttributeError от занятого COM-объекта


def _busy_error():
    import pywintypes

    return pywintypes.com_error(-2147418111, "busy", (0, None, "busy", None, 0, 0), None)  # RPC_E_CALL_REJECTED


class FlakyDispatch:
    """Ведёт себя как CDispatch: пока приложение «занято», любой атрибут превращается в AttributeError."""

    def __init__(self, busy_attempts):
        self.attempts = 0
        self.busy_attempts = busy_attempts
        self._oleobj_ = self

    def GetIDsOfNames(self, lcid, name):  # noqa: N802
        if self.attempts <= self.busy_attempts:
            raise _busy_error()
        return 1

    def __getattr__(self, name):
        if name.startswith("_"):
            raise AttributeError(name)
        self.attempts += 1
        if self.attempts <= self.busy_attempts:
            raise AttributeError(f"Excel.Application.{name}")  # так pywin32 прячет com_error
        return "value"


def test_proxy_retries_when_busy_hides_behind_attribute_error(monkeypatch):
    monkeypatch.setattr(com.time, "sleep", lambda s: None)
    flaky = FlakyDispatch(busy_attempts=3)
    assert com.Proxy(flaky).Workbooks == "value" and flaky.attempts == 4


def test_proxy_keeps_real_attribute_errors():
    class Plain:
        _oleobj_ = None

        def __getattr__(self, name):
            raise AttributeError(name)

    plain = Plain()

    class Ole:
        def GetIDsOfNames(self, lcid, name):  # noqa: N802
            import pywintypes

            raise pywintypes.com_error(-2147352570, "unknown name", None, None)  # DISP_E_UNKNOWNNAME

    plain._oleobj_ = Ole()
    with pytest.raises(AttributeError):
        com.Proxy(plain).NoSuchProperty
