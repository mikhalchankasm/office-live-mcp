"""Регрессии повторного аудита: обход зоны через выделение, проверки до изменений, срезы, события, порядок освобождения COM,
установщик (TOML без tomllib, generic --path, заголовки, копии), DOCX-инспектор, аннотации и журнал. Office не нужен."""

import asyncio
import json
import sys
import zipfile
from pathlib import Path
from types import SimpleNamespace

import pytest
import pywintypes

from office_live import bridge, cli, com, config, excel_core, excel_format, excel_pivot, inspect_file, registry, word_core, word_layout, xl_common
from office_live.errors import ToolError


def with_settings(monkeypatch, **env):
    monkeypatch.setattr(config, "SETTINGS", config.load(env))


def com_error(hresult=-2147352567):
    return pywintypes.com_error(hresult, "x", (0, None, "x", None, 0, 0), None)


class Coll:
    def __init__(self, items):
        self.items = items

    @property
    def Count(self):  # noqa: N802
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]


# ======================================================================== 1. выделение и «активная» книга вне зоны


class Sel:
    """Выделение, к которому нельзя обращаться, если книга вне зоны."""

    def __init__(self, forbidden=False):
        self._forbidden = forbidden
        self.Row = self.Column = 1
        self.Rows = SimpleNamespace(Count=1)
        self.Columns = SimpleNamespace(Count=1)
        self.Areas = SimpleNamespace(Count=1)

    @property
    def Address(self):  # noqa: N802
        assert not self._forbidden, "selection of a forbidden workbook was touched"
        return "$A$1"

    @property
    def Value(self):  # noqa: N802
        assert not self._forbidden, "selection of a forbidden workbook was read"
        return "PRIVATE"


def book(name, path, sheets=("S",)):
    return SimpleNamespace(Name=name, Path=path, FullName=(path + "\\" + name) if path else name, Saved=True, ReadOnly=False, AutoSaveOn=False, sheets=sheets)


def test_get_selection_does_not_read_a_workbook_outside_the_allowed_dirs(monkeypatch):
    secret = book("Secret.xlsx", r"C:\Secret")
    app = SimpleNamespace(ActiveWorkbook=secret, ActiveSheet=SimpleNamespace(Name="Hidden sheet"), Selection=Sel(forbidden=True), Workbooks=Coll([secret]))
    monkeypatch.setattr(com, "apps", lambda kind, launch=False: [app])
    with_settings(monkeypatch, OFFICE_LIVE_ALLOWED_DIRS=r"D:\Docs")
    out = excel_core.excel_get_selection()
    assert out["workbook"] is None and out["selection"] is None and "outside the folders" in out["note"]
    assert "Secret" not in json.dumps(out) and "Hidden sheet" not in json.dumps(out)
    app.ActiveWorkbook = book("Ok.xlsx", r"D:\Docs")
    app.Selection = Sel()
    out = excel_core.excel_get_selection()
    assert out["workbook"] == "Ok.xlsx" and out["selection"] == "A1" and out["values"] == [["PRIVATE"]]  # разрешённая книга читается как раньше


def test_list_workbooks_does_not_name_a_forbidden_active_workbook_or_sheet(monkeypatch):
    secret, ok = book("Secret.xlsx", r"C:\Secret"), book("Ok.xlsx", r"D:\Docs")
    app = SimpleNamespace(ActiveWorkbook=secret, ActiveSheet=SimpleNamespace(Name="Hidden sheet"), Workbooks=Coll([secret, ok]))
    monkeypatch.setattr(com, "apps", lambda kind, launch=False: [app])
    monkeypatch.setattr(excel_core, "_is_visible", lambda wb: True)
    monkeypatch.setattr(excel_core, "_sheet_list", lambda wb: list(wb.sheets))
    with_settings(monkeypatch, OFFICE_LIVE_ALLOWED_DIRS=r"D:\Docs")
    out = excel_core.excel_list_workbooks()
    assert [b["name"] for b in out["workbooks"]] == ["Ok.xlsx"]
    assert out["active_workbook"] is None and out["active_sheet"] is None
    assert "Secret" not in json.dumps(out) and "Hidden sheet" not in json.dumps(out)


def test_list_documents_does_not_name_a_forbidden_active_document(monkeypatch):
    secret = SimpleNamespace(Name="Secret.docx", Path=r"C:\Secret", FullName=r"C:\Secret\Secret.docx", Saved=True, ReadOnly=False, Tables=SimpleNamespace(Count=0))
    ok = SimpleNamespace(Name="Ok.docx", Path=r"D:\Docs", FullName=r"D:\Docs\Ok.docx", Saved=True, ReadOnly=False, Tables=SimpleNamespace(Count=0))
    app = SimpleNamespace(Documents=Coll([secret, ok]), ActiveDocument=secret)
    monkeypatch.setattr(com, "apps", lambda kind, launch=False: [app])
    monkeypatch.setattr(word_core, "para_count", lambda d: 1)
    with_settings(monkeypatch, OFFICE_LIVE_ALLOWED_DIRS=r"D:\Docs")
    out = word_core.word_list_documents()
    assert [d["name"] for d in out["documents"]] == ["Ok.docx"] and out["active_document"] is None


# ======================================================================== 2. слияние: имена уникальны после каждого переименования


def test_merge_names_stay_unique_after_renaming():
    used, renamed = set(), []
    names = [bridge._unique_stem(stem, i, used, renamed) for i, stem in enumerate(["x_003", "x", "x"], start=1)]
    assert len(set(n.lower() for n in names)) == 3, names  # раньше: x_003, x, x_003
    assert names[0] == "x_003" and names[1] == "x" and names[2] not in ("x", "x_003")
    assert renamed and renamed[0]["row"] == 3


def test_merge_names_ignore_case_and_chain_collisions():
    used, renamed = set(), []
    names = [bridge._unique_stem(s, i, used, renamed) for i, s in enumerate(["A", "a", "A", "a_002", "A_003"], start=1)]
    assert len({n.lower() for n in names}) == 5, names


# ======================================================================== 3. Excel: проверки до изменений


class Recorder:
    """Объект, который помнит порядок присваиваний своих свойств (в общий журнал)."""

    def __init__(self, log, prefix=""):
        object.__setattr__(self, "_log", log)
        object.__setattr__(self, "_prefix", prefix)

    def __setattr__(self, name, value):
        self._log.append(f"{self._prefix}{name}")
        object.__setattr__(self, name, value)

    def __getattr__(self, name):
        if name in ("Font", "Interior"):
            sub = Recorder(self._log, name + ".")
            object.__setattr__(self, name, sub)
            return sub
        raise AttributeError(name)


def make_range(log):
    r = Recorder(log)
    for k, v in (("Row", 1), ("Column", 1), ("Address", "$A$1:$B$2")):
        object.__setattr__(r, k, v)
    object.__setattr__(r, "Rows", SimpleNamespace(Count=2))
    object.__setattr__(r, "Columns", SimpleNamespace(Count=2))
    return r


@pytest.fixture
def fmt(monkeypatch):
    log: list = []
    rng = make_range(log)
    wb = SimpleNamespace(Name="B.xlsx")
    monkeypatch.setattr(excel_format, "pick_workbook", lambda name: (SimpleNamespace(), wb))
    monkeypatch.setattr(excel_format, "get_range", lambda wb_, sheet, cells, **kw: (SimpleNamespace(Name="S"), rng))
    with_settings(monkeypatch)
    return log


@pytest.mark.parametrize(
    "kwargs",
    [
        {"bold": True, "font_color": "definitely-not-a-color"},
        {"bold": True, "fill_color": "#12"},
        {"italic": True, "indent": 999},
        {"bold": True, "text_rotation": 180},
        {"bold": True, "horizontal_alignment": "sideways"},
        {"bold": True, "borders": "bogus"},
        {"bold": True, "borders": "all", "border_style": "wavy"},
        {"bold": True, "borders": "all", "border_color": "nope"},
        {"bold": True, "font_size": 0},
    ],
)
def test_format_range_rejects_bad_arguments_before_changing_anything(fmt, kwargs):
    with pytest.raises(ToolError):
        excel_format.excel_format_range("B.xlsx", "S", "A1:B2", **kwargs)
    assert fmt == [], f"bad arguments left partial changes: {fmt}"


def test_format_range_applies_everything_when_arguments_are_fine(fmt):
    out = excel_format.excel_format_range("B.xlsx", "S", "A1:B2", bold=True, font_color="#C00000", fill_color="none", indent=2)
    assert out["applied"] == ["bold=True", "font_color=#C00000", "fill_color=none", "indent=2"]
    assert "Font.Bold" in fmt and "IndentLevel" in fmt


def test_format_range_needs_at_least_one_parameter(fmt):
    with pytest.raises(ToolError, match="Nothing to apply"):
        excel_format.excel_format_range("B.xlsx", "S", "A1:B2")


class FakeSheetForFilter:
    def __init__(self):
        self.log = []
        self._mode = True

    @property
    def AutoFilterMode(self):  # noqa: N802
        return self._mode

    @AutoFilterMode.setter
    def AutoFilterMode(self, value):  # noqa: N802
        self.log.append(("AutoFilterMode", value))
        self._mode = value

    def Range(self, addr):  # noqa: N802
        return SimpleNamespace(Value=(("name", "qty"),))


def test_filter_range_validates_every_condition_before_removing_the_old_filter(monkeypatch):
    ws = FakeSheetForFilter()
    rng = SimpleNamespace(Row=1, Column=1, Rows=SimpleNamespace(Count=5), Columns=SimpleNamespace(Count=2), applied=[])
    rng.AutoFilter = lambda *a: rng.applied.append(a)
    monkeypatch.setattr(excel_core, "pick_workbook", lambda name: (SimpleNamespace(), SimpleNamespace(Name="B.xlsx")))
    monkeypatch.setattr(excel_core, "get_range", lambda wb, sheet, cells, **kw: (ws, rng))
    with_settings(monkeypatch)
    with pytest.raises(ToolError, match="Unknown operator"):
        excel_core.excel_filter_range("B.xlsx", "S", "A1:B5", filters=[{"column": "name", "criteria": "x"}, {"column": "qty", "operator": "typo", "criteria": "1"}])
    assert ws.log == [] and ws.AutoFilterMode is True and rng.applied == []  # прежний фильтр цел
    with pytest.raises(ToolError, match="not found"):
        excel_core.excel_filter_range("B.xlsx", "S", "A1:B5", filters=[{"column": "nope", "criteria": "x"}])
    assert ws.log == []


def test_sort_range_validates_keys_before_clearing_old_sort_fields(monkeypatch):
    cleared = []
    ws = SimpleNamespace(Range=lambda addr: SimpleNamespace(Value=(("name", "qty"),)), Sort=SimpleNamespace(SortFields=SimpleNamespace(Clear=lambda: cleared.append(1))))
    rng = SimpleNamespace(Row=1, Column=1, Rows=SimpleNamespace(Count=4), Columns=SimpleNamespace(Count=2))
    monkeypatch.setattr(excel_core, "pick_workbook", lambda name: (SimpleNamespace(), SimpleNamespace(Name="B.xlsx")))
    monkeypatch.setattr(excel_core, "get_range", lambda wb, sheet, cells, **kw: (ws, rng))
    with_settings(monkeypatch)
    with pytest.raises(ToolError, match="order must be"):
        excel_core.excel_sort_range("B.xlsx", "S", "A1:B4", keys=[{"column": "name"}, {"column": "qty", "order": "sideways"}])
    assert cleared == []


# ======================================================================== 4. сводная: фильтр и срезы


def test_pivot_filter_parses_operands_before_clearing_existing_filters(monkeypatch):
    cleared = []
    pf = SimpleNamespace(ClearAllFilters=lambda: cleared.append(1), PivotFilters=SimpleNamespace(Add2=lambda *a: None))
    monkeypatch.setattr(excel_pivot, "pick_workbook", lambda name: (SimpleNamespace(), SimpleNamespace(Name="B.xlsx")))
    monkeypatch.setattr(excel_pivot, "find_pivot", lambda wb, name: (SimpleNamespace(), SimpleNamespace()))
    monkeypatch.setattr(excel_pivot, "_field", lambda pt, name: pf)
    monkeypatch.setattr(excel_pivot, "_safe_orient", lambda f: 1)
    monkeypatch.setattr(excel_pivot, "_data_field", lambda pt, caption: object())
    with_settings(monkeypatch)
    with pytest.raises(ToolError, match="must be a number"):
        excel_pivot.excel_pivot_filter("B.xlsx", field="Region", action="value", operator="greater", value1="abc", data_field="Sum of X")
    assert cleared == []  # раньше прежние фильтры сбрасывались, а потом выяснялось, что операнд не число


class FakeSlicer:
    def __init__(self, name, caption=None):
        self.Name, self.Caption = name, caption or name
        self.Shape = SimpleNamespace(Left=0, Top=0, Width=10, Height=10)
        self.deleted = False
        self.NumberOfColumns = 1
        self.Style = ""

    def Delete(self):  # noqa: N802
        self.deleted = True


class FakeCache:
    def __init__(self, name, field, slicers):
        self.Name, self.SourceName = name, field
        self.slicers = slicers
        self.Slicers = Coll(slicers)
        self.deleted = False

    def Delete(self):  # noqa: N802
        self.deleted = True


@pytest.fixture
def slicer_env(monkeypatch):
    first, second = FakeSlicer("First", "First"), FakeSlicer("Second", "Second caption")
    cache = FakeCache("Slicer_Region", "Region", [first, second])
    wb = SimpleNamespace(Name="B.xlsx", SlicerCaches=Coll([cache]))
    monkeypatch.setattr(excel_pivot, "pick_workbook", lambda name: (SimpleNamespace(), wb))
    monkeypatch.setattr(excel_pivot, "_slicer_info", lambda sc, max_items=100: {"cache": sc.Name})
    with_settings(monkeypatch)
    return cache, first, second


def test_slicer_move_changes_the_requested_slicer_not_the_first(slicer_env):
    _, first, second = slicer_env
    excel_pivot.excel_manage_slicers("B.xlsx", action="move", slicer="Second", width=150, height=180, caption="New caption")
    assert second.Caption == "New caption" and second.Shape.Width == 150
    assert first.Caption == "First" and first.Shape.Width == 10  # раньше менялся sc.Slicers(1)
    excel_pivot.excel_manage_slicers("B.xlsx", action="move", slicer="New caption", width=99, height=99)  # и по (новой) подписи
    assert second.Shape.Width == 99 and first.Shape.Width == 10


def test_slicer_move_by_cache_name_with_several_slicers_asks_which_one(slicer_env):
    with pytest.raises(ToolError, match="Several slicers share"):
        excel_pivot.excel_manage_slicers("B.xlsx", action="move", slicer="Region", width=50, height=50)


def test_slicer_delete_by_name_removes_only_that_slicer(slicer_env):
    cache, first, second = slicer_env
    out = excel_pivot.excel_manage_slicers("B.xlsx", action="delete", slicer="Second")
    assert second.deleted and not first.deleted and not cache.deleted and out["deleted"] == "Second"


def test_slicer_delete_by_cache_name_or_for_the_last_slicer_removes_the_cache(slicer_env):
    cache, first, second = slicer_env
    out = excel_pivot.excel_manage_slicers("B.xlsx", action="delete", slicer="Slicer_Region")
    assert cache.deleted and set(out["deleted_slicers"]) == {"First", "Second"}
    lone = FakeCache("Slicer_X", "X", [FakeSlicer("Only")])
    wb = SimpleNamespace(Name="B.xlsx", SlicerCaches=Coll([lone]))
    excel_pivot.pick_workbook = lambda name: (SimpleNamespace(), wb)
    excel_pivot.excel_manage_slicers("B.xlsx", action="delete", slicer="Only")
    assert lone.deleted


# ======================================================================== 5. события Excel и порядок освобождения COM


class EventApp:
    def __init__(self, fail=False):
        self.fail = fail
        self._events = True

    @property
    def EnableEvents(self):  # noqa: N802
        return self._events

    @EnableEvents.setter
    def EnableEvents(self, value):  # noqa: N802
        if self.fail and value is False:
            raise com_error()
        self._events = value


def test_write_is_refused_when_events_cannot_be_switched_off(monkeypatch):
    with_settings(monkeypatch)
    monkeypatch.setitem(com._CTX, "cleanup", [])
    app = EventApp(fail=True)
    with pytest.raises(ToolError, match="Could not switch off Excel events"):
        xl_common.suspend_events(app)
    assert app.EnableEvents is True and com._CTX["cleanup"] == []  # молча продолжать с включёнными событиями нельзя


def test_events_failure_can_be_accepted_explicitly(monkeypatch):
    with_settings(monkeypatch, OFFICE_LIVE_ENABLE_EVENTS="1")
    xl_common.suspend_events(EventApp(fail=True))  # не трогает события вовсе


def test_cleanup_callbacks_are_released_before_com_is_uninitialised(monkeypatch):
    events = []

    class Holder:
        def __del__(self):
            events.append("released")

    def tool():
        holder = Holder()
        com.add_cleanup(lambda h=holder: events.append("cleanup"))

    monkeypatch.setattr(com.pythoncom, "CoUninitialize", lambda: events.append("CoUninitialize"))
    com.run_com(tool)
    assert events == ["cleanup", "released", "CoUninitialize"]


# ======================================================================== 6. колонтитулы: свойство страницы для связанных разделов


class FakeHF:
    def __init__(self, linked):
        self.LinkToPrevious = linked
        self.text_sets = 0
        self.Range = SimpleNamespace(
            ParagraphFormat=SimpleNamespace(Alignment=0), Font=SimpleNamespace(Size=0), End=1, SetRange=lambda a, b: None,
            InsertAfter=lambda t: None, Fields=SimpleNamespace(Add=lambda *a: None),
        )
        outer = self

        class R(SimpleNamespace):
            pass

        def set_text(value):
            outer.text_sets += 1

        self.Range = R(ParagraphFormat=SimpleNamespace(Alignment=0), Font=SimpleNamespace(Size=0), End=1, SetRange=lambda a, b: None, InsertAfter=lambda t: None, Fields=SimpleNamespace(Add=lambda *a: None))
        type(self.Range).Text = property(lambda s: "", lambda s, v: set_text(v))


class FakeSection:
    def __init__(self, linked):
        self.PageSetup = SimpleNamespace(DifferentFirstPageHeaderFooter=False)
        self.hf = FakeHF(linked)

    def Headers(self, t):  # noqa: N802
        return self.hf

    Footers = Headers


def test_headers_footers_apply_page_property_to_linked_sections_too(monkeypatch):
    sections = [FakeSection(linked=False), FakeSection(linked=True)]
    doc = SimpleNamespace(Name="D.docx", Sections=Coll(sections))
    monkeypatch.setattr(word_layout, "pick_document", lambda name: (SimpleNamespace(), doc))
    with_settings(monkeypatch)
    out = word_layout.word_headers_footers("D.docx", action="set", kind="footer", section=0, text="X", different_first_page=True)
    assert [s.PageSetup.DifferentFirstPageHeaderFooter for s in sections] == [True, True]  # раньше: [True, False]
    assert out["sections_changed"] == [1] and sections[0].hf.text_sets >= 1 and sections[1].hf.text_sets == 0  # общий текст правится один раз


def test_headers_footers_refuse_single_linked_section_without_break_link(monkeypatch):
    sections = [FakeSection(linked=False), FakeSection(linked=True)]
    doc = SimpleNamespace(Name="D.docx", Sections=Coll(sections))
    monkeypatch.setattr(word_layout, "pick_document", lambda name: (SimpleNamespace(), doc))
    with_settings(monkeypatch)
    with pytest.raises(ToolError, match="linked to the previous section"):
        word_layout.word_headers_footers("D.docx", action="set", kind="footer", section=2, text="X", different_first_page=True)
    assert sections[1].PageSetup.DifferentFirstPageHeaderFooter is False and sections[1].hf.text_sets == 0


# ======================================================================== 7. аннотации и журнал


def test_render_declares_that_it_is_not_read_only_and_merge_is_destructive():
    from office_live.app import mcp

    tools = {t.name: t for t in asyncio.run(mcp.list_tools())}
    assert tools["excel_render_range_image"].annotations.read_only_hint is False  # на мгновение ставит диаграмму на лист
    assert tools["excel_merge_cells"].annotations.destructive_hint is True
    assert tools["excel_read_range"].annotations.read_only_hint is True


class StubServer:
    def __init__(self):
        self.tools = {}

    def add_tool(self, fn, name=None, **kw):
        self.tools[name] = fn


def test_read_action_that_writes_a_file_is_audited(monkeypatch, tmp_path):
    stub = StubServer()
    monkeypatch.setattr(registry, "mcp", stub)
    log = tmp_path / "audit.jsonl"
    monkeypatch.setattr(config, "SETTINGS", config.load({"OFFICE_LIVE_AUDIT_LOG": str(log)}))

    @registry.office_tool("excel_analysis", "write", read_actions=("list", "export_image"), file_args=("export_path",))
    def zz_charts(action: str = "list", export_path: str = ""):
        """Chart tool double (used by the unit tests only)."""
        return action

    try:
        stub.tools["zz_charts"](action="export_image")  # без записи на диск — не аудируется
        assert not log.exists()
        stub.tools["zz_charts"](action="export_image", export_path=r"D:\x.png")  # с записью — аудируется
        rows = [json.loads(line) for line in log.read_text(encoding="utf-8").splitlines()]
        assert [r["phase"] for r in rows] == ["start", "end"]
    finally:
        registry.CATALOG.pop("zz_charts", None)


def test_live_fixtures_only_claim_files_of_their_own_run():
    from tests.live.conftest import RUN_TAG, _is_ours

    assert _is_ours("a.xlsx", "b.xlsx", rf"C:\Temp\{RUN_TAG}_x") is True
    assert _is_ours("a.xlsx", "b.xlsx", r"C:\Temp\ol_pytest_old_run_dir") is False  # чужой каталог и прежнего запуска — не наш
    assert _is_ours("Книга1", "Книга1", "") is True and _is_ours("x", "y", "") is False


# ======================================================================== 8. установщик


@pytest.fixture
def fake_home(tmp_path, monkeypatch):
    home = tmp_path / "home"
    (home / "AppData" / "Roaming").mkdir(parents=True)
    monkeypatch.setattr(Path, "home", classmethod(lambda cls: home))
    monkeypatch.setenv("APPDATA", str(home / "AppData" / "Roaming"))
    monkeypatch.setattr(cli.shutil, "which", lambda name: None)
    return home


READONLY_TOML = (
    'model = "x"\n\n[mcp_servers.office-live]\ncommand = "old"\nargs = ["a"]\n'
    '[mcp_servers.office-live.env]\nOFFICE_LIVE_MODE = "readonly"\nOFFICE_LIVE_ALLOWED_DIRS = "D:/Docs"\n'
)


def codex_cfg(home, text=READONLY_TOML):
    cfg = home / ".codex" / "config.toml"
    cfg.parent.mkdir(parents=True, exist_ok=True)
    cfg.write_text(text, encoding="utf-8")
    return cfg


def test_codex_reinstall_keeps_restrictions_without_tomllib(fake_home, monkeypatch):
    """Python 3.10 не имеет tomllib: прежний код возвращал {} и затирал ограничения."""
    monkeypatch.setitem(sys.modules, "tomllib", None)  # import tomllib -> ImportError
    cfg = codex_cfg(fake_home)
    assert cli.run("setup", ["--yes", "--clients", "codex"]) == 0
    text = cfg.read_text(encoding="utf-8")
    assert 'OFFICE_LIVE_MODE = "readonly"' in text and "OFFICE_LIVE_ALLOWED_DIRS" in text and 'command = "old"' not in text


@pytest.mark.parametrize(
    "text,expected",
    [
        ('[mcp_servers.office-live.env]\nA = "1"\nB = \'2\'  # c\n', {"A": "1", "B": "2"}),
        ('[mcp_servers."office-live"]\ncommand = "x"\nenv = { A = "1", "B" = \'two\' }\n', {"A": "1", "B": "two"}),
        ('[mcp_servers.office-live]\ncommand = "x"\nenv.A = "1"\n', {"A": "1"}),
        ('[other]\nA = "9"\n[mcp_servers.office-live]\ncommand = "x"\n', {}),
        ('[mcp_servers.office-live.env]\nP = "C:\\\\dir\\\\x"\n', {"P": "C:\\dir\\x"}),
    ],
)
def test_toml_fallback_parser_reads_our_entry(text, expected):
    assert cli._toml_env_fallback(text) == expected


def test_unreadable_entry_is_not_silently_overwritten(fake_home, monkeypatch, capsys):
    monkeypatch.setitem(sys.modules, "tomllib", None)
    cfg = codex_cfg(fake_home, '[mcp_servers.office-live]\ncommand = "old"\n[mcp_servers.office-live.env]\nOFFICE_LIVE_MODE = 5\n')
    before = cfg.read_text(encoding="utf-8")
    assert cli.run("setup", ["--yes", "--clients", "codex"]) == 1  # настройки прочитать нельзя -> отказ, файл не тронут
    assert cfg.read_text(encoding="utf-8") == before
    assert "Cannot read the settings" in capsys.readouterr().err
    assert cli.run("setup", ["--yes", "--clients", "codex", "--reset"]) == 0  # явное согласие потерять настройки


def test_invalid_toml_with_tomllib_is_also_refused(fake_home):
    pytest.importorskip("tomllib")
    cfg = codex_cfg(fake_home, '[mcp_servers.office-live]\ncommand = "old\n')
    before = cfg.read_text(encoding="utf-8")
    assert cli.run("setup", ["--yes", "--clients", "codex"]) == 1 and cfg.read_text(encoding="utf-8") == before


def test_generic_client_writes_to_the_path_it_was_given(fake_home, tmp_path):
    target = tmp_path / "x.json"
    assert cli.config_cmd(["generic", "--write", "--path", str(target), "--quiet"]) == 0
    assert "office-live" in json.loads(target.read_text(encoding="utf-8"))["mcpServers"]
    assert cli.config_cmd(["generic", "--write", "--quiet"]) == 1  # без пути писать некуда


def test_header_with_quoted_first_component_is_replaced_not_duplicated(fake_home):
    cfg = codex_cfg(fake_home, 'model = "x"\n\n["mcp_servers"."office-live"]\ncommand = "old"\n["mcp_servers"."office-live".env]\nA = "1"\n\n[mcp_servers.other]\ncommand = "keep"\n')
    assert cli.run("setup", ["--yes", "--clients", "codex"]) == 0
    text = cfg.read_text(encoding="utf-8")
    assert 'command = "old"' not in text and "[mcp_servers.other]" in text
    tomllib = pytest.importorskip("tomllib")
    servers = tomllib.loads(text)["mcp_servers"]  # корректный TOML без повторного объявления таблицы
    assert set(servers) == {"office-live", "other"} and servers["office-live"]["env"] == {"A": "1"}


def test_two_backups_in_the_same_second_do_not_overwrite_each_other(tmp_path, monkeypatch):
    monkeypatch.setattr(cli.time, "strftime", lambda fmt: "20260101-000000")
    f = tmp_path / "cfg.json"
    f.write_text("one", encoding="utf-8")
    first = cli._backup(f)
    f.write_text("two", encoding="utf-8")
    second = cli._backup(f)
    assert first != second and first.read_text(encoding="utf-8") == "one" and second.read_text(encoding="utf-8") == "two"


# ======================================================================== 9. DOCX-инспектор: любой префикс пространства имён


def test_docx_inspector_does_not_depend_on_the_namespace_prefix(tmp_path):
    body = (
        '<x:document xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:body>'
        '<x:p><x:pPr><x:pStyle x:val="Heading1"/></x:pPr><x:r><x:t>Title {{Name}}</x:t></x:r></x:p>'
        '<x:p><x:r><x:t>Body</x:t></x:r></x:p>'
        '<x:tbl><x:tr><x:tc><x:p><x:r><x:t>cell</x:t></x:r></x:p></x:tc></x:tr></x:tbl>'
        '<x:p><x:ins x:id="1"><x:r><x:t>added</x:t></x:r></x:ins><x:fldSimple x:instr="PAGE"/></x:p><x:sectPr/></x:body></x:document>'
    )
    path = tmp_path / "a.docx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", body)
    info = inspect_file.inspect_docx(str(path))
    assert info["paragraphs"] == 4 and info["tables"] == 1 and info["sections"] == 1 and info["fields"] == 1
    assert info["tracked_changes"] == {"insertions": 1, "deletions": 0}
    assert info["peek"][0] == {"style": "Heading1", "text": "Title {{Name}}"} and info["placeholders"] == ["Name"]
    assert info["paragraph_styles_used"] == {"Heading1": 1}
