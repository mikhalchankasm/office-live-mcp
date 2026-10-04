"""Регрессии третьего аудита: фильтр сводной в readonly, TOML без tomllib, удаление и правка срезов, очистка живых тестов,
аудит экспорта PNG, AutoSave при сбое снимка. Office не нужен."""

import json
import sys
from types import SimpleNamespace

import pytest
import pywintypes

from office_live import cli, config, excel_format, excel_pivot, registry
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


class StubServer:
    def __init__(self):
        self.tools = {}

    def add_tool(self, fn, name=None, **kw):
        self.tools[name] = fn


@pytest.fixture
def reregister(monkeypatch):
    """Регистрирует НАСТОЯЩИЙ инструмент заново с его же объявлением из CATALOG (группа, вид, read_actions, file_args)
    под заданными настройками; COM-контекст подменён прямым вызовом."""
    stub = StubServer()
    monkeypatch.setattr(registry, "mcp", stub)
    monkeypatch.setattr(registry, "run_com", lambda fn, args, kwargs, **kw: fn(*args, **kwargs))
    saved = dict(registry.CATALOG)

    def make(fn, **env):
        with_settings(monkeypatch, **env)
        info = saved[fn.__name__]
        registry.office_tool(info.group, info.kind, read_actions=info.read_actions or None, file_args=info.file_args or None)(fn)
        return stub.tools[fn.__name__]

    yield make
    registry.CATALOG.clear()
    registry.CATALOG.update(saved)


# ======================================================================== 1. фильтр сводной в readonly


class Item:
    def __init__(self, name, visible=True):
        self.Name, self._visible, self.writes = name, visible, 0

    @property
    def Visible(self):  # noqa: N802
        return self._visible

    @Visible.setter
    def Visible(self, value):  # noqa: N802
        self.writes += 1
        self._visible = value


@pytest.fixture
def pivot_env(monkeypatch):
    items = [Item("A"), Item("B", visible=False), Item("C")]
    pf = SimpleNamespace(Name="Region", Orientation=1, PivotItems=lambda: Coll(items))
    wb = SimpleNamespace(Name="B.xlsx")
    monkeypatch.setattr(excel_pivot, "pick_workbook", lambda name: (SimpleNamespace(), wb))
    monkeypatch.setattr(excel_pivot, "find_pivot", lambda wb, name: (SimpleNamespace(Name="S"), SimpleNamespace(Name="PT")))
    monkeypatch.setattr(excel_pivot, "_field", lambda pt, name: pf)
    monkeypatch.setattr(excel_pivot, "_summary", lambda ws, pt: {})
    return items


def test_pivot_item_filter_is_refused_in_readonly(pivot_env, reregister):
    tool = reregister(excel_pivot.excel_pivot_filter, OFFICE_LIVE_MODE="readonly")
    with pytest.raises(ToolError, match="Read-only mode"):
        tool(workbook="B.xlsx", field="Region", action="items", mode="only", items=["A"])
    with pytest.raises(ToolError, match="Read-only mode"):
        tool(workbook="B.xlsx", field="Region", items=["A"])  # action по умолчанию — изменяющий 'items'
    assert all(i.writes == 0 for i in pivot_env)


def test_pivot_list_action_reads_items_without_changing_them(pivot_env, reregister):
    tool = reregister(excel_pivot.excel_pivot_filter, OFFICE_LIVE_MODE="readonly")
    out = tool(workbook="B.xlsx", field="Region", action="list")
    assert out["items"] == [{"name": "A", "visible": True}, {"name": "B", "visible": False}, {"name": "C", "visible": True}]
    assert out["hidden_items"] == 1 and all(i.writes == 0 for i in pivot_env)


def test_only_list_is_a_read_action_of_the_pivot_filter():
    assert registry.CATALOG["excel_pivot_filter"].read_actions == ("list",)


# ======================================================================== 2. TOML без tomllib


TRIPLE_QUOTED = (
    '[mcp_servers.office-live]\ncommand = "old"\n'
    'env = { OFFICE_LIVE_MODE = """readonly""", OFFICE_LIVE_ALLOWED_DIRS = """D:/Docs""" }\n'
)


@pytest.fixture
def fake_home(tmp_path, monkeypatch):
    home = tmp_path / "home"
    home.mkdir()
    monkeypatch.setattr(cli.Path, "home", classmethod(lambda cls: home))
    monkeypatch.setenv("APPDATA", str(home / "AppData" / "Roaming"))
    monkeypatch.setattr(cli.shutil, "which", lambda name: None)
    return home


def codex_cfg(home, text):
    cfg = home / ".codex" / "config.toml"
    cfg.parent.mkdir(parents=True, exist_ok=True)
    cfg.write_text(text, encoding="utf-8")
    return cfg


def test_without_any_toml_parser_an_existing_entry_is_refused_not_emptied(fake_home, monkeypatch, capsys):
    """Прежний самодельный разбор превращал многострочные строки в "" и снимал ограничения с кодом 0."""
    monkeypatch.setitem(sys.modules, "tomllib", None)
    monkeypatch.setitem(sys.modules, "tomli", None)
    cfg = codex_cfg(fake_home, TRIPLE_QUOTED)
    assert cli.run("setup", ["--yes", "--clients", "codex"]) == 1
    assert cfg.read_text(encoding="utf-8") == TRIPLE_QUOTED
    assert "tomli" in capsys.readouterr().err


def test_without_any_toml_parser_a_file_without_our_entry_is_still_written(fake_home, monkeypatch):
    monkeypatch.setitem(sys.modules, "tomllib", None)
    monkeypatch.setitem(sys.modules, "tomli", None)
    cfg = codex_cfg(fake_home, 'model = "x"\n')
    assert cli.run("setup", ["--yes", "--clients", "codex"]) == 0
    assert "[mcp_servers.office-live]" in cfg.read_text(encoding="utf-8")


def test_tomli_backport_reads_any_valid_toml_and_keeps_the_restrictions(fake_home, monkeypatch):
    tomllib = pytest.importorskip("tomllib")
    monkeypatch.setitem(sys.modules, "tomllib", None)
    monkeypatch.setitem(sys.modules, "tomli", tomllib)  # бэкпорт с тем же API
    cfg = codex_cfg(fake_home, TRIPLE_QUOTED)
    assert cli.run("setup", ["--yes", "--clients", "codex"]) == 0
    env = tomllib.loads(cfg.read_text(encoding="utf-8"))["mcp_servers"]["office-live"]["env"]
    assert env == {"OFFICE_LIVE_MODE": "readonly", "OFFICE_LIVE_ALLOWED_DIRS": "D:/Docs"}


def test_entry_that_cannot_be_replaced_line_by_line_is_left_untouched(fake_home, capsys):
    """Запись как inline-таблица под [mcp_servers]: построчная замена дала бы дубликат ключа — отказ без записи."""
    pytest.importorskip("tomllib")
    text = '[mcp_servers]\noffice-live = { command = "old", env = { OFFICE_LIVE_MODE = "readonly" } }\n'
    cfg = codex_cfg(fake_home, text)
    assert cli.run("setup", ["--yes", "--clients", "codex"]) == 1
    assert cfg.read_text(encoding="utf-8") == text and "Refusing to rewrite" in capsys.readouterr().err
    assert not list(cfg.parent.glob("*.bak*"))  # и копию не плодим


# ======================================================================== 3, 7. срезы


class FakeSlicer:
    def __init__(self, name, caption=None):
        self.Name, self.Caption = name, caption or name
        self.Shape = SimpleNamespace(Left=0, Top=0, Width=300, Height=400)
        self.deleted = False
        self.NumberOfColumns = 1
        self.Style = ""

    def Delete(self):  # noqa: N802
        self.deleted = True


class FakeCache:
    def __init__(self, name, field, slicers):
        self.Name, self.SourceName = name, field
        self.Slicers = Coll(slicers)
        self.deleted = False

    def Delete(self):  # noqa: N802
        self.deleted = True


def slicer_book(monkeypatch, *caches):
    wb = SimpleNamespace(Name="B.xlsx", SlicerCaches=Coll(list(caches)))
    monkeypatch.setattr(excel_pivot, "pick_workbook", lambda name: (SimpleNamespace(), wb))
    monkeypatch.setattr(excel_pivot, "_slicer_info", lambda sc, max_items=100: {"cache": sc.Name})
    with_settings(monkeypatch)
    return wb


@pytest.fixture
def region(monkeypatch):
    """Поле 'Region', а срезы называются 'Region' и 'Region2' — имя поля совпадает с именем одного из срезов."""
    a, b = FakeSlicer("Region"), FakeSlicer("Region2")
    cache = FakeCache("Slicer_Region", "Region", [a, b])
    slicer_book(monkeypatch, cache)
    return cache, a, b


def test_slicer_name_wins_over_a_field_of_the_same_name(region):
    cache, a, b = region
    out = excel_pivot.excel_manage_slicers("B.xlsx", action="delete", slicer="Region")
    assert a.deleted and not b.deleted and not cache.deleted and out["deleted"] == "Region"


def test_delete_of_a_cache_with_several_slicers_is_refused_before_deleting(region):
    cache, a, b = region
    with pytest.raises(ToolError, match="delete_cache"):
        excel_pivot.excel_manage_slicers("B.xlsx", action="delete", slicer="Slicer_Region")
    assert not (cache.deleted or a.deleted or b.deleted)


def test_delete_cache_removes_the_cache_with_all_slicers(region):
    cache, _, _ = region
    out = excel_pivot.excel_manage_slicers("B.xlsx", action="delete_cache", slicer="Region2")
    assert cache.deleted and out["deleted_slicers"] == ["Region", "Region2"]


def test_ambiguous_caption_or_field_is_an_error_not_the_first_match(monkeypatch):
    c1 = FakeCache("Slicer_A", "Year", [FakeSlicer("S1", "Same")])
    c2 = FakeCache("Slicer_B", "Year", [FakeSlicer("S2", "Same")])
    slicer_book(monkeypatch, c1, c2)
    with pytest.raises(ToolError, match="Several slicers have the caption"):
        excel_pivot.excel_manage_slicers("B.xlsx", action="delete", slicer="Same")
    with pytest.raises(ToolError, match="Several slicer caches filter the field"):
        excel_pivot.excel_manage_slicers("B.xlsx", action="delete_cache", slicer="Year")
    assert not (c1.deleted or c2.deleted)


def test_move_changes_only_what_was_passed(region):
    _, a, _ = region
    excel_pivot.excel_manage_slicers("B.xlsx", action="move", slicer="Region", caption="New")
    assert a.Caption == "New" and (a.Shape.Width, a.Shape.Height) == (300, 400)  # раньше сбрасывалось в 150 x 180
    excel_pivot.excel_manage_slicers("B.xlsx", action="move", slicer="Region", width=120)
    assert (a.Shape.Width, a.Shape.Height) == (120, 400)


def test_move_can_return_to_a_single_column(region):
    _, a, _ = region
    excel_pivot.excel_manage_slicers("B.xlsx", action="move", slicer="Region", columns=3)
    assert a.NumberOfColumns == 3
    excel_pivot.excel_manage_slicers("B.xlsx", action="move", slicer="Region", columns=1)
    assert a.NumberOfColumns == 1


def test_bad_slicer_settings_are_refused_before_any_change(region):
    _, a, _ = region
    with pytest.raises(ToolError, match="columns"):
        excel_pivot.excel_manage_slicers("B.xlsx", action="move", slicer="Region", caption="X", columns=0)
    with pytest.raises(ToolError, match="action must be"):
        excel_pivot.excel_manage_slicers("B.xlsx", action="remove", slicer="Region")
    assert a.Caption == "Region"


# ======================================================================== 4. очистка живых тестов


class FakeSrv:
    """Сервер для фикстур: несколько экземпляров Excel/Word, закрытия записываются."""

    def __init__(self, books, names_by_book=None, ambiguous=()):
        self.books, self.names, self.ambiguous, self.closed = books, names_by_book or {}, set(ambiguous), []

    def call(self, tool, **kw):
        from tests.live.mcpclient import ToolFailed

        if tool in ("excel_list_workbooks", "word_list_documents"):
            return {"workbooks": self.books, "documents": self.books}
        target = kw.get("workbook") or kw.get("document")
        if target in self.ambiguous:
            raise ToolFailed(tool, f"'{target}' exists in several instances")
        if tool == "excel_manage_names":
            return {"names": [{"name": n} for n in self.names.get(target, [])]}
        if tool == "word_document_properties":
            return {"properties": {"Comments": self.names.get(target, [""])[0]}}
        if tool in ("excel_close_workbook", "word_close_document"):
            self.closed.append(target)
            return {"ok": True}
        raise AssertionError(tool)


def test_cleanup_does_not_close_a_users_book_with_the_original_name():
    """Тест создал «Книга1», сохранил её как own.xlsx в папке запуска; в другом экземпляре открыта «Книга1» пользователя."""
    from tests.live.conftest import RUN_TAG, close_own_documents, close_own_workbooks

    own_dir = rf"C:\Temp\{RUN_TAG}_x"
    srv = FakeSrv([{"name": "own.xlsx", "path": own_dir}, {"name": "Книга1", "path": None}], names_by_book={"Книга1": ["Users_name"]})
    close_own_workbooks(srv, "Книга1", f"{RUN_TAG}_m")
    assert srv.closed == [own_dir + "\\own.xlsx"]
    srv = FakeSrv([{"name": "Документ1", "path": None}], names_by_book={"Документ1": ["user comment"]})
    close_own_documents(srv, "Документ1", f"{RUN_TAG}_m")
    assert srv.closed == []


def test_cleanup_closes_its_own_marked_draft_and_skips_ambiguous_names():
    from tests.live.conftest import RUN_TAG, close_own_workbooks

    marker = f"{RUN_TAG}_m"
    srv = FakeSrv([{"name": "Книга1", "path": None}], names_by_book={"Книга1": [marker]})
    close_own_workbooks(srv, "Книга1", marker)
    assert srv.closed == ["Книга1"]
    srv = FakeSrv([{"name": "Книга1", "path": None}, {"name": "Книга1", "path": None}], ambiguous={"Книга1"})
    close_own_workbooks(srv, "Книга1", marker)  # своя и чужая с одним именем: лучше оставить свою, чем закрыть чужую
    assert srv.closed == []
    srv = FakeSrv([{"name": "a.xlsx", "path": r"C:\Temp\ol_pytest_old_run_dir"}])
    close_own_workbooks(srv, "Книга1", marker)  # каталог прежнего запуска — не наш
    assert srv.closed == []


# ======================================================================== 5. аудит экспорта PNG настоящим инструментом


def test_real_chart_tool_declares_its_export_path():
    assert registry.CATALOG["excel_manage_charts"].file_args == ("export_path",)


def test_real_chart_export_to_a_file_is_audited(monkeypatch, reregister, tmp_path):
    from office_live import excel_analysis

    log = tmp_path / "audit.jsonl"
    monkeypatch.setattr(excel_analysis, "pick_workbook", lambda name: (_ for _ in ()).throw(ToolError("stop after the audit start")))
    tool = reregister(excel_analysis.excel_manage_charts, OFFICE_LIVE_AUDIT_LOG=str(log))
    with pytest.raises(ToolError):
        tool(workbook="B.xlsx", action="export_image", chart_name="C1")
    assert not log.exists()  # картинка без записи на диск — чтение, не аудируется
    with pytest.raises(ToolError):
        tool(workbook="B.xlsx", action="export_image", chart_name="C1", export_path=str(tmp_path / "c.png"))
    rows = [json.loads(line) for line in log.read_text(encoding="utf-8").splitlines()]
    assert [r["phase"] for r in rows] == ["start", "end"] and rows[0]["tool"] == "excel_manage_charts"


# ======================================================================== 6. AutoSave при сбое подготовки снимка


class AutoSaveBook:
    def __init__(self):
        self.Name = "B.xlsx"
        self.AutoSaveOn = True


class BrokenViewApp:
    @property
    def ActiveWorkbook(self):  # noqa: N802
        raise com_error()


def test_autosave_is_restored_when_remembering_the_view_fails(monkeypatch):
    monkeypatch.setattr(excel_format, "bounds", lambda rng: (1, 1, 2, 2))
    monkeypatch.setattr(excel_format, "count_nonempty", lambda app, rng: 1)
    monkeypatch.setattr(excel_format, "suspend_events", lambda app: None)
    wb = AutoSaveBook()
    rng = SimpleNamespace(Width=100, Height=50)
    with pytest.raises(pywintypes.com_error):
        excel_format.render_range_png(BrokenViewApp(), wb, SimpleNamespace(Name="S"), rng)
    assert wb.AutoSaveOn is True
