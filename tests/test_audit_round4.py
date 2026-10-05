"""Регрессии четвёртой проверки (обработка ошибок): нечитаемые срезы, AutoSave при «Office занят», многострочные
строки TOML, битый исходный config.toml. Office не нужен."""

import re
import sys
from types import SimpleNamespace

import pytest
import pywintypes

from office_live import cli, config, excel_format, excel_pivot
from office_live.errors import AppBusyError, ToolError


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


# ======================================================================== 1. ошибка перечисления срезов


class BrokenSlicers:
    @property
    def Count(self):  # noqa: N802
        raise com_error()


class Cache:
    def __init__(self, name, field, slicers=None):
        self.Name, self.SourceName = name, field
        self.Slicers = BrokenSlicers() if slicers is None else Coll(slicers)
        self.deleted = False

    def Delete(self):  # noqa: N802
        self.deleted = True


def book(monkeypatch, *caches):
    wb = SimpleNamespace(Name="B.xlsx", SlicerCaches=Coll(list(caches)))
    monkeypatch.setattr(excel_pivot, "pick_workbook", lambda name: (SimpleNamespace(), wb))
    monkeypatch.setattr(excel_pivot, "_slicer_info", lambda sc, max_items=100: {"cache": sc.Name})
    monkeypatch.setattr(config, "SETTINGS", config.load({}))


@pytest.mark.parametrize("action", ["delete", "delete_cache", "move"])
def test_unlistable_slicers_stop_the_change_instead_of_counting_as_none(monkeypatch, action):
    """Раньше ошибка чтения Slicers.Count давала [] — и delete удалял кэш как «пустой», с ok=true."""
    cache = Cache("Slicer_Region", "Region")
    book(monkeypatch, cache)
    with pytest.raises(ToolError, match="could not list the slicers"):
        excel_pivot.excel_manage_slicers("B.xlsx", action=action, slicer="Slicer_Region", caption="X")
    assert not cache.deleted


def test_an_unreadable_cache_blocks_guessing_by_field(monkeypatch):
    """Искомый срез 'Region' может лежать в непрочитанном кэше — нельзя подменять его кэшем другого поля 'Region'."""
    other = SimpleNamespace(Name="Only", Caption="Only", Delete=lambda: None)
    broken, readable = Cache("Slicer_A", "A"), Cache("Slicer_B", "Region", [other])
    book(monkeypatch, broken, readable)
    with pytest.raises(ToolError, match="cannot be resolved safely"):
        excel_pivot.excel_manage_slicers("B.xlsx", action="delete", slicer="Region")
    assert not (broken.deleted or readable.deleted)


def test_a_slicer_name_in_a_readable_cache_still_works_next_to_a_broken_one(monkeypatch):
    target = SimpleNamespace(Name="Mine", Caption="Mine", deleted=False)
    target.Delete = lambda: setattr(target, "deleted", True)
    sibling = SimpleNamespace(Name="Sib", Caption="Sib")
    book(monkeypatch, Cache("Slicer_A", "A"), Cache("Slicer_B", "B", [target, sibling]))
    out = excel_pivot.excel_manage_slicers("B.xlsx", action="delete", slicer="Mine")
    assert target.deleted and out["deleted"] == "Mine"


# ======================================================================== 2. AutoSave при «Office занят» во время восстановления


def test_autosave_comes_back_even_when_restoring_the_view_reports_busy(monkeypatch):
    wb = SimpleNamespace(Name="B.xlsx", AutoSaveOn=True, Windows=lambda i: SimpleNamespace(WindowState=-4137))
    app = SimpleNamespace(WindowState=-4143, Goto=lambda *a: None)
    wb.Activate = lambda: None
    ws = SimpleNamespace(Name="S", Activate=lambda: None)
    rng = SimpleNamespace(Width=100, Height=50, Cells=lambda r, c: None)
    monkeypatch.setattr(excel_format, "bounds", lambda rng: (1, 1, 2, 2))
    monkeypatch.setattr(excel_format, "count_nonempty", lambda app, rng: 1)
    monkeypatch.setattr(excel_format, "suspend_events", lambda app: None)
    monkeypatch.setattr(excel_format.time, "sleep", lambda s: None)
    monkeypatch.setattr(excel_format, "_copy_picture_png", lambda *a: b"png")
    monkeypatch.setattr(excel_format, "_is_blank_png", lambda data: False)
    monkeypatch.setattr(excel_format, "_remember_view", lambda app: {"wb": None, "ws": None, "sel": None})

    def busy(app, state):
        raise AppBusyError("Office is busy")

    monkeypatch.setattr(excel_format, "_restore_view", busy)
    with pytest.raises(AppBusyError):
        excel_format.render_range_png(app, wb, ws, rng)
    assert wb.AutoSaveOn is True


# ======================================================================== 3, 4. установщик и config.toml


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


INSTRUCTIONS = '''developer_instructions = """
Example of a server entry:
[mcp_servers.office-live]
command = "python"
[example_end]
Keep this text.
"""
'''


def test_table_headers_inside_a_multiline_string_are_left_alone(fake_home):
    tomllib = pytest.importorskip("tomllib")
    cfg = codex_cfg(fake_home, 'model = "x"\n' + INSTRUCTIONS)
    assert cli.run("setup", ["--yes", "--clients", "codex"]) == 0
    data = tomllib.loads(cfg.read_text(encoding="utf-8"))
    assert data["developer_instructions"] == tomllib.loads(INSTRUCTIONS)["developer_instructions"]
    assert "office-live" in data["mcp_servers"]
    assert cli.run("setup", ["--remove", "--yes", "--clients", "codex"]) == 0  # и удаление не трогает пример
    data = tomllib.loads(cfg.read_text(encoding="utf-8"))
    assert "Keep this text." in data["developer_instructions"] and "mcp_servers" not in data


def test_brackets_inside_multiline_arrays_are_not_table_headers():
    text = '[mcp_servers.office-live]\nmatrix = [\n[1, 2],\n[3]\n]\n[mcp_servers.z]\nq = 1\n'
    out = cli._strip_toml_tables(text)
    assert out == "[mcp_servers.z]\nq = 1\n"  # прежде строка '[3]' считалась заголовком и обрывки массива оставались


def test_a_rewrite_that_changes_other_settings_is_refused(fake_home, monkeypatch, capsys):
    """Даже если построчная замена ошибётся (здесь — прежний наивный вариант), результат сверяется целиком."""
    pytest.importorskip("tomllib")
    header = re.compile(r"^\s*\[\[?[^\[\]=\n]*\]\]?\s*(?:#.*)?$")

    def naive(text):
        out, skipping = [], False
        for line in text.splitlines(keepends=True):
            if header.match(line):
                skipping = bool(cli._OURS.match(line))
            if not skipping:
                out.append(line)
        return "".join(out)

    monkeypatch.setattr(cli, "_strip_toml_tables", naive)
    text = 'model = "x"\n' + INSTRUCTIONS
    cfg = codex_cfg(fake_home, text)
    assert cli.run("setup", ["--yes", "--clients", "codex"]) == 1
    assert cfg.read_text(encoding="utf-8") == text and "change other settings" in capsys.readouterr().err


def test_invalid_existing_toml_is_not_made_worse(fake_home, capsys):
    """Битый исходный файл без нашей записи: прежде проверка результата пропускалась и установка «успешно» его меняла."""
    pytest.importorskip("tomllib")
    text = 'model = "unterminated\n'
    cfg = codex_cfg(fake_home, text)
    assert cli.run("setup", ["--yes", "--clients", "codex"]) == 1
    assert cfg.read_text(encoding="utf-8") == text and "Не удалось записать конфигурацию" in capsys.readouterr().err


def test_without_a_parser_a_file_that_mentions_the_server_is_refused(fake_home, monkeypatch):
    """Без парсера нельзя отличить запись от упоминания в тексте — отказ безопаснее, чем угадывать."""
    monkeypatch.setitem(sys.modules, "tomllib", None)
    monkeypatch.setitem(sys.modules, "tomli", None)
    text = 'model = "x"\n' + INSTRUCTIONS
    cfg = codex_cfg(fake_home, text)
    assert cli.run("setup", ["--yes", "--clients", "codex"]) == 1
    assert cfg.read_text(encoding="utf-8") == text
