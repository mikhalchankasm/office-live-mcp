"""Настройки, безопасность, COM-слой (на подделках, без Office), реестр инструментов, VBA-декомпрессор."""

import asyncio
import os
import re
from pathlib import Path

import pytest
import pywintypes

from office_live import com, config, safety, vba
from office_live.errors import AppBusyError, ToolError
from tests.history_fakes import fake_office  # noqa: F401

ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture
def import_office(office, monkeypatch):
    from tests.history_fakes import Range

    monkeypatch.setattr(Range, "HasArray", False, raising=False)
    return office


def test_import_acceptance(import_office):
    import json
    from office_live import undo

    o = import_office
    path = o.tmp / "table.csv"
    path.write_text("Код;Название\n" + "007;Пример\n" * 336, encoding="cp1251")
    result = o.call("excel_import_table", path=str(path), workbook=o.wb.Name, text_columns=["Код"])
    assert (result["rows"], result["columns"], result["sheet"]) == (337, 2, "Data")
    assert (result["delimiter"], result["encoding"]) == (";", "cp1251")
    assert o.ws.Range("A2").Value2 == "007" and o.ws.Range("A2").NumberFormat == "@"
    assert o.ws.Range("A1:B1").Font.Bold and o.ws.Range("A1:B1").WrapText
    assert o.ws.view.FreezePanes and o.ws.view.SplitRow == 1
    assert len(json.dumps(result).encode()) <= 2048
    stack = undo.STACKS[undo.stack_key("workbook", o.excel, o.wb)]
    assert len(stack) == 1 and len(stack[0].backup_sheets) == 1
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1:B337").Value2 == ((None, None),) * 337
    assert not o.ws.view.FreezePanes


@pytest.mark.parametrize("ext,encoding,content,expected", [
    ("csv", "utf-8-sig", "code,name\n007,тест", [["code", "name"], ["007", "тест"]]),
    ("tsv", "utf-16", "code\tname\n007\tтест", [["code", "name"], ["007", "тест"]]),
    ("json", "utf-8", '[{"a":"007"},{"b":2,"a":"008"}]', [["a", "b"], ["007", None], ["008", 2]]),
    ("json", "utf-8", '[["a","b"],["007"]]', [["a", "b"], ["007", None]]),
])
def test_import_formats(import_office, ext, encoding, content, expected):
    o = import_office
    path = o.tmp / ("table." + ext)
    path.write_text(content, encoding=encoding)
    result = o.call("excel_import_table", path=str(path), workbook=o.wb.Name, text_columns=[1], layout=False)
    assert result["sample"] == expected and result["encoding"] == encoding
    assert o.ws.Range("A2").Value2 == "007"


@pytest.mark.parametrize("problem", ["occupied", "protected", "merge", "array", "spill", "autosave", "nested", "limit", "header", "snapshot", "structure"])
def test_import_refuses_before_changes(import_office, monkeypatch, problem):
    import copy
    from office_live import undo
    from tests.history_fakes import Range

    o = import_office
    path = o.tmp / "table.json"
    path.write_text('[["code"],["007"]]', encoding="utf-8")
    args = {"path": str(path), "workbook": o.wb.Name}
    if problem == "occupied":
        o.ws.Range("A2").Formula = '=""'
    elif problem == "protected":
        o.ws.ProtectContents = True
    elif problem == "merge":
        o.ws.Range("A1").MergeCells = True
    elif problem == "array":
        monkeypatch.setattr(Range, "HasArray", True)
    elif problem == "spill":
        o.ws.spills = {(2, 1)}
    elif problem == "autosave":
        o.wb.AutoSaveOn = True
    elif problem == "nested":
        path.write_text('[{"code":{"x":1}}]', encoding="utf-8")
    elif problem == "limit":
        args["max_rows"] = 1
    elif problem == "header":
        args["text_columns"] = ["missing"]
    elif problem == "snapshot":
        monkeypatch.setattr(undo, "_snapshots", lambda *a, **k: (_ for _ in ()).throw(ValueError("snapshot failed")))
    else:
        o.wb.ProtectStructure = True
        args["sheet"] = "New"
    before = copy.deepcopy((o.ws.data, o.ws.formats))
    with pytest.raises(ToolError) as error:
        o.call("excel_import_table", **args)
    if problem == "nested":
        assert '$[0]["code"]' in str(error.value)
    if problem == "occupied":
        assert "A2" in str(error.value)
    assert (o.ws.data, o.ws.formats) == before and not undo.STACKS and o.wb.Sheets.Count == 1


def test_import_created_sheet_and_book(import_office):
    o = import_office
    path = o.tmp / "arbitrary.csv"
    path.write_text("code\n007", encoding="utf-8")
    result = o.call("excel_import_table", path=str(path), workbook=o.wb.Name, sheet="Imported", value_mode="text")
    assert result["undo"] == "available" and o.wb.Sheets.Count == 2
    o.call("office_undo", file=o.wb.Name)
    assert o.wb.Sheets.Count == 1
    result = o.call("excel_import_table", path=str(path), value_mode="text")
    assert result["workbook"] != o.wb.Name and result["sheet"] == "Data"
    assert "new workbook" in result["undo"] and "not available" in result["undo"]


def test_import_inactive_and_partial_undo(import_office, monkeypatch):
    from office_live import excel_core

    o = import_office
    other = o.wb.Worksheets.Add()
    other.Name = "Other"
    o.ws.Activate()
    monkeypatch.setattr(other, "Activate", lambda: pytest.fail("must not activate import target"))
    path = o.tmp / "table.csv"
    path.write_text("code\n007", encoding="utf-8")
    result = o.call("excel_import_table", path=str(path), workbook=o.wb.Name, sheet="Other")
    assert "Freeze skipped" in " ".join(result["warnings"]) and o.excel.ActiveSheet is o.ws
    o.call("office_undo", file=o.wb.Name)
    original = excel_core._write_values
    def fail(*args, **kwargs):
        original(*args, **kwargs)
        raise ValueError("interrupted after write")
    monkeypatch.setattr(excel_core, "_write_values", fail)
    with pytest.raises(ToolError, match="office_undo will restore"):
        o.call("excel_import_table", path=str(path), workbook=o.wb.Name, text_columns=[1])
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("A1:A2").Value2 == ((None,), (None,)) and o.ws.Range("A1").NumberFormat == "General"


def test_import_compact_overwrite_and_single_write(import_office, monkeypatch):
    import json
    from tests.history_fakes import Range

    o = import_office
    path = o.tmp / "table.json"
    path.write_text(json.dumps([["x" * 1000] * 20] * 4), encoding="utf-8")
    o.ws.Range("B3").Value2 = "before"
    writes = []
    original = Range.Value  # the import writes through write_range's path (Range.Value)
    def write(rng, value):
        writes.append(rng.Address)
        original.fset(rng, value)
    monkeypatch.setattr(Range, "Value", property(original.fget, write))
    result = o.call("excel_import_table", path=str(path), workbook=o.wb.Name, start_cell="B3", overwrite=True)
    assert writes == ["B3:U6"] and len(json.dumps(result).encode()) <= 2048
    assert len(result["sample"]) == 3 and any("truncated" in warning for warning in result["warnings"])
    o.call("office_undo", file=o.wb.Name)
    assert o.ws.Range("B3").Value2 == "before"


@pytest.mark.parametrize("content,options", [
    ('[[1]]', {"start_cell": "XFD1048577"}), ('[[1,2]]', {"start_cell": "XFD1"}),
    ('[[1]]', {"text_columns": [True]}), ('[[1]]', {"max_rows": 200001}),
    ('[[NaN]]', {}), ('[{"a":[]}]', {}), ('[]', {}),
    ('[[' + ','.join(['1'] * 201) + ']]', {}), ('[["' + 'x' * 32768 + '"]]', {}),
], ids=["start-out-of-sheet", "start-too-wide", "bad-text-columns", "too-many-rows", "nan", "nested-json",
        "empty", "too-many-columns", "cell-over-32767"])  # short ids: a 32 KB id overflows PYTEST_CURRENT_TEST
def test_import_invalid_file_never_creates_book(import_office, monkeypatch, content, options):
    from office_live import excel_core

    path = import_office.tmp / "bad.json"
    path.write_text(content, encoding="utf-8")
    monkeypatch.setattr(com, "primary_app", lambda *a, **k: pytest.fail("must validate before Office"))
    with pytest.raises(ToolError):
        excel_core.excel_import_table(str(path), **options)


def test_import_headerless_json_and_allowed_path(import_office, monkeypatch):
    from dataclasses import replace
    from office_live import excel_core

    o = import_office
    path = o.tmp / "table.json"
    path.write_text('[{"code":"007"},{"value":2}]', encoding="utf-8")
    result = o.call("excel_import_table", path=str(path), workbook=o.wb.Name, header=False, text_columns=["code"])
    assert result["sample"] == [["007", None], [None, 2]] and not o.ws.view.FreezePanes
    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, allowed_dirs=(o.tmp / "allowed",)))
    monkeypatch.setattr(com, "primary_app", lambda *a, **k: pytest.fail("must reject outside allowed dirs first"))
    with pytest.raises(ToolError, match="outside"):
        excel_core.excel_import_table(str(path))


# ------------------------------------------------------------------ config


def test_defaults():
    s = config.load({})
    assert s.mode == "full" and not s.readonly and not s.allow_eval
    assert s.groups == frozenset(config.GROUPS)
    assert s.busy_timeout == 20 and s.audit_log is None


def test_readonly_hides_every_writing_tool():
    s = config.load({"OFFICE_LIVE_MODE": "readonly"})
    assert s.enabled("excel_core", "read") and s.enabled("excel_core", "ui") and s.enabled("word_core", "open")
    for kind in ("write", "destructive", "save"):
        assert not s.enabled("excel_core", kind)


def test_unknown_mode_fails_safe():
    assert config.load({"OFFICE_LIVE_MODE": "banana"}).readonly


def test_toolset_presets_and_groups():
    assert config.load({"OFFICE_LIVE_TOOLSETS": "core"}).groups == frozenset({"excel_core", "word_core", "bridge", "window"})
    assert config.load({"OFFICE_LIVE_TOOLSETS": "word"}).groups == frozenset({"word_core", "word_tables", "word_layout", "window"})
    assert config.load({"OFFICE_LIVE_TOOLSETS": "excel_format, bridge"}).groups == frozenset({"excel_format", "bridge"})
    assert config.load({"OFFICE_LIVE_TOOLSETS": "nonsense, excel_core"}).groups == frozenset({"excel_core"})  # лишнее пропускается
    with pytest.raises(config.ConfigError):  # ни одного известного набора — НЕ «включить всё», а отказ запуска
        config.load({"OFFICE_LIVE_TOOLSETS": "nonsense"})


def test_eval_requires_flag_and_full_mode():
    assert not config.load({}).enabled("eval", "destructive")
    assert config.load({"OFFICE_LIVE_ALLOW_EVAL": "1"}).enabled("eval", "destructive")
    assert not config.load({"OFFICE_LIVE_ALLOW_EVAL": "1", "OFFICE_LIVE_MODE": "readonly"}).enabled("eval", "destructive")


# ------------------------------------------------------------------ safety


def test_allowed_dirs_enforced(tmp_path):
    inside = tmp_path / "ok"
    inside.mkdir()
    s = config.load({"OFFICE_LIVE_ALLOWED_DIRS": str(inside)})
    assert safety.check_path(str(inside / "a.xlsx"), "write", safety.EXCEL_EXTS, settings=s).endswith("a.xlsx")
    with pytest.raises(ToolError, match="outside the allowed"):
        safety.check_path(str(tmp_path / "other" / "a.xlsx"), "write", safety.EXCEL_EXTS, settings=s)
    with pytest.raises(ToolError, match="outside the allowed"):
        safety.check_path(str(inside / ".." / "a.xlsx"), "read", settings=s)


def test_write_extension_and_system_dirs():
    s = config.load({})
    with pytest.raises(ToolError, match="Unsupported file extension"):
        safety.check_path(str(Path.home() / "x.exe"), "write", safety.EXCEL_EXTS, settings=s)
    for bad in (os.environ.get("WINDIR", "C:\\Windows"), os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Excel", "XLSTART")):
        with pytest.raises(ToolError, match="system/startup"):
            safety.check_path(os.path.join(bad, "evil.xlsx"), "write", safety.EXCEL_EXTS, settings=s)
    with pytest.raises(ToolError):
        safety.check_path("   ", "read", settings=s)


# ------------------------------------------------------------------ COM-слой


def _com_error(hresult, text="x", scode=0):
    return pywintypes.com_error(hresult, text, (0, None, text, None, 0, scode), None)


def test_busy_detection_and_translation():
    busy = _com_error(-2147418111)  # RPC_E_CALL_REJECTED
    assert com.is_busy(busy) and not com.is_dead(busy)
    assert com.is_busy(_com_error(-2147352567, scode=0x800AC472 - (1 << 32)))  # Excel «занят» приходит через excepinfo
    assert com.is_dead(_com_error(0x80010108 - (1 << 32)))  # RPC_E_DISCONNECTED
    assert "1004" in com.translate(_com_error(-2147352567, "Exception occurred.", 0x800A03EC - (1 << 32)))
    assert com.translate(ToolError("plain")) == "plain"


def test_retry_waits_then_succeeds(monkeypatch):
    monkeypatch.setattr(com.time, "sleep", lambda s: None)
    calls = {"n": 0}

    def flaky():
        calls["n"] += 1
        if calls["n"] < 4:
            raise _com_error(-2147418111)
        return "done"

    assert com._retry(flaky) == "done" and calls["n"] == 4


def test_retry_gives_up_with_clear_message(monkeypatch):
    monkeypatch.setattr(com.config, "SETTINGS", config.load({"OFFICE_LIVE_BUSY_TIMEOUT": "0"}))
    monkeypatch.setattr(com.time, "sleep", lambda s: None)

    def always_busy():
        raise _com_error(-2147418111)

    with pytest.raises(AppBusyError, match="busy"):
        com._retry(always_busy)


def test_retry_does_not_swallow_real_errors():
    def broken():
        raise _com_error(-2147352567, "bad arg", 0x800A03EC - (1 << 32))

    with pytest.raises(pywintypes.com_error):
        com._retry(broken)


def test_proxy_forbids_keyword_arguments():
    class Fake:
        def method(self, *a):
            return a

    p = com.Proxy(Fake())
    # у Proxy.__getattr__ вызываемые значения оборачиваются в _Method, который отвергает kwargs
    with pytest.raises(TypeError, match="positional"):
        p.method(After=1)
    assert p.method(1, 2) == (1, 2)


def test_proxy_unwraps_proxy_arguments():
    class Fake:
        def same(self, other):
            return other is token

    token = object()
    wrapped = com.Proxy(token)
    assert com.raw(wrapped) is token
    assert com.Proxy(Fake()).same(wrapped) is True


def test_run_com_turns_exceptions_into_toolerror():
    def boom():
        raise KeyError("missing")

    with pytest.raises(ToolError, match="KeyError"):
        com.run_com(boom)
    assert com.run_com(lambda x, y=1: x + y, (2,), {"y": 3}) == 5


# ------------------------------------------------------------------ реестр инструментов


@pytest.fixture(scope="module")
def catalog():
    from office_live.app import CATALOG, mcp

    return CATALOG, mcp


def test_catalog_is_sane(catalog):
    cat, _ = catalog
    assert len(cat) >= 100
    for name, info in cat.items():
        assert re.fullmatch(r"[a-z][a-z0-9_]+", name), name
        assert info.group in config.GROUPS + ("eval", "history"), name
        assert info.kind in config.ALL_KINDS, name
        assert len(info.doc.split("\n")[0]) >= 15, f"{name}: add a one-line description"


def test_instruction_tool_count_matches_catalog(catalog):
    from office_live.registry import INSTRUCTIONS

    cat, _ = catalog
    count = re.search(r"The catalog contains (\d+) tools", INSTRUCTIONS)
    assert count is not None
    assert int(count.group(1)) == len(cat) == 125


def test_every_registered_tool_publishes_a_valid_schema(catalog):
    _, mcp = catalog
    tools = asyncio.run(mcp.list_tools())
    names = [t.name for t in tools]
    assert len(names) == len(set(names))
    assert "excel_write_range" in names and "word_create_table" in names and "bridge_word_table_to_excel" in names
    for t in tools:
        schema = t.input_schema
        assert schema.get("type") == "object", t.name
        assert t.annotations is not None, t.name
    by = {t.name: t for t in tools}
    assert by["excel_read_range"].annotations.read_only_hint is True
    assert by["excel_delete_sheet"].annotations.destructive_hint is True
    assert by["excel_write_range"].annotations.read_only_hint is False


def test_readonly_registers_no_writers(monkeypatch):
    s = config.load({"OFFICE_LIVE_MODE": "readonly"})
    from office_live.app import CATALOG

    writers = [n for n, i in CATALOG.items() if i.kind in ("write", "destructive", "save") and s.enabled(i.group, i.kind)]
    assert writers == []


def test_source_never_uses_known_bad_com_patterns():
    """Ловушки late-binding pywin32, найденные живыми пробами: Resize/Offset и именованные аргументы COM-вызовов."""
    bad = re.compile(r"(?<!lo)\.(Resize|Offset)\((?!\.\.\.)")  # ListObject.Resize(range) — настоящий метод, не свойство Range
    offenders = []
    for path in (ROOT / "office_live").glob("*.py"):
        for i, line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
            code = line.split("#")[0]
            if bad.search(code):
                offenders.append(f"{path.name}:{i}: {line.strip()}")
    assert offenders == []


# ------------------------------------------------------------------ VBA


def test_ovba_decompress_literals_and_copy_token():
    # контейнер: 01, заголовок чанка (сжатый, размер), флаг-байт 00 + 8 литералов
    assert vba.decompress(bytes([0x01, 0x08, 0xB0, 0x00]) + b"ABCDEFGH") == b"ABCDEFGH"
    # три литерала 'abc' и токен-копия (смещение 3, длина 6) -> 'abcabcabc'
    assert vba.decompress(bytes([0x01, 0x05, 0xB0, 0x08]) + b"abc" + bytes([0x03, 0x20])) == b"abcabcabc"
    with pytest.raises(ValueError):
        vba.decompress(b"\x00\x00")


def test_vba_project_without_olefile_reports_names(monkeypatch):
    monkeypatch.setattr(vba, "olefile", None)
    info = vba.read_vba_project(b"junk Module=Module1 junk Class=Sheet1")
    assert info["readable"] is False and "Module1" in info["module_names"]
