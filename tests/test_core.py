"""Настройки, безопасность, COM-слой (на подделках, без Office), реестр инструментов, VBA-декомпрессор."""

import asyncio
import os
import re
from pathlib import Path

import pytest
import pywintypes

from office_live import com, config, safety, vba
from office_live.errors import AppBusyError, ToolError

ROOT = Path(__file__).resolve().parent.parent


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
    assert config.load({"OFFICE_LIVE_TOOLSETS": "core"}).groups == frozenset({"excel_core", "word_core", "bridge"})
    assert config.load({"OFFICE_LIVE_TOOLSETS": "word"}).groups == frozenset({"word_core", "word_tables", "word_layout"})
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
