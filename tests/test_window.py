import copy
import ctypes
import os

import pytest

from office_live import config, protocol, registry, undo, win32, window
from office_live.errors import ToolError
from tests.history_fakes import Collection, fake_office  # noqa: F401
from tests.window_fakes import FakeWin32


@pytest.fixture
def desktop(office, monkeypatch):
    fake = FakeWin32()
    monkeypatch.setattr(window, "Win32", lambda: fake)
    monkeypatch.delenv("WT_SESSION", raising=False)
    return fake


def test_status_readonly_no_side_effects(office, desktop, monkeypatch):
    monkeypatch.setattr(config, "SETTINGS", config.load({"OFFICE_LIVE_MODE": "readonly"}))
    before = copy.deepcopy(desktop.data)
    result = office.call("office_window", workbook=office.wb.Name)
    assert registry.CATALOG["office_window"].kind == "ui"
    assert registry.CATALOG["office_window"].registered
    assert result["office_windows"][0]["hwnd"] == 42
    assert result["office_windows"][0]["busy"] is False
    assert result["office_windows"][0]["foreground"] is False
    assert result["chat_window"]["hwnd"] == 142 and result["chat_detection"] == "ancestor"
    assert len(result["monitors"]) == 2
    assert before == desktop.data and not desktop.calls and desktop.depth == 0
    assert not undo.STACKS and not list((office.tmp / "journal").glob("*"))


@pytest.mark.parametrize("layout", ["office_left_chat_right", "chat_left_office_right"])
def test_layout_physical_dpi_frames_vertical_taskbar_and_restore(office, desktop, layout):
    desktop.data[42]["state"] = "minimized"
    desktop.data[142]["state"] = "maximized"
    before = copy.deepcopy(desktop.data)
    out = office.call("office_window", action="arrange", workbook=office.wb.Name, layout=layout, office_share=0.4)
    assert out["ok"] and out["monitor"]["number"] == 2
    assert out["office_windows"][0]["monitor"]["number"] == 2
    assert out["office_windows"][0]["monitor"]["dpi"] == 144
    office_rect = out["office_windows"][0]["frame_rect"]
    chat_rect = out["chat_window"]["frame_rect"]
    assert office_rect == ([-2500, -400, -1500, 1040] if layout.startswith("office") else [-1000, -400, 0, 1040])
    assert chat_rect == ([-1500, -400, 0, 1040] if layout.startswith("office") else [-2500, -400, -1000, 1040])
    assert len([c for c in desktop.calls if c[:2] == ("position", 42)]) == 2  # borders remeasured at 150%
    assert desktop.foreground_hwnd == 142
    restored = office.call("office_window", action="restore")
    assert restored["ok"] and set(restored["restored"]) == {42, 142}
    assert desktop.data == before and not window._LAST


@pytest.mark.parametrize("monitor", ["primary", "office", 1, "1"])
def test_monitor_selection(office, desktop, monitor):
    out = office.call("office_window", action="arrange", workbook=office.wb.Name, monitor=monitor)
    assert out["monitor"]["number"] == 1
    assert out["office_windows"][0]["frame_rect"] == [0, 0, 640, 1040]
    assert out["chat_window"]["frame_rect"] == [640, 0, 1920, 1040]
    assert out["chat_window"]["monitor"]["number"] == 1


@pytest.mark.parametrize("layout", ["office_maximized", "excel_word_side_by_side"])
def test_office_only_layout_never_moves_chat(office, desktop, layout):
    before = copy.deepcopy(desktop.data)
    kwargs = {"document": office.doc.Name} if layout == "excel_word_side_by_side" else {}
    out = office.call("office_window", action="arrange", workbook=office.wb.Name, layout=layout, include_chat=False, **kwargs)
    assert out["ok"] and out["chat_window"] is None
    assert desktop.data[142] == before[142]
    assert {c[1] for c in desktop.calls} == ({42, 84} if kwargs else {42})
    assert out["office_windows"][0]["state"] == ("normal" if kwargs else "maximized")
    assert office.call("office_window", action="restore")["ok"]
    assert desktop.data == before


def test_unknown_chat_office_only_fallback(office, desktop):
    desktop.tree[os.getpid()]["parent"] = 0
    out = office.call("office_window", action="arrange", workbook=office.wb.Name)
    assert out["ok"] and out["chat_window"] is None
    assert out["monitor"]["number"] == 1 and len(out["warnings"]) == 2
    assert {c[1] for c in desktop.calls} == {42}


@pytest.mark.parametrize("name", ["ZCode.exe", "Claude.exe", "Code.exe", "Cursor.exe"])
def test_gui_ancestor_detection(desktop, name):
    desktop.tree[500]["name"] = name
    desktop.tree[600] = {"parent": 500, "name": "node.exe"}
    desktop.tree[os.getpid()]["parent"] = 600
    assert window.chat_window(desktop, desktop.tree)[0]["hwnd"] == 142


def test_ambiguous_ancestor_and_explicit_hint(desktop):
    desktop.add(143, 500, "Chrome_WidgetWin_1", [0, 0, 100, 100])
    assert window.chat_window(desktop, desktop.tree)[0] is None
    assert window.chat_window(desktop, desktop.tree, "code.exe")[0] is None
    assert window.chat_window(desktop, desktop.tree, "142")[0]["hwnd"] == 142
    assert window.chat_window(desktop, desktop.tree, "missing")[0] is None


@pytest.mark.parametrize("process,cls", [("EXCEL.EXE", "XLMAIN"), ("WINWORD.EXE", "OpusApp"),
                                          ("explorer.exe", "CabinetWClass"), ("other.exe", "WorkerW"), ("other.exe", "Shell_TrayWnd")])
def test_never_accept_office_or_shell_as_chat(desktop, process, cls):
    desktop.tree[500]["name"] = process
    desktop.data[142]["class"] = cls
    assert window.chat_window(desktop, desktop.tree, "Window")[0] is None
    assert window.chat_window(desktop, desktop.tree)[0] is None


@pytest.mark.parametrize("field,value", [("visible", False), ("cloaked", True), ("owner", 42), ("root", 900), ("tool_window", True)])
def test_chat_must_be_visible_main_window(desktop, field, value):
    desktop.data[142][field] = value
    assert window.chat_window(desktop, desktop.tree)[0] is None


def test_terminal_and_conhost_conservative_detection(desktop, monkeypatch):
    desktop.tree[os.getpid()]["parent"] = 0
    desktop.tree[500]["name"] = "conhost.exe"
    desktop.console = 142
    assert window.chat_window(desktop, desktop.tree)[1] == "inherited_console"
    desktop.console = 999  # invisible pseudoconsole/message-only handle
    desktop.tree[500]["name"] = "WindowsTerminal.exe"
    assert window.chat_window(desktop, desktop.tree)[1] == "unique_windows_terminal"
    desktop.console = 0
    assert window.chat_window(desktop, desktop.tree)[0] is None
    monkeypatch.setenv("WT_SESSION", "test-session")
    assert window.chat_window(desktop, desktop.tree)[0]["hwnd"] == 142
    desktop.add(143, 500, "CASCADIA_HOSTING_WINDOW_CLASS", [0, 0, 100, 100])
    assert window.chat_window(desktop, desktop.tree)[0] is None


def test_process_parent_cycle_is_bounded(desktop):
    desktop.tree[os.getpid()]["parent"] = 600
    desktop.tree[600] = {"parent": os.getpid(), "name": "node.exe"}
    assert window.chat_window(desktop, desktop.tree)[0] is None


@pytest.mark.parametrize("allowed", [True, False])
def test_focus_verified_and_denial_flashes(office, desktop, allowed):
    desktop.focus_allowed = allowed
    desktop.data[42]["state"] = "minimized"
    out = office.call("office_window", action="focus", workbook=office.wb.Name)
    assert out["foreground"] is allowed
    assert desktop.data[42]["state"] == "normal"
    assert (("flash", 42) in desktop.calls) is not allowed
    assert bool(out["reason"]) is not allowed


def test_foreground_owned_popup_but_not_another_document(desktop):
    desktop.add(44, 42, "#32770", [0, 0, 400, 200])
    desktop.data[44]["owner"] = 42
    desktop.foreground_hwnd = 44
    assert window.is_foreground(desktop, 42)
    desktop.data[44]["owner"] = 0
    assert not window.is_foreground(desktop, 42)
    desktop.data[42]["owner"] = 44
    assert window.is_foreground(desktop, 42)


@pytest.mark.parametrize("blocked", ["hung", "modal", "hidden", "cloaked"])
def test_no_focus_or_arrange_for_blocked_window(office, desktop, blocked):
    if blocked == "hung":
        desktop.hung.add(42)
    else:
        desktop.data[42][{"modal": "enabled", "hidden": "visible", "cloaked": "cloaked"}[blocked]] = blocked == "cloaked"
    out = office.call("office_window", action="status", workbook=office.wb.Name)
    assert out["office_windows"][0]["responding"] is (blocked != "hung")
    out = office.call("office_window", action="focus", workbook=office.wb.Name)
    assert out["foreground"] is False and not desktop.calls
    with pytest.raises(ToolError):
        office.call("office_window", action="arrange", workbook=office.wb.Name)
    assert not desktop.calls


def test_chat_hung_preflight_before_any_change(office, desktop):
    desktop.hung.add(142)
    with pytest.raises(ToolError, match="responding"):
        office.call("office_window", action="arrange", workbook=office.wb.Name)
    assert not desktop.calls


@pytest.mark.parametrize("property", ["Ready", "Interactive"])
def test_excel_busy_status_and_arrange_refused(office, desktop, property):
    setattr(office.excel, property, False)
    out = office.call("office_window", workbook=office.wb.Name)
    assert out["office_windows"][0]["busy"]
    with pytest.raises(ToolError, match="busy"):
        office.call("office_window", action="arrange", workbook=office.wb.Name)
    assert not desktop.calls


def test_excel_unknown_busy_is_not_claimed_ready(office, desktop):
    del office.excel.Ready
    assert office.call("office_window")["office_windows"][0]["busy"] is None
    with pytest.raises(ToolError, match="unknown"):
        office.call("office_window", action="arrange")
    assert not desktop.calls


@pytest.mark.parametrize("kwargs", [{"action": "bad"}, {"layout": "bad"}, {"application": "bad"}, {"office_share": 0.1},
                                    {"office_share": float("nan")}, {"office_share": True}, {"monitor": 0}, {"monitor": True},
                                    {"monitor": 33}, {"chat_hint": "a\n"}, {"document": "Draft.docx"}, {"layout": "excel_word_side_by_side"}])
def test_invalid_input_no_changes(office, desktop, kwargs):
    args = {"action": "arrange", "workbook": office.wb.Name, **kwargs}
    with pytest.raises(ToolError):
        office.call("office_window", **args)
    assert not desktop.calls and desktop.depth == 0


def test_partial_failure_rolls_back_both_windows(office, desktop):
    before = copy.deepcopy(desktop.data)
    desktop.fail_position = 142
    out = office.call("office_window", action="arrange", workbook=office.wb.Name)
    assert not out["ok"] and out["applied"] == [42, 142]
    assert set(out["rolled_back"]) == {42, 142} and not out["rollback_errors"]
    assert desktop.data == before


def test_restore_last_arrange_only_and_reused_handle(office, desktop):
    office.call("office_window", action="arrange", include_chat=False)
    before_second = copy.deepcopy(desktop.data)
    office.call("office_window", action="arrange", include_chat=False, office_share=0.7)
    assert office.call("office_window", action="restore")["ok"]
    assert desktop.data == before_second
    office.call("office_window", action="arrange", include_chat=False)
    desktop.data[42]["pid"] = 999
    desktop.calls.clear()
    out = office.call("office_window", action="restore")
    assert not out["ok"] and "identity" in out["errors"][0]["reason"] and not desktop.calls


def test_frame_unavailable_and_actual_constraints_reported(office, desktop, monkeypatch):
    desktop.data[42]["frame_available"] = False
    out = office.call("office_window", action="arrange", include_chat=False)
    assert "DWM" in out["warnings"][-1]
    assert out["office_windows"][0]["rect"] == [0, 0, 640, 1040]
    monkeypatch.setattr(desktop, "position", lambda *args: None)
    out = office.call("office_window", action="arrange", include_chat=False, office_share=0.5)
    assert "did not fit" in out["warnings"][-1]
    assert out["office_windows"][0]["rect"] != out["requested"][0]["frame_rect"]


def test_restore_preflights_chat_before_restoring_office(office, desktop):
    office.call("office_window", action="arrange")
    desktop.hung.add(142)
    desktop.calls.clear()
    out = office.call("office_window", action="restore")
    assert not out["ok"] and not desktop.calls and window._LAST
    desktop.hung.clear()
    assert office.call("office_window", action="restore")["ok"]


def test_restore_rechecks_allowed_directories(office, desktop, monkeypatch):
    office.wb.FullName, office.wb.Path = str(office.tmp / "Book.xlsx"), str(office.tmp)
    office.call("office_window", action="arrange")
    desktop.calls.clear()
    monkeypatch.setattr(config, "SETTINGS", config.load({"OFFICE_LIVE_ALLOWED_DIRS": str(office.tmp / "other")}))
    with pytest.raises(ToolError):
        office.call("office_window", action="restore")
    assert not desktop.calls


def test_office_target_is_exact_and_not_application_window(office, desktop):
    with pytest.raises(ToolError):
        office.call("office_window", workbook="Book")
    office.excel.Hwnd = 999
    assert office.call("office_window", workbook=office.wb.Name)["office_windows"][0]["hwnd"] == 42
    office.wb.Windows = Collection([])
    with pytest.raises(ToolError, match="no window"):
        office.call("office_window", workbook=office.wb.Name)
    assert not desktop.calls


def test_word_target_uses_document_window(office, desktop):
    out = office.call("office_window", application="word", include_chat=False)
    assert out["office_windows"][0]["hwnd"] == 84


def test_protocol_focus_uses_same_adapter(desktop, monkeypatch):
    monkeypatch.setattr(win32, "Win32", lambda: desktop)
    desktop.focus_allowed = False
    assert protocol.foreground(42)["foreground"] is False
    assert ("flash", 42) in desktop.calls


def test_native_window_guard_is_armed():
    for dll in (ctypes.windll.user32, ctypes.WinDLL("user32")):
        for name in ("SetWindowPos", "ShowWindow", "ShowWindowAsync", "SetForegroundWindow", "SetWindowPlacement"):
            with pytest.raises(AssertionError, match="fake Win32"):
                getattr(dll, name)(0)
    with pytest.raises(AssertionError, match="fake Win32"):
        win32.Win32()


@pytest.mark.parametrize("preset", ["all", "core", "excel", "word", "window"])
def test_group_presets_readonly(preset):
    assert config.load({"OFFICE_LIVE_TOOLSETS": preset, "OFFICE_LIVE_MODE": "readonly"}).enabled("window", "ui")
    assert not config.load({"OFFICE_LIVE_TOOLSETS": "excel_core"}).enabled("window", "ui")
