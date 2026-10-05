"""ctypes ABI/flags and failure paths using fake DLL exports, never native calls."""

import ctypes as C
from types import SimpleNamespace as NS

import pytest

from office_live import win32
from office_live.errors import ToolError


class Function:
    def __init__(self):
        self.calls = []
        self.impl = lambda *args: 1

    def __call__(self, *args):
        self.calls.append(args)
        return self.impl(*args)


class DLL:
    def __init__(self):
        self.entries = {}

    def __getattr__(self, name):
        return self.entries.setdefault(name, Function())


@pytest.fixture
def native(monkeypatch):
    dlls = {name: DLL() for name in ("user32", "kernel32", "dwmapi", "shcore")}
    monkeypatch.setattr(win32, "_load_dll", dlls.__getitem__)
    api = win32.Win32()
    api.u.IsHungAppWindow.impl = lambda _: 0
    return NS(api=api, u=api.u, k=api.k, d=api.d, s=api.s)


def test_pointer_sized_signatures(native):
    assert native.u.GetForegroundWindow.restype is win32.W.HWND
    assert native.u.SetWindowPos.argtypes[:2] == [win32.W.HWND, win32.W.HWND]
    assert native.u.SendMessageTimeoutW.restype is C.c_ssize_t
    assert native.k.CreateToolhelp32Snapshot.restype is win32.W.HANDLE
    assert native.u.SetThreadDpiAwarenessContext.restype is win32.W.HANDLE
    assert C.sizeof(win32.Placement) == 44
    assert C.sizeof(win32.ProcessEntry) == (568 if C.sizeof(C.c_void_p) == 8 else 556)


def test_dpi_v2_context_restored_even_on_failure(native):
    native.u.SetThreadDpiAwarenessContext.impl = lambda value: 1234
    with pytest.raises(RuntimeError), native.api.dpi_context():
        raise RuntimeError("test")
    calls = native.u.SetThreadDpiAwarenessContext.calls
    assert calls[0][0].value == C.c_void_p(-4).value
    assert calls[1] == (1234,)
    native.u.SetThreadDpiAwarenessContext.impl = lambda _: 0
    with pytest.raises(ToolError, match="DPI"):
        with native.api.dpi_context():
            pytest.fail("invalid DPI context must not proceed")


def test_hung_and_bounded_wm_null(native):
    assert native.api.responsive(42)
    args = native.u.SendMessageTimeoutW.calls[0]
    assert args[:6] == (42, 0, 0, 0, 0x23, 2000)
    native.u.SendMessageTimeoutW.impl = lambda *args: 0
    assert not native.api.responsive(42)
    native.u.IsHungAppWindow.impl = lambda _: 1
    count = len(native.u.SendMessageTimeoutW.calls)
    assert not native.api.responsive(42)
    assert len(native.u.SendMessageTimeoutW.calls) == count


def test_position_async_no_activate_no_zorder(native, monkeypatch):
    monkeypatch.setattr(native.api, "info", lambda _: {"rect": [-10, 20, 700, 800]})
    native.api.position(42, [-10, 20, 700, 800])
    assert native.u.SetWindowPos.calls == [(42, None, -10, 20, 710, 780, 0x4014)]
    native.u.SetWindowPos.impl = lambda *args: 0
    with pytest.raises(ToolError, match="refused"):
        native.api.position(42, [0, 0, 1, 1])


def test_placement_keeps_workspace_coordinates_opaque(native):
    stored = win32.Placement()
    stored.length, stored.flags, stored.show = 44, 2, 2
    stored.normal = win32.W.RECT(-2400, 40, -1200, 900)

    def get(hwnd, ptr):
        assert ptr._obj.length == C.sizeof(stored)
        C.memmove(ptr, C.byref(stored), C.sizeof(stored))
        return 1

    native.u.GetWindowPlacement.impl = get
    saved = native.api.placement(42)
    native.api.restore_placement(42, saved)
    restored = native.u.SetWindowPlacement.calls[0][1]._obj
    assert restored.flags == 6 and restored.show == 2
    assert win32.rect(restored.normal) == [-2400, 40, -1200, 900]
    assert not native.u.SetWindowPos.calls


def test_monitor_work_area_scaling_and_primary_order(native):
    def enum(dc, rect, callback, data):
        callback(2, None, None, 0)
        callback(1, None, None, 0)
        return 1

    def info(handle, ptr):
        value = ptr._obj
        value.device = f"DISPLAY{handle}"
        value.flags = int(handle == 1)
        value.monitor = win32.W.RECT(-2560, -400, 0, 1040) if handle == 2 else win32.W.RECT(0, 0, 1920, 1080)
        value.work = win32.W.RECT(-2500, -400, 0, 1040) if handle == 2 else win32.W.RECT(0, 0, 1920, 1040)
        return 1

    def scale(handle, ptr):
        ptr._obj.value = 150 if handle == 2 else 100
        return 0

    native.u.EnumDisplayMonitors.impl = enum
    native.u.GetMonitorInfoW.impl = info
    native.s.GetScaleFactorForMonitor.impl = scale
    result = native.api.monitors()
    assert [m["number"] for m in result] == [1, 2]
    assert result[1]["dpi"] == 144 and result[1]["scale"] == 1.5
    assert result[1]["work_area"] == [-2500, -400, 0, 1040]
    native.s.GetScaleFactorForMonitor.impl = lambda *args: -1
    assert native.api.monitors()[0]["dpi"] is None


def test_snapshot_handles_closed_and_parent_fields(native):
    native.k.CreateToolhelp32Snapshot.impl = lambda *args: 101

    def first(handle, ptr):
        value = ptr._obj
        assert value.size == C.sizeof(win32.ProcessEntry)
        value.pid, value.parent, value.exe = 50, 40, "Code.exe"
        return 1

    native.k.Process32FirstW.impl = first
    native.k.Process32NextW.impl = lambda *args: 0
    assert native.api.processes() == {50: {"parent": 40, "name": "Code.exe"}}
    assert native.k.CloseHandle.calls == [(101,)]
    native.k.CreateToolhelp32Snapshot.impl = lambda *args: C.c_void_p(-1).value
    with pytest.raises(ToolError):
        native.api.processes()


def test_foreground_request_does_not_inject_keys_or_attach_threads(native):
    native.api.request_foreground(42, 100)
    assert native.u.AllowSetForegroundWindow.calls == [(100,)]
    assert native.u.SetForegroundWindow.calls == [(42,)]
    assert "AttachThreadInput" not in native.u.entries and "keybd_event" not in native.u.entries
    native.api.flash(42)
    value = native.u.FlashWindowEx.calls[0][0]._obj
    assert value.hwnd == 42 and value.flags == 2 and value.count == 3


def test_minimized_from_maximized_restores_twice_for_normal(native, monkeypatch):
    state = {"value": "minimized"}

    def show(hwnd, command):
        assert hwnd == 42 and command == 9
        state["value"] = "maximized" if state["value"] == "minimized" else "normal"
        return 1

    native.u.ShowWindowAsync.impl = show
    monkeypatch.setattr(native.api, "info", lambda _: {"state": state["value"]})
    monkeypatch.setattr(win32.time, "sleep", lambda _: None)
    native.api.show(42, "normal")
    assert native.u.ShowWindowAsync.calls == [(42, 9), (42, 9)]


def test_window_info_dwm_frame_cloaking_and_minimized(native):
    native.u.GetWindowThreadProcessId.impl = lambda hwnd, ptr: setattr(ptr._obj, "value", 100) or 101
    native.u.GetWindowTextW.impl = lambda hwnd, buf, size: setattr(buf, "value", "Excel") or 5
    native.u.GetClassNameW.impl = lambda hwnd, buf, size: setattr(buf, "value", "XLMAIN") or 6
    native.u.GetWindowRect.impl = lambda hwnd, ptr: C.memmove(ptr, C.byref(win32.W.RECT(-2512, -400, 12, 1052)), 16) and 1
    native.u.GetWindow.impl = lambda *args: 0
    native.u.GetAncestor.impl = lambda hwnd, flags: hwnd
    native.u.IsIconic.impl = native.u.IsZoomed.impl = lambda _: 0
    native.u.GetWindowLongW.impl = lambda *args: 0
    native.u.GetDpiForWindow.impl = lambda _: 144

    def dwm(hwnd, attr, ptr, size):
        if attr == 9:
            C.memmove(ptr, C.byref(win32.W.RECT(-2500, -400, 0, 1040)), 16)
        else:
            assert attr == 14
            ptr._obj.value = 1
        return 0

    native.d.DwmGetWindowAttribute.impl = dwm
    out = native.api.info(42)
    assert out["rect"] == [-2512, -400, 12, 1052]
    assert out["frame_rect"] == [-2500, -400, 0, 1040]
    assert out["pid"] == 100 and out["thread"] == 101
    assert out["window_dpi"] == 144 and out["cloaked"]
    native.u.IsIconic.impl = lambda _: 1
    assert native.api.info(42)["frame_rect"] is None
    native.u.IsWindow.impl = lambda _: 0
    with pytest.raises(ToolError, match="closed"):
        native.api.info(42)
