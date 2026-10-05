"""Small, lazy ctypes boundary. No DLL loading or desktop access on import.

All rectangles are physical screen pixels inside dpi_context(). WINDOWPLACEMENT
is deliberately opaque: its workspace coordinates must never go to SetWindowPos.
"""

import ctypes as C
import time
from contextlib import contextmanager
from ctypes import wintypes as W

from .errors import ToolError


class MonitorInfo(C.Structure):
    _fields_ = [("size", W.DWORD), ("monitor", W.RECT), ("work", W.RECT), ("flags", W.DWORD), ("device", W.WCHAR * 32)]


class Placement(C.Structure):
    _fields_ = [("length", W.UINT), ("flags", W.UINT), ("show", W.UINT), ("minimum", W.POINT), ("maximum", W.POINT), ("normal", W.RECT)]


class ProcessEntry(C.Structure):
    _fields_ = [("size", W.DWORD), ("usage", W.DWORD), ("pid", W.DWORD), ("heap", C.c_size_t), ("module", W.DWORD),
                ("threads", W.DWORD), ("parent", W.DWORD), ("priority", W.LONG), ("flags", W.DWORD), ("exe", W.WCHAR * 260)]


class FlashInfo(C.Structure):
    _fields_ = [("size", W.UINT), ("hwnd", W.HWND), ("flags", W.DWORD), ("count", W.UINT), ("timeout", W.DWORD)]


def rect(value):
    return [value.left, value.top, value.right, value.bottom]


def _load_dll(name):
    return C.WinDLL(name, use_last_error=True)


def _bind(dll, name, result, *args):
    fn = getattr(dll, name)
    fn.restype, fn.argtypes = result, list(args)
    return fn


class Win32:
    def __init__(self):
        self.u, self.k, self.d, self.s = (_load_dll(n) for n in ("user32", "kernel32", "dwmapi", "shcore"))
        self.enum_proc = C.WINFUNCTYPE(W.BOOL, W.HWND, W.LPARAM)
        self.monitor_proc = C.WINFUNCTYPE(W.BOOL, W.HMONITOR, W.HDC, C.POINTER(W.RECT), W.LPARAM)
        for name in ("IsWindow", "IsWindowVisible", "IsIconic", "IsZoomed", "IsWindowEnabled", "IsHungAppWindow", "SetForegroundWindow"):
            _bind(self.u, name, W.BOOL, W.HWND)
        _bind(self.u, "GetForegroundWindow", W.HWND)
        _bind(self.u, "GetWindow", W.HWND, W.HWND, W.UINT)
        _bind(self.u, "GetAncestor", W.HWND, W.HWND, W.UINT)
        _bind(self.u, "GetLastActivePopup", W.HWND, W.HWND)
        _bind(self.u, "GetWindowThreadProcessId", W.DWORD, W.HWND, C.POINTER(W.DWORD))
        for name in ("GetWindowTextW", "GetClassNameW"):
            _bind(self.u, name, C.c_int, W.HWND, W.LPWSTR, C.c_int)
        _bind(self.u, "GetWindowRect", W.BOOL, W.HWND, C.POINTER(W.RECT))
        _bind(self.u, "GetWindowPlacement", W.BOOL, W.HWND, C.POINTER(Placement))
        _bind(self.u, "SetWindowPlacement", W.BOOL, W.HWND, C.POINTER(Placement))
        _bind(self.u, "ShowWindowAsync", W.BOOL, W.HWND, C.c_int)
        _bind(self.u, "SetWindowPos", W.BOOL, W.HWND, W.HWND, C.c_int, C.c_int, C.c_int, C.c_int, W.UINT)
        _bind(self.u, "SendMessageTimeoutW", C.c_ssize_t, W.HWND, W.UINT, W.WPARAM, W.LPARAM, W.UINT, W.UINT, C.POINTER(C.c_size_t))
        _bind(self.u, "AllowSetForegroundWindow", W.BOOL, W.DWORD)
        _bind(self.u, "FlashWindowEx", W.BOOL, C.POINTER(FlashInfo))
        _bind(self.u, "SetThreadDpiAwarenessContext", W.HANDLE, W.HANDLE)
        _bind(self.u, "GetDpiForWindow", W.UINT, W.HWND)
        _bind(self.u, "GetWindowLongW", W.LONG, W.HWND, C.c_int)
        _bind(self.u, "EnumWindows", W.BOOL, self.enum_proc, W.LPARAM)
        _bind(self.u, "EnumDisplayMonitors", W.BOOL, W.HDC, C.POINTER(W.RECT), self.monitor_proc, W.LPARAM)
        _bind(self.u, "GetMonitorInfoW", W.BOOL, W.HMONITOR, C.POINTER(MonitorInfo))
        _bind(self.u, "MonitorFromWindow", W.HMONITOR, W.HWND, W.DWORD)
        _bind(self.d, "DwmGetWindowAttribute", W.LONG, W.HWND, W.DWORD, C.c_void_p, W.DWORD)
        _bind(self.s, "GetScaleFactorForMonitor", W.LONG, W.HMONITOR, C.POINTER(C.c_int))
        _bind(self.k, "CreateToolhelp32Snapshot", W.HANDLE, W.DWORD, W.DWORD)
        _bind(self.k, "Process32FirstW", W.BOOL, W.HANDLE, C.POINTER(ProcessEntry))
        _bind(self.k, "Process32NextW", W.BOOL, W.HANDLE, C.POINTER(ProcessEntry))
        _bind(self.k, "CloseHandle", W.BOOL, W.HANDLE)
        _bind(self.k, "GetConsoleWindow", W.HWND)

    @contextmanager
    def dpi_context(self):
        old = self.u.SetThreadDpiAwarenessContext(C.c_void_p(-4))  # PER_MONITOR_AWARE_V2
        if not old:
            raise ToolError("Per-monitor DPI awareness v2 is unavailable; no windows were changed.")
        try:
            yield
        finally:
            self.u.SetThreadDpiAwarenessContext(old)

    def processes(self):
        handle = self.k.CreateToolhelp32Snapshot(2, 0)  # TH32CS_SNAPPROCESS
        if handle in (None, C.c_void_p(-1).value):
            raise ToolError("Cannot read the process tree.")
        result, entry = {}, ProcessEntry()
        entry.size = C.sizeof(entry)
        try:
            ok = self.k.Process32FirstW(handle, C.byref(entry))
            while ok:
                result[entry.pid] = {"parent": entry.parent, "name": entry.exe}
                ok = self.k.Process32NextW(handle, C.byref(entry))
        finally:
            self.k.CloseHandle(handle)
        return result

    def windows(self):
        handles = []
        callback = self.enum_proc(lambda hwnd, _: handles.append(int(hwnd)) or True)
        if not self.u.EnumWindows(callback, 0):
            raise ToolError("Cannot enumerate desktop windows.")
        return handles

    def console_window(self):
        # Only OUR associated console. Never AttachConsole/FreeConsole on the MCP server.
        return int(self.k.GetConsoleWindow() or 0)

    def info(self, hwnd):
        if not hwnd or not self.u.IsWindow(hwnd):
            raise ToolError("The window was closed or its handle is no longer valid.")
        pid = W.DWORD()
        thread = self.u.GetWindowThreadProcessId(hwnd, C.byref(pid))
        title, cls, bounds = C.create_unicode_buffer(1024), C.create_unicode_buffer(256), W.RECT()
        self.u.GetWindowTextW(hwnd, title, len(title))
        self.u.GetClassNameW(hwnd, cls, len(cls))
        if not self.u.GetWindowRect(hwnd, C.byref(bounds)):
            raise ToolError("Cannot read the window rectangle.")
        state = "minimized" if self.u.IsIconic(hwnd) else "maximized" if self.u.IsZoomed(hwnd) else "normal"
        visible = bool(self.u.IsWindowVisible(hwnd))
        frame = W.RECT()
        has_frame = visible and state != "minimized" and self.d.DwmGetWindowAttribute(hwnd, 9, C.byref(frame), C.sizeof(frame)) == 0
        cloaked = W.DWORD()
        self.d.DwmGetWindowAttribute(hwnd, 14, C.byref(cloaked), C.sizeof(cloaked))
        return {"hwnd": hwnd, "pid": pid.value, "thread": thread, "title": title.value, "class": cls.value,
                "owner": int(self.u.GetWindow(hwnd, 4) or 0), "root": int(self.u.GetAncestor(hwnd, 2) or 0),
                "state": state, "visible": visible, "cloaked": bool(cloaked.value), "enabled": bool(self.u.IsWindowEnabled(hwnd)),
                "tool_window": bool(self.u.GetWindowLongW(hwnd, -20) & 0x80),
                "rect": rect(bounds), "frame_rect": rect(frame) if has_frame else None,
                "monitor_id": int(self.u.MonitorFromWindow(hwnd, 2) or 0), "window_dpi": self.u.GetDpiForWindow(hwnd)}

    def monitors(self):
        handles, result = [], []
        callback = self.monitor_proc(lambda monitor, dc, bounds, data: handles.append(int(monitor)) or True)
        if not self.u.EnumDisplayMonitors(None, None, callback, 0):
            raise ToolError("Cannot enumerate monitors.")
        for handle in handles:
            info = MonitorInfo()
            info.size = C.sizeof(info)
            if not self.u.GetMonitorInfoW(handle, C.byref(info)):
                raise ToolError("Cannot read the monitor work area.")
            scale = C.c_int()
            ok = self.s.GetScaleFactorForMonitor(handle, C.byref(scale)) == 0
            result.append({"id": handle, "name": info.device, "rect": rect(info.monitor), "work_area": rect(info.work),
                           "primary": bool(info.flags & 1), "scale": scale.value / 100 if ok else None,
                           "dpi": round(96 * scale.value / 100) if ok else None, "dpi_source": "monitor_scale_factor" if ok else "unavailable"})
        result.sort(key=lambda m: (not m["primary"], m["name"]))
        return [dict(m, number=i) for i, m in enumerate(result, 1)]

    def foreground(self):
        return int(self.u.GetForegroundWindow() or 0)

    def popup(self, hwnd):
        popup = int(self.u.GetLastActivePopup(hwnd) or 0)
        return popup if popup != hwnd and self.u.IsWindowVisible(popup) else 0

    def responsive(self, hwnd):
        if self.u.IsHungAppWindow(hwnd):
            return False
        result = C.c_size_t()
        return bool(self.u.SendMessageTimeoutW(hwnd, 0, 0, 0, 0x23, 2000, C.byref(result)))  # WM_NULL, BLOCK|ABORTIFHUNG|ERRORONEXIT

    def placement(self, hwnd):
        value = Placement()
        value.length = C.sizeof(value)
        if not self.u.GetWindowPlacement(hwnd, C.byref(value)):
            raise ToolError("Cannot save the window placement.")
        return bytes(value)

    def restore_placement(self, hwnd, saved):
        value = Placement.from_buffer_copy(saved)
        value.flags |= 4  # WPF_ASYNCWINDOWPLACEMENT
        if not self.u.SetWindowPlacement(hwnd, C.byref(value)):
            raise ToolError("Windows refused to restore the window placement.")
        self.settle(hwnd)
        deadline = time.monotonic() + 2
        while True:
            actual = Placement.from_buffer_copy(self.placement(hwnd))
            if actual.show == value.show and rect(actual.normal) == rect(value.normal):
                return
            if time.monotonic() >= deadline:
                raise ToolError("Window placement restore timed out or was constrained by Windows.")
            time.sleep(0.025)

    def settle(self, hwnd):
        # Bounded cross-thread synchronization, also lets async positioning take effect.
        if not self.responsive(hwnd):
            raise ToolError("The window stopped responding; restore can be retried.")

    def show(self, hwnd, state):
        command = {"normal": 9, "maximized": 3}[state]
        if not self.u.ShowWindowAsync(hwnd, command):
            raise ToolError("Windows refused to change the window state.")
        deadline = time.monotonic() + 2
        restored_from_max = False
        while self.info(hwnd)["state"] != state and time.monotonic() < deadline:
            # Minimized-from-maximized restores to maximized first (WPF_RESTORETOMAXIMIZED).
            if state == "normal" and self.info(hwnd)["state"] == "maximized" and not restored_from_max:
                if not self.u.ShowWindowAsync(hwnd, 9):
                    raise ToolError("Windows refused to restore the maximized window.")
                restored_from_max = True
            time.sleep(0.025)
        if self.info(hwnd)["state"] != state:
            raise ToolError("Window state change timed out.")

    def position(self, hwnd, bounds):
        left, top, right, bottom = bounds
        # No Z-order change, no activation, async to avoid an unbounded cross-process send.
        if not self.u.SetWindowPos(hwnd, None, left, top, right - left, bottom - top, 0x4014):
            raise ToolError("Windows refused to move the window.")
        self.settle(hwnd)
        deadline = time.monotonic() + 0.5
        while self.info(hwnd)["rect"] != list(bounds) and time.monotonic() < deadline:
            time.sleep(0.025)

    def request_foreground(self, hwnd, pid):
        # ASFW may itself be denied. Never synthesize Alt or attach to a hung input queue.
        self.u.AllowSetForegroundWindow(pid)
        self.u.SetForegroundWindow(hwnd)
        self.responsive(hwnd)  # activation may be asynchronous across input queues

    def flash(self, hwnd):
        value = FlashInfo(C.sizeof(FlashInfo), hwnd, 2, 3, 0)  # taskbar only, three flashes
        self.u.FlashWindowEx(C.byref(value))
