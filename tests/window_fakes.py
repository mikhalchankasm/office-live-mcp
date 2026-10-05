"""In-memory desktop. Never subclasses the real adapter or loads a DLL."""

import copy
import os
from contextlib import contextmanager

from office_live.errors import ToolError


class FakeWin32:
    def __init__(self):
        self.calls, self.depth, self.foreground_hwnd = [], 0, 142
        self.focus_allowed = True
        self.fail_position = None
        self.hung, self.popups = set(), {}
        self.console = 0
        self.tree = {os.getpid(): {"parent": 500, "name": "python.exe"}, 500: {"parent": 1, "name": "Code.exe"},
                     42: {"parent": 1, "name": "EXCEL.EXE"}, 84: {"parent": 1, "name": "WINWORD.EXE"}}
        self.screens = [
            {"id": 1, "number": 1, "name": "DISPLAY1", "primary": True, "rect": [0, 0, 1920, 1080], "work_area": [0, 0, 1920, 1040], "dpi": 96, "scale": 1},
            {"id": 2, "number": 2, "name": "DISPLAY2", "primary": False, "rect": [-2560, -400, 0, 1040], "work_area": [-2500, -400, 0, 1040], "dpi": 144, "scale": 1.5},
        ]
        self.data = {}
        self.add(42, 42, "XLMAIN", [100, 100, 1000, 800])
        self.add(84, 84, "OpusApp", [200, 120, 1100, 900])
        self.add(142, 500, "Chrome_WidgetWin_1", [-2000, -200, -800, 800])

    def add(self, hwnd, pid, cls, bounds, **kwargs):
        self.data[hwnd] = dict(hwnd=hwnd, pid=pid, thread=pid + 1, title=f"Window {hwnd}", **{"class": cls},
                               owner=0, root=hwnd, state="normal", visible=True, cloaked=False, enabled=True, tool_window=False,
                               rect=list(bounds), monitor_id=2 if bounds[0] < 0 else 1, frame_available=True, **kwargs)

    @contextmanager
    def dpi_context(self):
        self.depth += 1
        try:
            yield
        finally:
            self.depth -= 1

    def processes(self):
        return self.tree

    def windows(self):
        return list(self.data)

    def console_window(self):
        return self.console

    def info(self, hwnd):
        if hwnd not in self.data:
            raise ToolError("Window closed")
        item = copy.deepcopy(self.data[hwnd])
        dpi = 144 if item["monitor_id"] == 2 else 96
        border = round(8 * dpi / 96)
        r = item["rect"]
        item["frame_rect"] = [r[0] + border, r[1], r[2] - border, r[3] - border] if item.pop("frame_available") else None
        item["window_dpi"] = dpi
        return item

    def monitors(self):
        assert self.depth
        return self.screens

    def foreground(self):
        return self.foreground_hwnd

    def popup(self, hwnd):
        return self.popups.get(hwnd, 0)

    def responsive(self, hwnd):
        return hwnd not in self.hung

    def placement(self, hwnd):
        return copy.deepcopy(self.data[hwnd])

    def restore_placement(self, hwnd, saved):
        self.calls.append(("restore", hwnd))
        self.data[hwnd] = copy.deepcopy(saved)

    def show(self, hwnd, state):
        self.calls.append(("show", hwnd, state))
        self.data[hwnd]["state"] = state

    def position(self, hwnd, bounds):
        assert self.depth
        self.calls.append(("position", hwnd, list(bounds)))
        if self.fail_position == hwnd:
            raise ToolError("simulated positioning failure")
        self.data[hwnd]["rect"] = list(bounds)
        self.data[hwnd]["monitor_id"] = 2 if (bounds[0] + bounds[2]) / 2 < 0 else 1

    def request_foreground(self, hwnd, pid):
        self.calls.append(("focus", hwnd, pid))
        if self.focus_allowed:
            self.foreground_hwnd = hwnd

    def flash(self, hwnd):
        self.calls.append(("flash", hwnd))
