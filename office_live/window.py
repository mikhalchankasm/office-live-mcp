"""Office/chat window policy and layouts; every desktop operation goes through Win32."""

import math
import os

from . import com, wd_common as wd, xl_common as xl
from .errors import ToolError
from .registry import office_tool
from .win32 import Win32

_LAST = []  # Only plain data, never COM references; protected by com.LOCK.
_EXCLUDED = {"excel.exe", "winword.exe", "explorer.exe", "dwm.exe", "shellexperiencehost.exe", "startmenuexperiencehost.exe"}
_SHELL_CLASSES = {"progman", "workerw", "shell_traywnd", "shell_secondarytraywnd", "xlmain", "opusapp"}
_LAYOUTS = {"office_left_chat_right", "chat_left_office_right", "excel_word_side_by_side", "office_maximized"}


def chat_window(api, processes, hint=""):
    candidates = []
    for hwnd in api.windows():
        try:
            info = api.info(hwnd)
        except ToolError:  # disappearing windows are normal during enumeration
            continue
        name = processes.get(info["pid"], {}).get("name", "").lower()
        if (not name or name in _EXCLUDED or info["class"].lower() in _SHELL_CLASSES or not info["visible"]
                or info["cloaked"] or info["owner"] or info["tool_window"] or info["root"] != hwnd or not info["title"]):
            continue
        candidates.append(dict(info, process=name))

    def unique(items, source):
        if len(items) == 1:
            return items[0], source
        return None, f"{source}: {'ambiguous' if items else 'not found'}"

    if hint:
        key = hint.casefold()
        return unique([w for w in candidates if key in w["title"].casefold() or key == w["process"].casefold()], "chat_hint")
    pid, seen = os.getpid(), set()
    while pid and pid not in seen:
        seen.add(pid)
        pid = processes.get(pid, {}).get("parent", 0)
        matches = [w for w in candidates if w["pid"] == pid]
        if matches:
            return unique(matches, "ancestor")
    console = api.console_window()
    matches = [w for w in candidates if w["hwnd"] == console]
    if matches:
        return unique(matches, "inherited_console")
    # WT_SESSION is evidence of a terminal session, not a window/tab identifier.
    # Never choose among multiple Terminal windows and never attach/detach consoles.
    if os.environ.get("WT_SESSION") or console:
        return unique([w for w in candidates if w["process"] == "windowsterminal.exe"], "unique_windows_terminal")
    return None, "No unambiguous chat window; only Office will be arranged. Use chat_hint to clarify."


def is_foreground(api, hwnd):
    fg = api.foreground()
    if not fg:
        return False
    if fg == hwnd:
        return True
    try:
        # Only an ownership relationship; another document in the same process does not count.
        target = api.info(hwnd)
        current = api.info(fg)
        if target["owner"] == fg:
            return True
        seen = set()
        while current["owner"] and current["hwnd"] not in seen:
            seen.add(current["hwnd"])
            if current["owner"] == hwnd:
                return True
            current = api.info(current["owner"])
    except ToolError:
        pass
    return False


def status(api, hwnd, processes, monitors):
    info = api.info(hwnd)
    popup = api.popup(hwnd)
    return dict(info, process=processes.get(info["pid"], {}).get("name"),
                monitor=next((m for m in monitors if m["id"] == info["monitor_id"]), None),
                foreground=is_foreground(api, hwnd), responding=api.responsive(hwnd),
                modal=not info["enabled"], popup_hwnd=popup or None)


def focus(api, hwnd):
    info = api.info(hwnd)
    if not api.responsive(hwnd):
        return {"foreground": is_foreground(api, hwnd), "responding": False, "reason": "Window is not responding; focus was not requested."}
    popup = api.popup(hwnd)
    if not info["enabled"]:
        return {"foreground": is_foreground(api, hwnd), "responding": True, "reason": "A modal dialog blocks the Office window; close it first.", "popup_hwnd": popup or None}
    if not info["visible"] or info["cloaked"]:
        return {"foreground": False, "responding": True, "reason": "Window is hidden or on another virtual desktop."}
    if info["state"] == "minimized":
        api.show(hwnd, "normal")
    api.request_foreground(hwnd, info["pid"])
    foreground = is_foreground(api, hwnd)
    if not foreground:
        api.flash(hwnd)
    return {"foreground": foreground, "responding": True,
            "reason": None if foreground else "Windows denied foreground activation; the taskbar button was flashed."}


def office_target(api, kind, name):
    app, obj = (xl.pick_workbook if kind == "excel" else wd.pick_document)(name, exact_only=True)
    if kind == "excel":
        handles = [int(obj.Windows(i).Hwnd) for i in range(1, obj.Windows.Count + 1)]
        if not handles:
            raise ToolError("The workbook has no window.")
        if len(handles) == 1:
            hwnd = handles[0]
        else:
            active = int(app.ActiveWindow.Hwnd)
            if active not in handles:
                raise ToolError("The workbook has several windows; activate the intended workbook window first.")
            hwnd = active
    else:
        hwnd = int(obj.ActiveWindow.Hwnd)
    # Office COM exposes HWND as a signed 32-bit integer on some versions.
    hwnd &= 0xFFFFFFFF
    info = api.info(hwnd)
    if info["root"] != hwnd or info["class"].lower() != {"excel": "xlmain", "word": "opusapp"}[kind]:
        raise ToolError("COM did not identify a top-level Office window; nothing was changed.")
    busy, detail = None, None
    if kind == "excel" and api.responsive(hwnd):
        try:
            raw = com.raw(app)  # status must not spend the normal busy-retry budget on these getters
            ready, interactive = bool(raw.Ready), bool(raw.Interactive)
            busy, detail = not (ready and interactive), {"ready": ready, "interactive": interactive}
        except Exception:
            detail = {"reason": "Excel Ready/Interactive are unavailable (edit mode, dialog or COM refusal)."}
    return {"kind": kind, "name": str(obj.FullName), "hwnd": hwnd, "busy": busy, "busy_detail": detail}


def _identity(info):
    return info["pid"], info["thread"], info["class"]


def _preflight(api, hwnd):
    info = api.info(hwnd)
    if not info["visible"] or info["cloaked"] or not info["enabled"] or not api.responsive(hwnd):
        raise ToolError("A target window is hidden, modal, on another desktop or not responding; nothing was changed.")
    return info


def _fit(api, hwnd, wanted, warnings):
    if api.info(hwnd)["state"] != "normal":
        api.show(hwnd, "normal")
    # Moving to another DPI monitor changes borders. Measure again there and correct,
    # bounded to three attempts (applications can enforce a minimum window size).
    for _ in range(3):
        info = api.info(hwnd)
        outer, frame = info["rect"], info["frame_rect"]
        if frame is None:
            warnings.append(f"DWM frame bounds unavailable for {hwnd}; using outer rectangle.")
            frame = outer
        bounds = [wanted[i] + outer[i] - frame[i] for i in range(4)]
        api.position(hwnd, bounds)
        actual = api.info(hwnd)
        if (actual["frame_rect"] or actual["rect"]) == wanted:
            return
    warnings.append(f"Window {hwnd} did not fit the requested frame exactly (minimum size, DPI transition or application constraints); see actual rectangles.")


def _restore(api, entries):
    errors, restored = [], []
    for entry in reversed(entries):
        hwnd = entry["hwnd"]
        try:
            if _identity(api.info(hwnd)) != entry["identity"]:
                raise ToolError("Window identity changed; refusing a reused handle.")
            if not api.responsive(hwnd):
                raise ToolError("Window is not responding.")
            api.restore_placement(hwnd, entry["placement"])
            restored.append(hwnd)
        except Exception as exc:
            errors.append({"hwnd": hwnd, "reason": str(exc)})
    return restored, errors


def _readback(api, hwnd, monitors=None):
    try:
        info = api.info(hwnd)
        result = dict(info, foreground=is_foreground(api, hwnd), responding=api.responsive(hwnd),
                      modal=not info["enabled"], popup_hwnd=api.popup(hwnd) or None)
        if monitors is not None:
            result["monitor"] = next((m for m in monitors if m["id"] == info["monitor_id"]), None)
        return result
    except ToolError as exc:
        return {"hwnd": hwnd, "readback_error": str(exc)}


@office_tool("window", "ui", title="Office and chat windows", idempotent=False)
def office_window(action: str = "status", application: str = "excel", workbook: str = "", document: str = "",
                  layout: str = "office_left_chat_right", office_share: float = 1 / 3,
                  monitor: str | int = "chat", chat_hint: str = "", include_chat: bool = True) -> dict:
    """Inspect, focus, arrange or restore already open Office windows. Available in readonly; never edits documents. Arrange only on the user's request/consent. Status is read-only. Chat detection is conservative; an unknown chat is never moved. Restore undoes the last arrange in this server session, not through office_undo.

    Args:
        action: status | focus | arrange | restore.
        application: excel (default) | word, used when no explicit file is given.
        workbook, document: exact open name/path; empty selects active in application. Both required for excel_word_side_by_side.
        layout: office_left_chat_right | chat_left_office_right | excel_word_side_by_side (Excel left) | office_maximized.
        office_share: Office width fraction, 0.25..0.75; Excel fraction in the Excel/Word layout.
        monitor: chat | office | primary | 1-based number from status. Missing chat falls back to Office's monitor.
        chat_hint: unique title substring or exact process name, e.g. Code.exe; never accepts Office or Explorer.
        include_chat: false disables chat discovery/movement, including for live tests. Restore still restores the saved set.
    """
    global _LAST
    if action not in {"status", "focus", "arrange", "restore"} or application not in {"excel", "word"} or layout not in _LAYOUTS:
        raise ToolError("Invalid action, application or layout.")
    if isinstance(office_share, bool) or not isinstance(office_share, (int, float)) or not math.isfinite(office_share) or not 0.25 <= office_share <= 0.75:
        raise ToolError("office_share must be between 0.25 and 0.75.")
    if not (type(monitor) is int and monitor >= 1 or type(monitor) is str and (monitor in {"chat", "office", "primary"} or monitor.isdecimal() and int(monitor) >= 1)):
        raise ToolError("monitor must be chat, office, primary or a positive monitor number.")
    if len(chat_hint) > 256 or any(ord(c) < 32 for c in chat_hint):
        raise ToolError("chat_hint must be a short title substring or process name.")
    pair = action == "arrange" and layout == "excel_word_side_by_side"
    if pair and not (workbook and document) or not pair and workbook and document:
        raise ToolError("Pass both workbook and document only for excel_word_side_by_side; otherwise pass one target.")
    api = Win32()
    with api.dpi_context():
        if action == "restore":
            # Recheck file permissions and current Office HWND before restoring any saved window.
            for entry in _LAST:
                if entry.get("target"):
                    target = entry["target"]
                    from .navigation import renamed_target

                    name = renamed_target(target["kind"], target["name"]) or target["name"]
                    current = office_target(api, target["kind"], name)
                    if current["hwnd"] != entry["hwnd"]:
                        raise ToolError("Office window changed since arrange; restore was refused.")
            # Validate the entire saved set before restoring the first window. Rollback
            # after an arrange failure still uses best effort for surviving windows.
            errors = []
            for entry in _LAST:
                try:
                    if _identity(_preflight(api, entry["hwnd"])) != entry["identity"]:
                        raise ToolError("Window identity changed; refusing a reused handle.")
                except ToolError as exc:
                    errors.append({"hwnd": entry["hwnd"], "reason": str(exc)})
            if errors:
                return {"ok": False, "action": action, "restored": [], "errors": errors, "windows": []}
            restored, errors = _restore(api, _LAST)
            result = {"ok": not errors, "action": action, "restored": restored, "errors": errors,
                      "windows": [_readback(api, h) for h in restored]}
            if not errors:
                _LAST = []
            return result
        targets = ([office_target(api, "excel", workbook), office_target(api, "word", document)] if pair else
                   [office_target(api, "word" if document else "excel" if workbook else application, document or workbook)])
        processes, monitors = api.processes(), api.monitors()
        chat, detection = chat_window(api, processes, chat_hint.strip()) if include_chat else (None, "disabled")
        windows = [dict(status(api, t["hwnd"], processes, monitors), **{k: v for k, v in t.items() if k != "hwnd"}) for t in targets]
        result = {"action": action, "office_windows": windows, "chat_window": status(api, chat["hwnd"], processes, monitors) if chat else None,
                  "chat_detection": detection, "monitors": monitors, "warnings": []}
        if action == "status":
            return result
        if action == "focus":
            result.update(focus(api, targets[0]["hwnd"]))
            result["office_windows"][0].update(_readback(api, targets[0]["hwnd"], monitors))
            return result
        if not monitors:
            raise ToolError("No monitor work area is available.")
        if monitor == "primary":
            chosen = next((m for m in monitors if m["primary"]), None)
        elif monitor in ("office", "chat"):
            origin = chat if monitor == "chat" and chat else windows[0]
            chosen = next((m for m in monitors if m["id"] == origin["monitor_id"]), None)
        else:
            chosen = next((m for m in monitors if m["number"] == int(monitor)), None)
        if chosen is None:
            raise ToolError("The requested monitor is unavailable; run status again.")
        left, top, right, bottom = chosen["work_area"]
        if right <= left or bottom <= top:
            raise ToolError("Invalid monitor work area.")
        width = round((right - left) * office_share)
        office_rect = [left, top, left + width, bottom]
        chat_rect = [left + width, top, right, bottom]
        if layout == "chat_left_office_right":
            office_rect, chat_rect = [right - width, top, right, bottom], [left, top, right - width, bottom]
        moves = [(targets[0]["hwnd"], office_rect, targets[0])]
        if pair:
            moves.append((targets[1]["hwnd"], chat_rect, targets[1]))
        elif layout == "office_maximized":
            moves[0] = (targets[0]["hwnd"], [left, top, right, bottom], targets[0])
        elif chat:
            moves.append((chat["hwnd"], chat_rect, None))
        else:
            result["warnings"].append("Chat window unavailable or disabled; only Office was arranged.")
        if monitor == "chat" and not chat:
            result["warnings"].append("Chat monitor unavailable; using the Office monitor.")
        saved = []
        for hwnd, _, target in moves:
            if target and target["kind"] == "excel" and target["busy"] is not False:
                raise ToolError("Excel is busy or its Ready/Interactive state is unknown; nothing was changed.")
            info = _preflight(api, hwnd)
            saved.append({"hwnd": hwnd, "identity": _identity(info), "placement": api.placement(hwnd), "target": target})
        _LAST = saved
        applied = []
        try:
            for (hwnd, wanted, _), entry in zip(moves, saved, strict=True):
                if _identity(_preflight(api, hwnd)) != entry["identity"]:
                    raise ToolError("Window identity changed before moving; refusing a reused handle.")
                applied.append(hwnd)  # include an operation that may fail after restoring/moving
                _fit(api, hwnd, wanted, result["warnings"])
                if layout == "office_maximized":
                    api.show(hwnd, "maximized")
            result.update(ok=True, applied=applied)
        except Exception as exc:
            restored, errors = _restore(api, [e for e in saved if e["hwnd"] in applied])
            result.update(ok=False, error=str(exc), applied=applied, rolled_back=restored, rollback_errors=errors,
                          note="Use restore to retry restoring the saved placements; office_undo does not manage windows.")
        result["requested"] = [{"hwnd": h, "frame_rect": r} for h, r, _ in moves]
        result["monitor"] = chosen
        result["office_windows"] = [{**w, **_readback(api, w["hwnd"], monitors)} for w in windows]
        if chat:
            result["chat_window"].update(_readback(api, chat["hwnd"], monitors))
        return result
