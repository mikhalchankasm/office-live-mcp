"""Уборка после живых тестов.

Приложения пользователя не трогаем: Word закрывается ТОЛЬКО если его запустили сами тесты (до прогона он не работал)
и в нём не осталось ни одного документа. Если Word был запущен до тестов — он остаётся как есть, даже пустой.
Excel — так же, но по процессам: закрывается только экземпляр, которого не было до прогона, и только без книг.
Иначе тестовый Excel живёт часами между прогонами и упирается в лимит GDI-объектов Windows (10 000 на процесс).
"""

import ctypes
import subprocess

import pythoncom
import win32com.client


def word_running() -> bool:
    """Запущен ли Word сейчас (не запуская его)."""
    pythoncom.CoInitialize()
    try:
        try:
            win32com.client.GetActiveObject("Word.Application")
            return True
        except pythoncom.com_error:
            return False
    finally:
        pythoncom.CoUninitialize()


def quit_word_if_idle(started_by_tests: bool = False) -> str:
    if not started_by_tests:
        return "word not started by the tests - left alone"
    pythoncom.CoInitialize()
    try:
        try:
            w = win32com.client.GetActiveObject("Word.Application")
        except pythoncom.com_error:
            return "word not running"
        if w.Documents.Count == 0:
            w.Quit(0)
            return "word quit (started by the tests and idle)"
        return f"word left running ({w.Documents.Count} documents open)"
    finally:
        pythoncom.CoUninitialize()


def excel_pids() -> set[int]:
    """PID всех процессов EXCEL.EXE сейчас (по tasklist; Excel не запускается и не трогается)."""
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"], capture_output=True, text=True, errors="replace").stdout
    return {int(line.split('","')[1]) for line in out.splitlines() if line.startswith('"EXCEL.EXE"')}


def _pid_of(app) -> int:
    pid = ctypes.c_ulong()
    ctypes.windll.user32.GetWindowThreadProcessId(int(app.Hwnd), ctypes.byref(pid))
    return pid.value


def gdi_objects(pid: int) -> int:
    handle = ctypes.windll.kernel32.OpenProcess(0x0400, False, pid)  # PROCESS_QUERY_INFORMATION
    try:
        return int(ctypes.windll.user32.GetGuiResources(handle, 0)) if handle else 0
    finally:
        if handle:
            ctypes.windll.kernel32.CloseHandle(handle)


def quit_excel_if_idle(pids_before: set[int], gdi_over: int = 0) -> str:
    """Закрывает Excel, ТОЛЬКО если его процесс появился во время прогона и в нём нет ни одной книги.

    gdi_over > 0 — закрывать лишь при утечке: Excel 16 теряет ~120–250 GDI-объектов на каждую пару
    Workbooks.Add/Close, и за один прогон тестовый экземпляр подходит к лимиту Windows (10 000).
    """
    new = excel_pids() - pids_before
    if not new:
        return "no excel started by the tests"
    pythoncom.CoInitialize()
    try:
        try:
            app = win32com.client.GetActiveObject("Excel.Application")
        except pythoncom.com_error:
            return "excel not registered - left alone"
        if _pid_of(app) not in new:
            return "active excel was running before the tests - left alone"
        if app.Workbooks.Count:
            return f"excel left running ({app.Workbooks.Count} workbooks open)"
        if gdi_over and gdi_objects(_pid_of(app)) <= gdi_over:
            return "excel kept (GDI below the threshold)"
        app.Quit()
        return "excel quit (started by the tests and idle)"
    finally:
        pythoncom.CoUninitialize()
