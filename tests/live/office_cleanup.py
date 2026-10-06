"""Уборка после живых тестов.

Приложения пользователя не трогаем: Word закрывается ТОЛЬКО если его запустили сами тесты (до прогона он не работал)
и в нём не осталось ни одного документа. Если Word был запущен до тестов — он остаётся как есть, даже пустой.
Excel — так же, но по процессам: закрывается только экземпляр, которого не было до прогона, и только без книг.
Иначе тестовый Excel живёт часами между прогонами и упирается в лимит GDI-объектов Windows (10 000 на процесс).
"""

import ctypes
import gc
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
        count = w.Documents.Count
        if count == 0:
            w.Quit(0)
        del w  # ссылку отпускаем ДО CoUninitialize: иначе Word прячется, но процесс остаётся жить
        return "word quit (started by the tests and idle)" if count == 0 else f"word left running ({count} documents open)"
    finally:
        gc.collect()
        pythoncom.CoUninitialize()


def excel_pids() -> set[int]:
    """PID всех процессов EXCEL.EXE сейчас (по tasklist; Excel не запускается и не трогается)."""
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"], capture_output=True, text=True, errors="replace").stdout
    return {int(line.split('","')[1]) for line in out.splitlines() if line.startswith('"EXCEL.EXE"')}


def hung_office_windows() -> list[tuple[str, int, str]]:
    """Видимые окна Excel/Word, которые Windows считает зависшими (IsHungAppWindow: не отвечают ~5 с). Только чтение."""
    from ctypes import wintypes

    names = {}
    for image in ("EXCEL.EXE", "WINWORD.EXE"):
        out = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {image}", "/FO", "CSV", "/NH"], capture_output=True, text=True, errors="replace").stdout
        names.update({int(line.split('","')[1]): image for line in out.splitlines() if line.startswith(f'"{image}"')})
    user32, hung = ctypes.windll.user32, []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def visit(hwnd, _):
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value in names and user32.IsWindowVisible(hwnd) and user32.IsHungAppWindow(hwnd):
            title = ctypes.create_unicode_buffer(200)
            user32.GetWindowTextW(hwnd, title, 200)
            hung.append((names[pid.value], pid.value, title.value))
        return True

    user32.EnumWindows(visit, 0)
    return hung


def hidden_idle_excel_pids() -> set[int]:
    """Скрытый Excel без единой книги — остаток прошлого прогона (Excel пользователя видим): его можно считать тестовым.

    Иначе он попадал в «Excel пользователя», не перезапускался и между прогонами раздувался до 1 ГБ и зависания."""
    pythoncom.CoInitialize()
    app = None
    try:
        try:
            app = win32com.client.GetActiveObject("Excel.Application")
        except pythoncom.com_error:
            return set()
        if bool(app.Visible) or int(app.Workbooks.Count):
            return set()
        # Остаток прошлого прогона (Quit, пока сервер тестов держал ссылки): окно спрятано, процесс жив, и тесты
        # окон получают «скрытое окно». Книг нет — делаем видимым, дальше он закрывается как тестовый.
        app.Visible = True
        return {_pid_of(app)}
    finally:
        app = None  # noqa: F841 — ссылки отпускаем до CoUninitialize, иначе процесс не завершится после Quit
        gc.collect()
        pythoncom.CoUninitialize()


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


def working_set_mb(pid: int) -> int:
    """Рабочий набор процесса в МБ (GetProcessMemoryInfo; Excel не трогается)."""
    class Counters(ctypes.Structure):
        _fields_ = [("cb", ctypes.c_ulong), ("PageFaultCount", ctypes.c_ulong)] + [
            (name, ctypes.c_size_t) for name in ("PeakWorkingSetSize", "WorkingSetSize", "QuotaPeakPagedPoolUsage",
                                                 "QuotaPagedPoolUsage", "QuotaPeakNonPagedPoolUsage", "QuotaNonPagedPoolUsage",
                                                 "PagefileUsage", "PeakPagefileUsage")]
    handle = ctypes.windll.kernel32.OpenProcess(0x0410, False, pid)  # QUERY_INFORMATION | VM_READ
    if not handle:
        return 0
    try:
        counters = Counters()
        counters.cb = ctypes.sizeof(Counters)
        ok = ctypes.windll.psapi.GetProcessMemoryInfo(handle, ctypes.byref(counters), counters.cb)
        return int(counters.WorkingSetSize // (1024 * 1024)) if ok else 0
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def _is_undo_book(book) -> bool:
    try:
        book.Names("__OfficeLiveUndo")  # служебная скрытая книга снимков отмены (office_live.undo.MARKER)
        return True
    except pythoncom.com_error:
        return False


def quit_excel_if_idle(pids_before: set[int], gdi_over: int = 0, mem_over_mb: int = 0) -> str:
    """Закрывает Excel, ТОЛЬКО если его процесс появился во время прогона и в нём нет книг, кроме служебной книги отмены.

    Пороги (gdi_over/mem_over_mb > 0) — закрывать лишь при разбухании: Excel 16 теряет память и ~120–250 GDI-объектов на
    каждую пару Workbooks.Add/Close, а почти каждый тест создаёт свою книгу. Без порогов — закрывать всегда (конец прогона).
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
        pid = _pid_of(app)
        if pid not in new:
            return "active excel was running before the tests - left alone"
        books = [app.Workbooks(i) for i in range(1, app.Workbooks.Count + 1)]
        if any(not _is_undo_book(b) for b in books):
            return f"excel left running ({len(books)} workbooks open)"
        if (gdi_over or mem_over_mb) and gdi_objects(pid) <= gdi_over and working_set_mb(pid) <= mem_over_mb:
            return "excel kept (below the thresholds)"
        for book in books:  # снимки отмены закрытых тестовых книг больше не нужны; без этого Quit спросил бы о сохранении
            book.Close(False)
        app.Quit()
        return "excel quit (started by the tests and idle)"
    finally:
        # все ссылки на Excel отпускаем ДО CoUninitialize: иначе после Quit он прячется, но процесс живёт (~400 МБ)
        app = books = book = None  # noqa: F841
        gc.collect()
        pythoncom.CoUninitialize()
