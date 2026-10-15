"""Office ownership and recovery for opt-in live tests. Never infer ownership from an image name."""

import ctypes
import gc
import json
import ntpath
import os
import re
import tempfile
import threading
import time
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta, timezone

import pythoncom
import win32api
import win32com.client
import win32gui
import win32process


@dataclass(frozen=True)
class OfficeProcess:
    pid: int
    created: int  # UTC microseconds, same precision as WMI CreationDate / GetProcessTimes
    image: str
    command: str

    @property
    def identity(self):
        return self.pid, self.created

    @property
    def automation(self):
        tokens = [token.strip('"') for token in re.findall(r'"[^"]*"|\S+', self.command.lower())]
        return '/automation' in tokens and '-embedding' in tokens


def timestamp_us(value):
    epoch = datetime(1970, 1, 1, tzinfo=timezone.utc)
    delta = value.astimezone(timezone.utc) - epoch
    return (delta.days * 86400 + delta.seconds) * 1_000_000 + delta.microseconds


def wmi_created(value):
    dt = datetime.strptime(value[:21], "%Y%m%d%H%M%S.%f")
    offset = int(value[22:25]) * (1 if value[21] == '+' else -1)
    return timestamp_us(dt.replace(tzinfo=timezone(timedelta(minutes=offset))))


def office_processes():
    """Read command lines via WMI, without connecting to or starting Office."""
    pythoncom.CoInitialize()
    try:
        service = win32com.client.GetObject("winmgmts:")
        rows = service.ExecQuery("SELECT ProcessId, CreationDate, Name, CommandLine FROM Win32_Process "
                                 "WHERE Name='EXCEL.EXE' OR Name='WINWORD.EXE'")
        # Missing command line => foreign; missing identity => fail closed (abort initialization/poll).
        return [OfficeProcess(int(row.ProcessId), wmi_created(str(row.CreationDate)),
                              str(row.Name).upper(), str(row.CommandLine or '')) for row in rows]
    finally:
        pythoncom.CoUninitialize()


def excel_pids():
    return {p.pid for p in office_processes() if p.image == 'EXCEL.EXE'}


_EVENT_LOCK = threading.Lock()


def live_event(event, **fields):
    entry = {"event": event, "time": time.time(), **fields}
    path = os.environ.get('OFFICE_LIVE_LIVE_EVENTS')
    if path:
        with _EVENT_LOCK, open(path, 'a', encoding='utf-8') as stream:
            stream.write(json.dumps(entry, ensure_ascii=False) + '\n')
    return entry


class Ownership:
    def __init__(self, started, before):
        self.started = started
        self.foreign = {p.identity for p in before}
        self.owned = {}
        self.lock = threading.RLock()

    def observe(self, processes):
        with self.lock:
            for p in processes:
                if p.image in ('EXCEL.EXE', 'WINWORD.EXE') and p.identity not in self.foreign and p.created > self.started and p.automation:
                    if p.identity not in self.owned:
                        self.owned[p.identity] = p
                        live_event('owned', process=asdict(p))

    def claim_leftover(self, p):
        with self.lock:
            self.foreign.discard(p.identity)
            self.owned[p.identity] = p

    def owns(self, p):
        with self.lock:
            return p.identity in self.owned


# Set only by the live session fixture. Older cleanup callers cannot bypass ownership.
OWNERSHIP = None


def handle_created_us(handle):
    """Keep FILETIME precision: pywin32's datetime conversion truncates to milliseconds."""
    from ctypes import wintypes

    get_times = ctypes.windll.kernel32.GetProcessTimes
    get_times.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
    get_times.restype = wintypes.BOOL
    values = [wintypes.FILETIME() for _ in range(4)]
    if not get_times(int(handle), *(ctypes.byref(value) for value in values)):
        raise ctypes.WinError()
    created = values[0]
    ticks = (created.dwHighDateTime << 32) | created.dwLowDateTime
    return (ticks - 116444736000000000) // 10  # 1601 -> 1970; 100ns -> microseconds (WMI precision)


def terminate_owned(process, ownership, *, open_process=None, creation_time=None, terminate=None):
    """Check identity through the SAME handle used for termination (including on 64-bit Windows)."""
    if not ownership.owns(process):
        return False
    open_process = open_process or win32api.OpenProcess
    creation_time = creation_time or handle_created_us
    terminate = terminate or win32api.TerminateProcess
    try:
        handle = open_process(0x0001 | 0x1000, False, process.pid)  # TERMINATE | QUERY_LIMITED_INFORMATION
    except OSError:
        return False
    try:
        if creation_time(handle) != process.created:
            return False
        terminate(handle, 1)
        return True
    finally:
        handle.Close()


def _office_windows(dialog=False):
    from ctypes import wintypes

    is_hung = ctypes.windll.user32.IsHungAppWindow
    is_hung.argtypes, is_hung.restype = [wintypes.HWND], wintypes.BOOL
    names = {p.pid: p.image for p in office_processes()}
    found = []

    def visit(hwnd, _):
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if pid not in names or (dialog and not win32gui.IsWindowVisible(hwnd)):
            return
        title = win32gui.GetWindowText(hwnd)
        if dialog:
            cls = win32gui.GetClassName(hwnd)
            if cls in ('NUIDialog', '#32770'):
                texts = [title]

                def child(ch, _):
                    text = win32gui.GetWindowText(ch)
                    if text:
                        texts.append(text)

                try:
                    win32gui.EnumChildWindows(hwnd, child, None)
                except win32gui.error:
                    pass  # a disappearing child must not hide the dialog itself
                found.append((pid, cls, ' | '.join(dict.fromkeys(texts))))
        elif is_hung(hwnd):
            found.append((names[pid], pid, title))

    win32gui.EnumWindows(visit, None)
    return found


def hung_office_windows():
    return _office_windows()


def office_dialogs():
    """Visible Office dialogs: (PID, class, title + available child text); read only."""
    return _office_windows(dialog=True)


class OfficeWatch:
    """One synchronous watchdog tick, with injectable observations/actions for Office-free tests."""
    def __init__(self, ownership, kill_clients, hung_for=45, dialog_for=20, terminate=None, emit=None):
        self.ownership, self.kill_clients = ownership, kill_clients
        self.hung_for, self.dialog_for = hung_for, dialog_for
        self.terminate = terminate or terminate_owned
        self.emit = emit or live_event
        self.since = {}
        self.recoveries = []
        self.fatal = ''
        self.handled = set()
        self.lock = threading.RLock()

    def checkpoint(self):
        # Pytest must not start a new MCP server between terminating Office and killing old servers.
        with self.lock:
            return self.fatal, list(self.recoveries)

    def tick(self, now, processes, hung, dialogs):
        with self.lock:
            self._tick(now, processes, hung, dialogs)

    def _tick(self, now, processes, hung, dialogs):
        self.ownership.observe(processes)
        by_pid = {p.pid: p for p in processes}
        observations = [(pid, 'hung', text) for _, pid, text in hung]
        observations += [(pid, 'dialog', text) for pid, _, text in dialogs]
        current = {}
        for pid, reason, text in observations:
            p = by_pid.get(pid)
            if p is None or p.identity in self.handled:
                continue
            key = (p.identity, reason)
            current[key] = self.since.get(key, now)
            threshold = self.hung_for if reason == 'hung' else self.dialog_for
            if now - current[key] < threshold:
                continue
            message = f'{p.image} PID {pid}: {reason} >= {threshold}s; {text}'
            if not self.ownership.owns(p):
                self.fatal = f'Чужой Office: {message}. Процесс не закрыт; прогон остановлен.'
                self.emit('foreign_office', process=asdict(p), reason=message, text=text)
                self.kill_clients()
                return
            self.handled.add(p.identity)
            # Log evidence BEFORE any destructive action, including a blocked Quit in pytest itself.
            self.emit('recovery_start', process=asdict(p), reason=message, text=text)
            try:
                terminated = self.terminate(p, self.ownership)
            except Exception as exc:
                terminated = False
                message += f'; termination failed: {exc}'
            if not terminated:
                self.fatal = f'Не удалось безопасно закрыть свой Office: {message}'
                self.emit('recovery_failed', process=asdict(p), reason=self.fatal)
            else:
                self.recoveries.append(message)
                self.emit('recovery', process=asdict(p), reason=message, text=text, number=len(self.recoveries))
                if len(self.recoveries) >= 3:
                    self.fatal = 'Третье восстановление Office; прогон остановлен: ' + '; '.join(self.recoveries)
            self.kill_clients()
            # One recovery per tick, not one per window of the same process.
            break
        self.since = current


def is_test_document(path, temp_root=None):
    if not path or not ntpath.isabs(path):
        return False
    root = ntpath.normcase(ntpath.abspath(temp_root or tempfile.gettempdir()))
    path = ntpath.normcase(ntpath.abspath(path))
    try:
        if ntpath.commonpath([root, path]) != root:
            return False
    except ValueError:
        return False
    parts = ntpath.relpath(path, root).split('\\')
    return len(parts) >= 2 and parts[0].startswith('ol_pytest_')


def check_leftover(process, paths, ownership, emit=live_event):
    """A complete document inventory is required before claiming an old automation process."""
    if not process.automation:
        return False
    if any(not is_test_document(path) for path in paths):
        message = f'{process.image} PID {process.pid}: automation содержит чужой документ; ничего не закрыто: {paths}'
        emit('leftover_refused', process=asdict(process), reason=message)
        raise RuntimeError(message)
    ownership.claim_leftover(process)
    emit('leftover', process=asdict(process), documents=paths)
    return True


def _application_for(process):
    # Existing ROT objects only; apps(..., launch=False) never launches Office.
    from office_live.com import apps

    kind = 'excel' if process.image == 'EXCEL.EXE' else 'word'
    before = {p.identity for p in office_processes() if p.image == process.image}
    if process.identity not in before:
        raise RuntimeError(f'Office PID {process.pid} изменился до подключения к COM')
    for app in apps(kind, launch=False):
        try:
            hwnd = app.Hwnd if kind == 'excel' else app.ActiveWindow.Hwnd
            pid = win32process.GetWindowThreadProcessId(int(hwnd) & 0xFFFFFFFF)[1]
            if pid == process.pid:
                if process.identity in {p.identity for p in office_processes() if p.image == process.image}:
                    return app
        except Exception:
            # Word.Application has no Hwnd. An empty Word can be bound only when it is
            # the sole Word instance BOTH before and after obtaining the existing COM object.
            # Ownership is still separately checked using automation flags + creation time.
            if kind == 'word' and int(app.Documents.Count) == 0 and before == {process.identity}:
                after = {p.identity for p in office_processes() if p.image == process.image}
                if after == before:
                    return app
    raise RuntimeError(f'Не удалось проверить документы {process.image} PID {process.pid}; процесс оставлен без изменений')


TEST_MARK = 'ol_pytest_'  # prefix of RUN_TAG: workbook names / Word Comments set by the live fixtures


def own_unsaved(doc, image) -> bool:
    """An unsaved document created by our fixtures: Excel book with an ol_pytest_* defined name or the undo service
    book, Word document whose Comments property starts with ol_pytest_. Anything unreadable counts as foreign."""
    try:
        if image == 'EXCEL.EXE':
            if _is_undo_book(doc):
                return True
            names = doc.Names
            return any(str(names(i).Name).startswith(TEST_MARK) for i in range(1, int(names.Count) + 1))
        return str(doc.BuiltInDocumentProperties(5).Value).startswith(TEST_MARK)  # 5 = wdPropertyComments
    except Exception:  # noqa: BLE001 — fail closed: not provably ours
        return False


def clean_leftovers(before, ownership, inspect_and_quit=None, terminate=None, wait_seconds=15):
    """Bound COM inspection/Quit, including broken old ROT objects; unknown documents fail closed."""
    terminate = terminate or terminate_owned
    for process in before:
        if not process.automation:
            continue
        done = threading.Event()
        errors = []
        cancelled = threading.Event()
        claim_lock = threading.Lock()
        claimed_at = []

        def clean(p=process, done=done, errors=errors, cancelled=cancelled, claim_lock=claim_lock, claimed_at=claimed_at):
            pythoncom.CoInitialize()
            app = docs = doc = None
            try:
                if inspect_and_quit is not None:
                    inspect_and_quit(p, ownership)
                    return
                app = _application_for(p)
                # Revalidate the identity before inspecting/claiming a ROT application.
                if p.identity not in {v.identity for v in office_processes()}:
                    raise RuntimeError(f'Office PID {p.pid} изменился при проверке остатка')
                collection = app.Workbooks if p.image == 'EXCEL.EXE' else app.Documents
                docs = [collection(i) for i in range(1, int(collection.Count) + 1)]
                # Unsaved books of our own fixtures (marked ol_pytest_*) and the undo service book are ours too;
                # any other unsaved document stays foreign ('' is never a test document).
                paths = [str(doc.FullName) if str(doc.Path) else '' for doc in docs if str(doc.Path) or not own_unsaved(doc, p.image)]
                with claim_lock:
                    if cancelled.is_set():
                        return
                    check_leftover(p, paths, ownership)
                    claimed_at.append(time.monotonic())
                for doc in docs:
                    doc.Close(False if p.image == 'EXCEL.EXE' else 0)
                app.Quit() if p.image == 'EXCEL.EXE' else app.Quit(0)
            except Exception as exc:
                errors.append(exc)
            finally:
                app = docs = doc = collection = None
                gc.collect()
                pythoncom.CoUninitialize()
                done.set()

        thread = threading.Thread(target=clean, daemon=True, name='office-leftover')
        thread.start()
        deadline = time.monotonic() + wait_seconds
        done.wait(wait_seconds)
        with claim_lock:
            if not done.is_set() and not ownership.owns(process):
                cancelled.set()
                raise RuntimeError(f'Не удалось проверить документы {process.image} PID {process.pid} за {wait_seconds} с; процесс не закрыт')
            if claimed_at:
                deadline = claimed_at[0] + wait_seconds
        if errors and not ownership.owns(process):
            raise RuntimeError(f'Проверка остатка Office PID {process.pid}: {errors[0]}') from errors[0]
        while time.monotonic() < deadline:
            if process.identity not in {p.identity for p in office_processes()}:
                break
            time.sleep(min(0.25, max(0, deadline - time.monotonic())))
        else:
            if process.identity in {p.identity for p in office_processes()}:
                if not terminate(process, ownership):
                    raise RuntimeError(f'Не удалось безопасно закрыть остаток Office PID {process.pid}')
                live_event('leftover_terminated', process=asdict(process))
        # A stuck worker must never be allowed to mutate a later Office instance.
        if not done.wait(1):
            raise RuntimeError(f'COM очистки остатка PID {process.pid} не завершился после закрытия процесса')
        live_event('leftover_closed', process=asdict(process))


class DialogWatch:
    """Опрашивает office_dialogs() в фоне, пока открыт блок with; увиденные окна — в .seen."""

    def __init__(self, every: float = 0.2):
        self.every, self.seen = every, []
        self._stop = threading.Event()

    def _run(self):
        while not self._stop.wait(self.every):
            for item in office_dialogs():
                if item not in self.seen:
                    self.seen.append(item)

    def __enter__(self):
        self._thread = threading.Thread(target=self._run, daemon=True, name="dialog-watch")
        self._thread.start()
        return self

    def __exit__(self, *exc):
        self._stop.set()
        self._thread.join(timeout=5)
        return False


def gdi_objects(pid: int) -> int:
    handle = win32api.OpenProcess(0x0400, False, pid)  # PROCESS_QUERY_INFORMATION
    try:
        return int(ctypes.windll.user32.GetGuiResources(ctypes.c_void_p(int(handle)), 0))
    finally:
        handle.Close()


def working_set_mb(pid: int) -> int:
    """Рабочий набор процесса в МБ (GetProcessMemoryInfo; Excel не трогается)."""
    class Counters(ctypes.Structure):
        _fields_ = [("cb", ctypes.c_ulong), ("PageFaultCount", ctypes.c_ulong)] + [
            (name, ctypes.c_size_t) for name in ("PeakWorkingSetSize", "WorkingSetSize", "QuotaPeakPagedPoolUsage",
                                                 "QuotaPagedPoolUsage", "QuotaPeakNonPagedPoolUsage", "QuotaNonPagedPoolUsage",
                                                 "PagefileUsage", "PeakPagefileUsage")]
    handle = win32api.OpenProcess(0x0410, False, pid)  # QUERY_INFORMATION | VM_READ
    try:
        counters = Counters()
        counters.cb = ctypes.sizeof(Counters)
        ok = ctypes.windll.psapi.GetProcessMemoryInfo(ctypes.c_void_p(int(handle)), ctypes.byref(counters), counters.cb)
        return int(counters.WorkingSetSize // (1024 * 1024)) if ok else 0
    finally:
        handle.Close()


def _is_undo_book(book) -> bool:
    try:
        book.Names("__OfficeLiveUndo")  # служебная скрытая книга снимков отмены (office_live.undo.MARKER)
        return True
    except pythoncom.com_error:
        return False


def quit_word_if_idle():
    return _quit_idle('WINWORD.EXE')


def quit_excel_if_idle(gdi_over=0, mem_over_mb=0):
    return _quit_idle('EXCEL.EXE', gdi_over, mem_over_mb)


def _quit_idle(image, gdi_over=0, mem_over_mb=0):
    if OWNERSHIP is None:
        return 'ownership unavailable - left alone'
    processes = office_processes()
    OWNERSHIP.observe(processes)
    for process in processes:
        if process.image != image or not OWNERSHIP.owns(process):
            continue
        if (gdi_over or mem_over_mb) and gdi_objects(process.pid) <= gdi_over and working_set_mb(process.pid) <= mem_over_mb:
            continue
        pythoncom.CoInitialize()
        app = docs = doc = collection = None
        try:
            app = _application_for(process)
            collection = app.Workbooks if image == 'EXCEL.EXE' else app.Documents
            docs = [collection(i) for i in range(1, int(collection.Count) + 1)]
            if docs and (image != 'EXCEL.EXE' or any(not _is_undo_book(doc) for doc in docs)):
                continue
            for doc in docs:
                doc.Close(False)
            app.Quit() if image == 'EXCEL.EXE' else app.Quit(0)
        finally:
            app = docs = doc = collection = None
            gc.collect()
            pythoncom.CoUninitialize()
        deadline = time.monotonic() + 15
        while process.identity in {p.identity for p in office_processes()}:
            if time.monotonic() >= deadline:
                # Still our process, verified again via the termination handle.
                if terminate_owned(process, OWNERSHIP):
                    live_event('idle_terminated', process=asdict(process))
                else:
                    raise RuntimeError(f'Не удалось безопасно закрыть свой Office PID {process.pid} после Quit')
                break
            time.sleep(0.25)
    return 'owned idle Office cleanup finished'
