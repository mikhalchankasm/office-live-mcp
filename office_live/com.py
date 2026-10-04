"""COM-слой: инициализация потока, повторы при «Office занят», поиск экземпляров, перевод ошибок.

Три правила, выведенные живыми пробами (см. README «Особенности COM»):
  1. Именованные аргументы в late-binding pywin32 молча превращаются в позиционные по порядку
     (Worksheets.Add(After=x) вставляет ПЕРЕД x). Поэтому все COM-вызовы — только позиционные,
     пропущенные необязательные Variant-аргументы — None (pythoncom.Missing в первой позиции роняет
     остальные аргументы, а для типизированных параметров вроде AutoFilter.Operator None недопустим —
     там передаём явное значение по умолчанию). Прокси ниже запрещает kwargs.
  2. Range.Resize(...)/Offset(...) в late-binding возвращают не то, что нужно — адреса считаем сами.
  3. Объекты COM нельзя отпускать после CoUninitialize: run_com() сначала освобождает прокси
     (в том числе из трейсбека исключения), потом gc.collect(), и только затем CoUninitialize().
"""

import contextlib
import gc
import sys
import threading
import time
import traceback

import pythoncom
import pywintypes
import win32com.client
from win32com.client import CDispatch

from . import config
from .errors import AppBusyError, ToolError

LOCK = threading.RLock()  # весь COM-доступ сериализован: Excel/Word всё равно однопоточны

# HRESULT «приложение занято / отклонило вызов» — вызов можно повторить без побочных эффектов
BUSY = {0x80010001, 0x8001010A, 0x800AC472}
# HRESULT «приложение закрылось / недоступно»
DEAD = {0x80010108, 0x800706BA, 0x800706BE, 0x80010012, 0x800706BF}

_GENERIC_TEXTS = {"", "exception occurred.", "exception occurred", "произошло исключение."}
_KNOWN_SCODES = {
    0x800A03EC: "Excel rejected the operation (error 1004): invalid range/argument/name, protected sheet or workbook, or the object is in an unsupported state",
    0x800A01A8: "Object required (the target object does not exist)",
    0x800A0009: "Subscript out of range (no such sheet/item/index)",
    0x800A000D: "Type mismatch",
    0x800A9C68: "Operation not allowed in the current state",
}


# ------------------------------------------------------------------ ошибки


def hresult_of(exc) -> int:
    h = getattr(exc, "hresult", None)
    if h is None:
        h = exc.args[0] if exc.args else 0
    return (h or 0) & 0xFFFFFFFF


def scode_of(exc) -> int:
    try:
        info = exc.excepinfo
        if info and len(info) > 5 and info[5]:
            return info[5] & 0xFFFFFFFF
    except Exception:
        pass
    return 0


def com_error_text(exc) -> str:
    desc = ""
    try:
        info = exc.excepinfo
        if info and info[2]:
            desc = str(info[2]).strip()
    except Exception:
        pass
    h, s = hresult_of(exc), scode_of(exc)
    if desc.lower() in _GENERIC_TEXTS:
        desc = _KNOWN_SCODES.get(s) or _KNOWN_SCODES.get(h) or (str(exc.args[1]).strip() if len(exc.args) > 1 else "")
    code = f"0x{(s or h):08X}"
    return f"{desc} [COM {code}]" if desc else f"COM error {code}"


def is_busy(exc) -> bool:
    return isinstance(exc, pywintypes.com_error) and (hresult_of(exc) in BUSY or scode_of(exc) in BUSY)


def is_dead(exc) -> bool:
    return isinstance(exc, pywintypes.com_error) and (hresult_of(exc) in DEAD or scode_of(exc) in DEAD)


def translate(exc: BaseException) -> str:
    """Любое исключение -> сообщение для агента."""
    if isinstance(exc, ToolError):
        return str(exc)
    if isinstance(exc, pywintypes.com_error):
        if is_dead(exc):
            return (
                "The Office application was closed or stopped responding while the call was running. "
                "Check that it is running and retry — it will be found again. " + com_error_text(exc)
            )
        return com_error_text(exc)
    if isinstance(exc, (ValueError, TypeError, KeyError, IndexError, AttributeError)):
        print("[office-live] bad input or tool bug:\n" + "".join(traceback.format_exception(exc)), file=sys.stderr)
        return f"{type(exc).__name__}: {exc}"
    if isinstance(exc, OSError):
        return f"File system error: {exc}"
    print("[office-live] unexpected error:\n" + "".join(traceback.format_exception(exc)), file=sys.stderr)
    return f"{type(exc).__name__}: {exc}"


# ------------------------------------------------------------------ прокси с повторами


def _retry(call):
    """Выполняет call(); при «Office занят» ждёт и повторяет только этот вызов (побочных эффектов нет)."""
    deadline = None
    delay = 0.05
    while True:
        try:
            return call()
        except pywintypes.com_error as exc:
            if not is_busy(exc):
                raise
            now = time.monotonic()
            if deadline is None:
                deadline = now + config.SETTINGS.busy_timeout
            if now >= deadline:
                raise AppBusyError(
                    "Excel/Word is busy and did not respond (a dialog is open or a cell/text is being edited). "
                    "Ask the user to press Esc / close the dialog, then retry."
                ) from None
            time.sleep(delay)
            delay = min(delay * 1.6, 1.0)


def _diagnose_attribute_error(o, name: str, exc: AttributeError):
    """CDispatch превращает ЛЮБОЙ сбой GetIDsOfNames в AttributeError, в том числе «приложение занято» и «приложение закрыто».

    Выясняем настоящую причину: если это занятость — пробрасываем com_error (его повторит _retry), иначе остаётся AttributeError.
    """
    try:
        o._oleobj_.GetIDsOfNames(0, name)
    except pywintypes.com_error as com_exc:
        if is_busy(com_exc) or is_dead(com_exc):
            raise com_exc from None
    except AttributeError:
        pass
    raise exc


def _raw(x):
    return object.__getattribute__(x, "_o") if isinstance(x, Proxy) else x


def _wrap(v):
    if isinstance(v, CDispatch):
        return Proxy(v)
    if callable(v) and not isinstance(v, type):
        return _Method(v)
    return v


class _Method:
    __slots__ = ("_f",)

    def __init__(self, f):
        self._f = f

    def __call__(self, *args, **kwargs):
        if kwargs:
            raise TypeError("COM calls must use positional arguments only (named arguments are silently reordered by pywin32)")
        args = tuple(_raw(a) for a in args)
        f = self._f
        return _wrap(_retry(lambda: f(*args)))


class Proxy:
    """Прозрачная обёртка над CDispatch: повтор при «занято», запрет kwargs, разворачивание аргументов-прокси."""

    __slots__ = ("_o",)

    def __init__(self, o):
        object.__setattr__(self, "_o", o)

    def __getattr__(self, name):
        o = object.__getattribute__(self, "_o")
        if name.startswith("_"):
            return getattr(o, name)

        def fetch():
            try:
                return getattr(o, name)
            except AttributeError as exc:
                _diagnose_attribute_error(o, name, exc)

        return _wrap(_retry(fetch))

    def __setattr__(self, name, value):
        o = object.__getattribute__(self, "_o")
        value = _raw(value)

        def store():
            try:
                setattr(o, name, value)
            except AttributeError as exc:
                _diagnose_attribute_error(o, name, exc)

        _retry(store)

    def __call__(self, *args, **kwargs):
        if kwargs:
            raise TypeError("COM calls must use positional arguments only")
        o = object.__getattribute__(self, "_o")
        args = tuple(_raw(a) for a in args)
        return _wrap(_retry(lambda: o(*args)))

    def __getitem__(self, key):
        o = object.__getattribute__(self, "_o")
        return _wrap(_retry(lambda: o[key]))

    def __iter__(self):
        o = object.__getattribute__(self, "_o")
        items = _retry(lambda: list(iter(o)))
        return iter([_wrap(i) for i in items])

    def __eq__(self, other):
        o = object.__getattribute__(self, "_o")
        return o == _raw(other)

    def __ne__(self, other):
        return not self.__eq__(other)

    def __hash__(self):
        return id(object.__getattribute__(self, "_o"))

    def __bool__(self):
        return True

    def __repr__(self):
        return f"<COM {type(object.__getattribute__(self, '_o')).__name__}>"


def raw(x):
    """Прокси -> исходный CDispatch (нужно, например, для win32com.client.Dispatch)."""
    return _raw(x)


# ------------------------------------------------------------------ экземпляры приложений

APPS = {
    "excel": {
        "progid": "Excel.Application",
        "label": "Excel",
        "exts": (".xlsx", ".xlsm", ".xlsb", ".xls", ".xltx", ".xltm", ".xlam", ".csv"),
    },
    "word": {
        "progid": "Word.Application",
        "label": "Word",
        "exts": (".docx", ".docm", ".doc", ".dotx", ".dotm", ".rtf"),
    },
}


def _active_object(progid):
    try:
        return win32com.client.GetActiveObject(progid)
    except pywintypes.com_error:
        return None


def _same_app(a, b) -> bool:
    try:
        if a._oleobj_ == b._oleobj_:
            return True
    except Exception:
        pass
    return False


def _rot_apps(kind: str) -> list:
    """Экземпляры Office, найденные через Running Object Table по открытым файлам (все процессы)."""
    exts = APPS[kind]["exts"]
    found = []
    try:
        ctx = pythoncom.CreateBindCtx(0)
        rot = pythoncom.GetRunningObjectTable()
        monikers = list(rot.EnumRunning())
    except pywintypes.com_error:
        return found
    for mk in monikers:
        try:
            name = mk.GetDisplayName(ctx, None)
        except pywintypes.com_error:
            continue
        if not name or not name.lower().endswith(exts):
            continue
        try:
            unk = rot.GetObject(mk)
            disp = win32com.client.Dispatch(unk.QueryInterface(pythoncom.IID_IDispatch))
            app = disp.Application
        except Exception:
            continue
        if not any(_same_app(app, f) for f in found):
            found.append(app)
    return found


def launch_app(kind: str):
    info = APPS[kind]
    try:
        app = win32com.client.Dispatch(info["progid"])
        app.Visible = True
    except pywintypes.com_error as exc:
        raise ToolError(f"Could not start {info['label']} (is it installed?): " + com_error_text(exc)) from None
    return app


def apps(kind: str, launch: bool = False) -> list:
    """Все запущенные экземпляры приложения (активный в ROT — первым). launch=True — запустить, если нет ни одного."""
    primary = _active_object(APPS[kind]["progid"])
    out = []
    if primary is not None:
        out.append(primary)
    for extra in _rot_apps(kind):
        if not any(_same_app(extra, o) for o in out):
            out.append(extra)
    if not out:
        if not launch:
            label = APPS[kind]["label"]
            raise ToolError(f"{label} is not running. Open {label} (or use the *_open/*_new tool to start it) and retry.")
        out.append(launch_app(kind))
    return [Proxy(a) for a in out]


def primary_app(kind: str, launch: bool = False):
    return apps(kind, launch)[0]


@contextlib.contextmanager
def macros_disabled(app):
    """Макросы и автозапуск отключены на время открытия файла. Не удалось отключить — файл не открываем."""
    try:
        old = app.AutomationSecurity
        app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
    except pywintypes.com_error as exc:
        raise ToolError(
            "Could not disable macros before opening the file (Application.AutomationSecurity), so it was not opened: " + com_error_text(exc)
        ) from None
    try:
        yield
    finally:
        try:
            app.AutomationSecurity = old
        except pywintypes.com_error:
            pass


# ------------------------------------------------------------------ запуск инструмента

_CTX = {"tool": None, "kind": None, "cleanup": [], "targets": []}


def current_kind() -> str | None:
    return _CTX["kind"]


def add_cleanup(fn) -> None:
    """Действие, которое run_com выполнит в конце вызова ДО освобождения COM (например, вернуть EnableEvents)."""
    _CTX["cleanup"].append(fn)


def note_target(text: str) -> None:
    """Запоминает фактически выбранный объект (книга/документ) для журнала аудита."""
    if text not in _CTX["targets"]:
        _CTX["targets"].append(text)


def run_com(fn, args=(), kwargs=None, tool: str | None = None, kind: str | None = None, meta: dict | None = None):
    """Выполняет fn в COM-контексте текущего потока и возвращает результат; ошибки -> ToolError.

    meta (если передан) получает {"targets": [...]} — какие объекты вызов реально выбрал (для аудита).
    """
    message = None
    result = None
    with LOCK:
        pythoncom.CoInitialize()
        _CTX["tool"], _CTX["kind"] = tool, kind
        _CTX["cleanup"], _CTX["targets"] = [], []
        try:
            try:
                result = fn(*args, **(kwargs or {}))
            except Exception as exc:  # noqa: BLE001 — все ошибки превращаем в ToolError ниже, после очистки COM
                message = translate(exc)
            finally:
                for undo in reversed(_CTX["cleanup"]):  # вернуть состояние приложения (события и т. п.), пока COM жив
                    try:
                        undo()
                    except Exception as exc:  # noqa: BLE001
                        print(f"[office-live] cleanup failed: {exc}", file=sys.stderr)
                if meta is not None:
                    meta["targets"] = list(_CTX["targets"])
                _CTX["tool"] = _CTX["kind"] = None
                _CTX["cleanup"], _CTX["targets"] = [], []
        finally:
            # трейсбек исключения уже отпущен (except завершён) — освобождаем прокси ДО CoUninitialize
            gc.collect()
            pythoncom.CoUninitialize()
    if message is not None:
        raise ToolError(message)
    return result
