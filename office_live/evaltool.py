"""Аварийный люк «полный доступ»: произвольный Python с живыми COM-объектами Excel/Word.

Регистрируется ТОЛЬКО при OFFICE_LIVE_ALLOW_EVAL=1 и не в режиме readonly. Код выполняется в процессе сервера с
правами пользователя — это эквивалент выдачи агенту консоли; включайте осознанно.
"""

import contextlib
import io
import json
import traceback

import pythoncom
import win32com.client

from . import com
from .errors import ToolError
from .registry import office_tool
from .xl_common import pick_workbook


def _jsonable(v):
    try:
        json.dumps(v)
        return v
    except (TypeError, ValueError):
        return repr(v)[:2000]


@office_tool("eval", "destructive", title="Run Python against Office COM (advanced)")
def office_run_python(code: str, workbook: str = "", document: str = "") -> dict:
    """ADVANCED / DANGEROUS: run Python code with live COM objects when no dedicated tool covers a need. Predefined names: `excel` (Excel.Application or None), `wb` (workbook or None), `ws` (its active sheet), `word` (Word.Application or None), `doc` (document or None), `pythoncom`, `win32com`. Set `result = ...` to return a value; print() output is returned too. COM rules: pass arguments POSITIONALLY (no name=value), use None for skipped optional arguments, and avoid Range.Resize/Offset (they misbehave in late binding). Only available when the server was started with OFFICE_LIVE_ALLOW_EVAL=1.

    Args:
        code: Python source to execute.
        workbook: workbook to bind to `wb` ('' = the active one, if any).
        document: Word document to bind to `doc` ('' = the active one, if any).
    """
    ns: dict = {"excel": None, "wb": None, "ws": None, "word": None, "doc": None, "result": None, "pythoncom": pythoncom, "win32com": win32com}
    try:
        apps = com.apps("excel")
        ns["excel"] = apps[0]
        _, ns["wb"] = pick_workbook(workbook)
        ns["ws"] = ns["wb"].ActiveSheet
    except ToolError:
        pass
    try:
        from .wd_common import pick_document

        wapps = com.apps("word")
        ns["word"] = wapps[0]
        _, ns["doc"] = pick_document(document)
    except ToolError:
        pass
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            exec(compile(code, "<office_run_python>", "exec"), ns)  # noqa: S102 — намеренно: инструмент включается явно
    except Exception as exc:  # noqa: BLE001
        if isinstance(exc, ToolError):
            raise
        tb = traceback.format_exc(limit=3)
        raise ToolError(f"{type(exc).__name__}: {exc}\n{tb[-800:]}") from None
    return {"ok": True, "result": _jsonable(ns.get("result")), "stdout": buf.getvalue()[-6000:]}
