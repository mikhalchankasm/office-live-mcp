"""Регрессии дефектов, найденных на живых Excel/Word при съёмке демо 2026-10-05 и 2026-10-06 (подделки COM, без Office).

Отмена условного форматирования — в test_undo.py, формат чисел с LCID 1033 — в test_util.py; живые сценарии всех
пунктов — в tests/live/test_live_demo_fixes.py.
"""

from types import SimpleNamespace as NS

import pytest
import pywintypes

from office_live import excel_analysis, excel_pivot, word_core
from office_live.errors import ToolError


def com_error(text="Ошибка"):
    return pywintypes.com_error(-2147467259, text, None, None)


# ------------------------------------------------------------------ word_select


@pytest.fixture
def selected(monkeypatch):
    out = []
    tables = type("Tables", (), {"Count": 2, "__call__": lambda self, i: NS(Range=f"table {i}")})()
    doc = NS(Name="Doc.docx", Tables=tables, Content="content")
    monkeypatch.setattr(word_core, "pick_document", lambda name: ("app", doc))
    monkeypatch.setattr(word_core, "paragraphs_range", lambda d, a, b: f"paragraphs {a}..{b}")
    monkeypatch.setattr(word_core, "find_all", lambda content, text, *args, limit: [f"hit {text}"])
    monkeypatch.setattr(word_core, "select_resolved_range", lambda app, d, rng: out.append(rng) or {"ok": True})
    return out


@pytest.mark.parametrize("kwargs,expected", [
    ({"find_text": "Итого"}, "hit Итого"),  # демо: без target='find' молча выделялся абзац 1
    ({"table_index": 2}, "table 2"),
    ({"start_paragraph": 3}, "paragraphs 3..3"),
    ({}, "paragraphs 1..1"),
    ({"target": "find", "find_text": "x"}, "hit x"),
])
def test_word_select_infers_the_target_from_its_arguments(selected, kwargs, expected):
    word_core.word_select("Doc.docx", **kwargs)
    assert selected == [expected]


@pytest.mark.parametrize("kwargs,message", [
    ({"target": "paragraphs", "find_text": "x"}, "does not use find_text"),
    ({"target": "start", "table_index": 1}, "does not use table_index"),
    ({"find_text": "x", "table_index": 1}, "Pass only one of find_text, table_index"),
    ({"target": "find"}, "needs find_text"),
    ({"target": "middle"}, "target must be"),
])
def test_word_select_refuses_arguments_its_target_would_ignore(selected, kwargs, message):
    with pytest.raises(ToolError, match=message):
        word_core.word_select("Doc.docx", **kwargs)
    assert selected == []


# ------------------------------------------------------------------ сводная диаграмма


class Cells:
    def __init__(self, log, name):
        self.log, self.name = log, name

    def __call__(self, r, c):
        return NS(Select=lambda: self.log.append(f"select {self.name}"))


def pivot_range(log, name, row, column):
    return NS(Row=row, Column=column, Rows=NS(Count=5), Columns=NS(Count=2), Address=f"${name}", Cells=Cells(log, name))


class Chart:
    def __init__(self, log, bound, fail):
        self.log, self.bound, self.fail = log, bound, fail

    def SetSourceData(self, src, *plot_by):
        self.log.append("set source")
        if self.fail:
            raise com_error()
        self.bound = src.Address.lstrip("$")

    def SeriesCollection(self):
        return NS(Count=0)

    @property
    def PivotLayout(self):
        return NS(PivotTable=NS(Name=self.bound))


def chart_book(monkeypatch, *, fail=False, auto_bind="Спецификация"):
    log = []
    pivots = [NS(Name=name, TableRange1=pivot_range(log, name, 3, col), TableRange2=pivot_range(log, name, 1, col)) for name, col in
              (("Спецификация", 1), ("МассаПоТипам", 7))]

    def add_chart(style, code, left, top, width, height):
        log.append("add chart")
        shape = NS(Name="Диаграмма 1", Chart=Chart(log, auto_bind, fail))
        shape.Delete = lambda: log.append("delete chart")
        return shape

    ws = NS(Name="Сводная", PivotTables=lambda: NS(Count=2, __call__=None), Shapes=NS(AddChart2=add_chart),
            Range=lambda addr: NS(Left=10.0, Top=20.0), Activate=lambda: None)
    ws.PivotTables = lambda: type("Pivots", (), {"Count": 2, "__call__": lambda self, j: pivots[j - 1]})()
    wb = NS(Name="Book.xlsx", Worksheets=type("Sheets", (), {"Count": 1, "__call__": lambda self, i: ws})(), Activate=lambda: None)
    app = NS(ActiveWorkbook=wb, ActiveSheet=ws, Selection=NS(Address="$A$3"))
    monkeypatch.setattr(excel_analysis, "pick_workbook", lambda name: (app, wb))
    monkeypatch.setattr(excel_analysis, "_style_chart", lambda *args: None)
    from office_live import excel_format

    monkeypatch.setattr(excel_format, "_restore_view", lambda app, state: log.append("restore view"))
    return log


def test_pivot_chart_selects_its_own_pivot_before_adding(monkeypatch):
    # Демо: активная ячейка в другой сводной — AddChart2 привязывал диаграмму к ней, SetSourceData падал с «Ошибка».
    log = chart_book(monkeypatch)
    result = excel_analysis.excel_create_chart("Book.xlsx", "МассаПоТипам", chart_type="bar")
    assert result["chart"] == "Диаграмма 1"
    assert log[:4] == ["select МассаПоТипам", "add chart", "restore view", "set source"]


def test_failed_pivot_chart_is_removed(monkeypatch):
    log = chart_book(monkeypatch, fail=True)
    with pytest.raises(pywintypes.com_error):
        excel_analysis.excel_create_chart("Book.xlsx", "МассаПоТипам", chart_type="bar")
    assert log[-1] == "delete chart"


def test_chart_bound_to_another_pivot_is_refused_and_removed(monkeypatch):
    log = chart_book(monkeypatch)
    monkeypatch.setattr(Chart, "SetSourceData", lambda self, src, *a: None)  # Excel «принял» данные, но оставил чужую сводную
    with pytest.raises(ToolError, match="bound the new chart to pivot 'Спецификация' instead of pivot 'МассаПоТипам'"):
        excel_analysis.excel_create_chart("Book.xlsx", "МассаПоТипам", chart_type="bar")
    assert log[-1] == "delete chart"


# ------------------------------------------------------------------ срезы


def slicer_book(monkeypatch, caches, fail_connect=False):
    log = []
    pivots = {name: NS(Name=name, CacheIndex=cache, TableRange2=NS(Row=1, Column=1, Columns=NS(Count=2)),
                       PivotFields=lambda: [NS(Name="Тип")]) for name, cache in caches.items()}
    ws = NS(Name="Сводная", Cells=lambda r, c: NS(Left=0.0, Top=0.0))

    def add_pivot(pt):
        log.append(f"connect {pt.Name}")
        if fail_connect:
            raise com_error("1004")

    def add2(src, field, *args):
        log.append("add cache")
        shape = NS(Left=0, Top=0, Width=0, Height=0)
        sc = NS(Name="Срез_Тип", Slicers=NS(Add=lambda *a: NS(Name="Тип", Shape=shape)), PivotTables=NS(AddPivotTable=add_pivot, Count=0))
        sc.Delete = lambda: log.append("delete cache")
        return sc

    wb = NS(Name="Book.xlsx", SlicerCaches=NS(Add2=add2))
    monkeypatch.setattr(excel_pivot, "pick_workbook", lambda name: ("app", wb))
    monkeypatch.setattr(excel_pivot, "find_pivot", lambda book, name="": (ws, pivots[name]))
    monkeypatch.setattr(excel_pivot, "_slicer_info", lambda sc: {"cache": sc.Name})
    return log


def test_slicer_for_pivots_with_different_caches_is_refused_before_creation(monkeypatch):
    log = slicer_book(monkeypatch, {"Спецификация": 1, "МассаПоТипам": 2})
    with pytest.raises(ToolError, match=r"\['МассаПоТипам'\].*different PivotCache.*Nothing was changed"):
        excel_pivot.excel_manage_slicers("Book.xlsx", action="add", pivot="Спецификация", field="Тип", connect_pivots=["МассаПоТипам"])
    assert log == []


def test_slicer_add_is_atomic(monkeypatch):
    log = slicer_book(monkeypatch, {"Спецификация": 1, "МассаПоТипам": 1}, fail_connect=True)
    with pytest.raises(pywintypes.com_error):
        excel_pivot.excel_manage_slicers("Book.xlsx", action="add", pivot="Спецификация", field="Тип", connect_pivots=["МассаПоТипам"])
    assert log == ["add cache", "connect МассаПоТипам", "delete cache"]


def test_slicer_connect_checks_the_cache_first(monkeypatch):
    log = slicer_book(monkeypatch, {"Спецификация": 1, "МассаПоТипам": 2})
    connected = type("Connected", (), {"Count": 1, "__call__": lambda self, k: NS(Name="Спецификация", CacheIndex=1),
                                       "AddPivotTable": lambda self, pt: log.append(f"connect {pt.Name}")})()
    monkeypatch.setattr(excel_pivot, "_find_slicer", lambda wb, name: (NS(Name="Срез_Тип", PivotTables=connected), None))
    with pytest.raises(ToolError, match="different PivotCache"):
        excel_pivot.excel_manage_slicers("Book.xlsx", action="connect", slicer="Тип", connect_pivots=["МассаПоТипам"])
    assert log == []


class _Ole:
    """Ячейка с COM-интерфейсом: Invoke с LCID 1033 сначала отвечает заданными ошибками."""

    def __init__(self, failures):
        self.failures, self.calls = list(failures), []

    def GetIDsOfNames(self, name):  # noqa: N802
        return 1

    def Invoke(self, dispid, lcid, flags, result, *args):  # noqa: N802
        self.calls.append((lcid, args))
        if self.failures:
            raise self.failures.pop(0)
        return "dd.mm.yyyy"


def _com_error(hresult):
    return pywintypes.com_error(hresult - (1 << 32), "x", None, None)


def test_number_format_busy_excel_retries_instead_of_local_fallback(monkeypatch):
    from office_live import config, xl_common
    from dataclasses import replace

    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, busy_timeout=2))
    ole = _Ole([_com_error(0x8001010A)])  # RPC_E_SERVERCALL_RETRYLATER
    cell = NS(_oleobj_=ole, NumberFormat="old")
    xl_common.set_number_format(NS(International={2: ",", 3: " ", 4: ";"}), cell, "dd.mm.yyyy")
    assert [c[0] for c in ole.calls] == [1033, 1033] and cell.NumberFormat == "old"  # локальная запись не понадобилась


def test_number_format_busy_timeout_raises_busy_not_wrong_format(monkeypatch):
    from office_live import config, xl_common
    from office_live.errors import AppBusyError
    from dataclasses import replace

    monkeypatch.setattr(config, "SETTINGS", replace(config.SETTINGS, busy_timeout=0))
    ole = _Ole([_com_error(0x80010001)] * 50)
    cell = NS(_oleobj_=ole, NumberFormat="old")
    with pytest.raises(AppBusyError):
        xl_common.set_number_format(NS(International={2: ",", 3: " ", 4: ";"}), cell, "dd.mm.yyyy")
    assert cell.NumberFormat == "old"


def test_number_format_unsupported_lcid_write_falls_back_to_local(monkeypatch):
    from office_live import xl_common

    ole = _Ole([_com_error(0x80020003)])  # DISP_E_MEMBERNOTFOUND — не «занято»
    cell = NS(_oleobj_=ole, NumberFormat="old")
    monkeypatch.setattr(xl_common, "number_format_for_write", lambda app, fmt, shortcuts=None: "local:" + fmt)
    xl_common.set_number_format(NS(), cell, "0.00")
    assert cell.NumberFormat == "local:0.00"


# ------------------------------------------------------------------ слияние Excel -> Word без окон (демо 2026-10-06)


class MergeDoc:
    def __init__(self, word, template, visible):
        self.word, self.template, self.visible, self.closed = word, template, visible, False

    def SaveAs2(self, path, fmt, /):
        assert fmt == 16 and self.word.DisplayAlerts == 0
        if self.word.fail_save_at == len(self.word.added):
            raise com_error("Диск переполнен")
        open(path, "w").close()

    def ExportAsFixedFormat(self, path, *args):
        open(path, "w").close()

    def Close(self, save, /):
        assert save == 0
        self.closed = True
        self.word.docs.remove(self)


class PrivateWord:
    """Фоновый Word слияния (DispatchEx): невидимый, свой для каждого вызова."""

    def __init__(self, desktop):
        self.desktop, self.docs, self.added, self.quits = desktop, [], [], []
        self.AutomationSecurity, self.DisplayAlerts, self.Visible = 1, True, False
        self.NormalTemplate = NS(Saved=False)
        self.fail_save_at, self.steal_focus = 0, False
        self.Documents = NS(Add=self.add)

    def add(self, template, new_template, doc_type, visible, /):
        assert self.AutomationSecurity == 3 and not self.quits  # макросы шаблона отключены
        doc = MergeDoc(self, template, visible)
        self.docs.append(doc)
        self.added.append(doc)
        if self.steal_focus:  # например, диалог Word вышел на передний план
            self.desktop.foreground_hwnd = 84
        return doc

    def Quit(self, save, /):
        assert self.NormalTemplate.Saved is True  # без вопроса о сохранении Normal.dotm
        self.quits.append(save)
        self.docs.clear()


@pytest.fixture
def merge(monkeypatch, tmp_path):
    from office_live import bridge, com
    from tests.window_fakes import FakeWin32

    desktop = FakeWin32()
    desktop.foreground_hwnd = 42  # пользователь работает в Excel
    word = PrivateWord(desktop)
    rows = [["Name", "Region"], ["Ivanov", "North"], ["Petrov", "South"], ["Sidorov", "East"]]
    monkeypatch.setattr(bridge, "Win32", lambda: desktop)
    launched = []
    monkeypatch.setattr(com, "private_app", lambda kind: launched.append(kind) or word)

    def user_apps(kind, launch=False):  # демо: документы создавались в Word пользователя и показывались поверх Excel
        raise AssertionError(f"the user's {kind} must not be used for the merge")

    monkeypatch.setattr(com, "apps", user_apps)
    monkeypatch.setattr(bridge, "pick_workbook", lambda name: ("excel", "book"))
    monkeypatch.setattr(bridge, "get_range", lambda wb, sheet, cells: ("sheet", NS(Rows=NS(Count=len(rows)))))
    monkeypatch.setattr(bridge, "read_grid", lambda rng, what: rows)

    def fill(doc, values, left, right, scope):
        assert not doc.closed and doc.visible is False
        return {"replaced": {}, "unfilled_placeholders": ["Manager"]}

    monkeypatch.setattr(bridge, "fill_placeholders", fill)
    template = tmp_path / "letter.docx"
    template.write_text("template")

    def run(**kwargs):
        return bridge.bridge_excel_to_word_documents("Book.xlsx", "A1:B4", str(template), str(tmp_path / "out"), filename_column="Name", **kwargs)

    return NS(word=word, desktop=desktop, run=run, launched=launched, out=tmp_path / "out", template=template)


def test_merge_runs_in_a_private_hidden_word_and_quits_it(merge):
    result = merge.run(export_pdf=True)
    assert result["files"] == ["Ivanov.docx", "Petrov.docx", "Sidorov.docx"]
    assert result["pdf_files"] == ["Ivanov.pdf", "Petrov.pdf", "Sidorov.pdf"] and result["unfilled_placeholders"] == ["Manager"]
    assert merge.launched == ["word"] and merge.word.Visible is False and merge.word.quits == [0]
    assert len(merge.word.added) == 3 and all(d.visible is False and d.closed for d in merge.word.added)
    assert all(d.template == str(merge.template) for d in merge.word.added) and merge.word.AutomationSecurity == 1
    assert merge.desktop.foreground_hwnd == 42 and not merge.desktop.calls  # передний план не трогали
    assert sorted(p.name for p in merge.out.iterdir()) == ["Ivanov.docx", "Ivanov.pdf", "Petrov.docx", "Petrov.pdf", "Sidorov.docx", "Sidorov.pdf"]


def test_merge_failure_closes_the_document_quits_word_and_restores_the_foreground(merge):
    merge.word.fail_save_at, merge.word.steal_focus = 2, True
    with pytest.raises(ToolError, match=r"Stopped at data row 2 \(Petrov.docx\).*Created before the failure: 1 document"):
        merge.run()
    assert [d.closed for d in merge.word.added] == [True, True] and merge.word.quits == [0]
    assert merge.desktop.foreground_hwnd == 42 and ("focus", 42, 42) in merge.desktop.calls  # Excel снова впереди
    assert sorted(p.name for p in merge.out.iterdir()) == ["Ivanov.docx"]


def test_merge_does_not_take_the_foreground_back_from_another_app(merge):
    original_add = merge.word.add

    def add_and_switch(*args):  # пользователь сам ушёл в браузер — не возвращаем его в Excel
        merge.desktop.foreground_hwnd = 142
        return original_add(*args)

    merge.word.Documents.Add = add_and_switch
    merge.run()
    assert merge.desktop.foreground_hwnd == 142 and not merge.desktop.calls


def test_merge_without_win32_still_works(merge, monkeypatch):
    from office_live import bridge

    def unavailable():
        raise OSError("user32 unavailable")

    monkeypatch.setattr(bridge, "Win32", unavailable)
    assert merge.run()["documents_created"] == 3 and merge.word.quits == [0]
