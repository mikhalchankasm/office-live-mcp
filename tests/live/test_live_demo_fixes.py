"""Дефекты, найденные на живых Excel/Word при съёмке демо 2026-10-05 и 2026-10-06. Только с OFFICE_LIVE_LIVE_TESTS=1.

Отмена условного форматирования — в test_live_undo.py.
"""

import pytest

from tests.live.conftest import read
from tests.live.mcpclient import ToolFailed
from tests.live.test_live_stage1 import in_excel, in_word, stage_book, stage_doc, word_state  # noqa: F401 — фикстуры

pytestmark = pytest.mark.live

ROWS = [["Type", "Name", "Mass"], ["Pipe", "P1", 10.5], ["Pipe", "P2", 4], ["Valve", "V1", 2.25], ["Flange", "F1", 1]]


def number_format(book, sheet, getter):
    from office_live.xl_common import read_number_format

    return in_excel(book, lambda app, ws: read_number_format(app, getter(ws.Parent.Worksheets(sheet))))


def test_live_english_number_formats_in_any_locale(srv, stage_book):  # noqa: F811
    # Русский Excel записывал 'dd.mm.yyyy' буквально (ячейка показывала 'dd.mm.yyyy' или ####), а 'General' отвергал.
    book = stage_book
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=[[45934, 1234.5], [45935, 10]])
    for fmt in ("dd.mm.yyyy", "DD.MM.YYYY"):
        srv.call("excel_format_range", workbook=book, sheet="Data", cells="A1:A2", number_format=fmt)
        assert read(srv, book, "A1:A2", display_text=True) == [["04.10.2025"], ["05.10.2025"]]
    srv.call("excel_format_range", workbook=book, sheet="Data", cells="A1", number_format="dd.mm")
    assert read(srv, book, "A1", display_text=True) == [["04.10"]]
    assert srv.call("excel_get_format", workbook=book, sheet="Data", cells="A1")["number_format"] == "dd.mm"
    srv.call("excel_format_range", workbook=book, sheet="Data", cells="B1", number_format="general")
    assert read(srv, book, "B1", display_text=True)[0][0] in {"1234,5", "1234.5"}
    srv.call("excel_format_range", workbook=book, sheet="Data", cells="B1", number_format="[Red]#,##0.00")
    assert srv.call("excel_get_format", workbook=book, sheet="Data", cells="B1")["number_format"] == "[Red]#,##0.00"
    srv.call("excel_format_range", workbook=book, sheet="Data", cells="B2", number_format="ДД.ММ.ГГГГ")  # местная запись по-прежнему работает
    srv.call("office_undo", file=book)

    srv.call("excel_write_range", workbook=book, sheet="Data", cells="D1", values=[["2025-10-04"]], value_mode="text")
    srv.call("excel_clean_text", workbook=book, sheet="Data", cells="D1", operations=["text_to_date"], date_format="YYYY-MM-DD",
             date_number_format="dd.mm.yyyy", preview=False)
    assert number_format(book, "Data", lambda ws: ws.Range("D1")) == "dd.mm.yyyy"
    assert in_excel(book, lambda app, ws: ws.Range("D1").Value2) == 45934


def make_pivots(srv, book):
    srv.call("excel_add_worksheet", workbook=book, name="P")
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=ROWS)
    # две сводные по одному диапазону, но с разными кэшами, как в демо
    srv.call("excel_create_pivot_table", workbook=book, source="A1:C5", source_sheet="Data", rows=["Type", "Name"],
             values=[{"field": "Mass", "function": "sum", "number_format": "dd.mm"}], dest_sheet="P", dest_cell="A3", name="Spec")
    srv.call("excel_create_pivot_table", workbook=book, source="A1:C5", source_sheet="Data", rows=["Type"],
             values=[{"field": "Mass", "function": "sum", "number_format": "#,##0.00"}], dest_sheet="P", dest_cell="G3", name="ByType")


def test_live_pivot_chart_binds_to_the_requested_pivot(srv, stage_book):  # noqa: F811
    book = stage_book
    make_pivots(srv, book)
    assert [d["number_format"] for d in srv.call("excel_pivot_info", workbook=book, pivot="Spec")["data_fields"]] == ["dd.mm"]
    srv.call("excel_select_range", workbook=book, sheet="P", cells="A3")  # активная ячейка — в ДРУГОЙ сводной
    result = srv.call("excel_create_chart", workbook=book, sheet="P", source="ByType", chart_type="bar", name="ch_a",
                      value_axis_number_format="#,##0.0", data_labels=True, data_label_number_format="#,##0.0")
    assert result["chart"] == "ch_a"
    srv.call("excel_create_chart", workbook=book, sheet="P", source="H4", chart_type="bar", name="ch_b", anchor_cell="G20")  # ячейка сводной

    def bound(ws):
        return {co.Name: co.Chart.PivotLayout.PivotTable.Name for co in ws.Parent.Worksheets("P").ChartObjects()}

    assert in_excel(book, lambda app, ws: bound(ws)) == {"ch_a": "ByType", "ch_b": "ByType"}
    # формат диаграммы Excel хранит строкой как есть (разделители местные) и читает его искажённым: проверяем подписи
    labels = in_excel(book, lambda app, ws: [p.DataLabel.Text for p in ws.Parent.Worksheets("P").ChartObjects("ch_a").Chart.SeriesCollection(1).Points()])
    assert sorted(label.strip().replace(",", ".") for label in labels) == ["1.0", "14.5", "2.3"], labels
    assert in_excel(book, lambda app, ws: (app.ActiveSheet.Name, app.Selection.Address)) == ("P", "$A$3")  # выделение возвращено
    srv.call("office_undo", file=book, steps=2)
    assert in_excel(book, lambda app, ws: int(ws.Parent.Worksheets("P").ChartObjects().Count)) == 0


def test_live_slicer_for_pivots_with_different_caches_is_refused_without_leftovers(srv, stage_book):  # noqa: F811
    book = stage_book
    make_pivots(srv, book)
    with pytest.raises(ToolFailed, match="different PivotCache"):
        srv.call("excel_manage_slicers", workbook=book, action="add", pivot="Spec", field="Type", connect_pivots=["ByType"])
    assert srv.call("excel_manage_slicers", workbook=book, action="list")["slicers"] == []


def test_live_word_select_find_text_without_target(srv, doc):
    srv.call("word_insert_text", document=doc, text="Alpha")
    srv.call("word_insert_text", document=doc, text="Total: Beta")
    start, end = srv.call("word_select", document=doc, find_text="Beta")["selected"]
    assert end - start == 4 and start > len("Alpha")  # раньше выделялся абзац 1
    with pytest.raises(ToolFailed, match="does not use find_text"):
        srv.call("word_select", document=doc, target="paragraphs", find_text="Beta")


@pytest.mark.parametrize("hide_column", [False, True])
def test_live_bridge_copies_looks_by_rows(srv, stage_book, stage_doc, hide_column):  # noqa: F811
    # Демо: 44x5 с keep_formatting шло ~41 с — по COM-вызову на каждое свойство каждой ячейки. Одинаковые строки — одним вызовом.
    import time


    book = stage_book
    rows = [["Type", "Name", "Unit", "Qty", "Mass"]] + [[f"T{i % 4}", f"N{i}", "pc", i, i * 1.5] for i in range(43)]
    srv.call("excel_write_range", workbook=book, sheet="Data", cells="A1", values=rows)
    srv.call("excel_format_range", workbook=book, sheet="Data", cells="A1:E1", bold=True)
    srv.call("excel_conditional_format", workbook=book, sheet="Data", cells="A2:E44", rule="formula", formula="=MOD(ROW(),2)=0", fill_color="#DDEBF7")
    srv.call("excel_format_range", workbook=book, sheet="Data", cells="B5", font_color="#C00000")  # неоднородная строка
    if hide_column:
        in_excel(book, lambda app, ws: setattr(ws.Columns(3), "Hidden", True))
    started = time.perf_counter()
    info = srv.call("bridge_excel_range_to_word_table", workbook=book, sheet="Data", cells="A1:E44", document=stage_doc)
    elapsed = time.perf_counter() - started
    cols = 4 if hide_column else 5
    assert info["columns_copied"] == cols and info["cells_with_copied_look"] == cols * 23 + 1, info  # шапка, 22 чётные строки, B5

    def looks(doc):
        t = doc.Tables(info["table"])
        return {"header_bold": int(t.Cell(1, 1).Range.Font.Bold), "even": int(t.Cell(2, cols).Shading.BackgroundPatternColor),
                "odd": int(t.Cell(3, 2).Shading.BackgroundPatternColor), "b5": int(t.Cell(5, 2).Range.Font.Color),
                "a5": int(t.Cell(5, 1).Range.Font.Color)}

    got = in_word(stage_doc, looks)
    assert got["header_bold"] == -1 and got["even"] == 0xF7EBDD and got["odd"] == -16777216, got  # BGR #DDEBF7; авто
    assert got["b5"] == 0x0000C0 and got["a5"] != 0x0000C0, got
    assert elapsed < 30, elapsed


def test_live_mail_merge_shows_no_windows_and_keeps_the_user_document(srv, stage_book, stage_doc, tmp):  # noqa: F811
    # Демо 2026-10-06: на каждый документ слияния Word показывал окно поверх Excel — сначала шаблон с {{метками}}, потом письмо.
    import os
    import subprocess
    import threading

    import win32gui
    import win32process

    def process_name(hwnd):
        pid = win32process.GetWindowThreadProcessId(hwnd)[1]
        out = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"], capture_output=True, text=True, errors="replace").stdout
        return out.split('","')[0].strip('"').lower() if out.startswith('"') else ""

    tpl = srv.call("word_new_document", text="Dear {{Name}}, region {{Region}}.")["document"]
    srv.call("word_headers_footers", document=tpl, action="set", kind="header", text="To {{Name}}")
    tpl_path = os.path.join(tmp, "letter_tpl.docx")
    srv.call("word_save_as", document=tpl, path=tpl_path)
    srv.call("word_close_document", document="letter_tpl.docx", discard=True)
    rows = [["Name", "Region"], ["Ivanov", "North"], ["Petrov", "South"], ["Sidorov", "East"]]
    srv.call("excel_write_range", workbook=stage_book, sheet="Data", cells="A1", values=rows)
    in_word(stage_doc, lambda d: d.Activate())
    before = word_state(stage_doc)

    def word_windows():
        found = set()
        win32gui.EnumWindows(lambda h, _: found.add(h) if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == "OpusApp" else None, None)
        return found

    existing, foreground = word_windows(), win32gui.GetForegroundWindow()
    shown, focus, seen, done = set(), set(), set(), threading.Event()

    def watch():
        while not done.is_set():
            shown.update(word_windows() - existing)
            now = win32gui.GetForegroundWindow()
            if now and now != foreground and now not in seen:
                seen.add(now)
                focus.add((win32gui.GetClassName(now), process_name(now), win32gui.GetWindowText(now)))
            done.wait(0.005)

    watcher = threading.Thread(target=watch)
    watcher.start()
    try:
        out = os.path.join(tmp, "letters")
        r = srv.call("bridge_excel_to_word_documents", workbook=stage_book, sheet="Data", cells="A1:B4", template=tpl_path,
                     output_dir=out, filename_column="Name", export_pdf=True)
    finally:
        done.set()
        watcher.join()
    assert r["files"] == ["Ivanov.docx", "Petrov.docx", "Sidorov.docx"] and len(r["pdf_files"]) == 3 and r["unfilled_placeholders"] == []
    assert not shown, [win32gui.GetWindowText(h) for h in shown if win32gui.IsWindow(h)]
    # Word не выходил на передний план; прочие смены фокуса — рабочий стол пользователя (уведомления, его клики), не слияние
    assert not [f for f in focus if f[0] == "OpusApp" or f[1] == "winword.exe"], focus
    if not focus:
        assert win32gui.GetForegroundWindow() == foreground
    assert word_state(stage_doc) == before  # слияние шло в отдельном Word: документ пользователя не тронут
    assert in_word(stage_doc, lambda d: d.Application.ActiveDocument.FullName) == stage_doc
    names = {d["name"] for d in srv.call("word_list_documents")["documents"]}
    assert not names & {"Ivanov.docx", "Petrov.docx", "Sidorov.docx", "letter_tpl.docx"}


def test_live_mail_merge_from_a_template_open_in_the_users_word(srv, wb, tmp):
    import os

    tpl = srv.call("word_new_document", text="Dear {{Name}}.")["document"]
    tpl_path = os.path.join(tmp, "open_tpl.docx")
    srv.call("word_save_as", document=tpl, path=tpl_path)
    try:
        srv.call("excel_write_range", workbook=wb, sheet="Data", cells="A1", values=[["Name"], ["Ivanov"], ["Petrov"]])
        before = word_state(tpl_path)
        r = srv.call("bridge_excel_to_word_documents", workbook=wb, sheet="Data", cells="A1:A2", template="open_tpl.docx",
                     output_dir=os.path.join(tmp, "out"), filename_column="Name")
        assert r["files"] == ["Ivanov.docx"] and r["unfilled_placeholders"] == []
        assert sorted(os.listdir(os.path.join(tmp, "out"))) == ["Ivanov.docx"]  # без ~$-файлов: документ закрыт
        assert word_state(tpl_path) == before  # открытый шаблон пользователя не тронут
        srv.call("word_open_document", path=os.path.join(tmp, "out", "Ivanov.docx"))
        assert "Dear Ivanov." in srv.call("word_read_document", document="Ivanov.docx")["text"]
        srv.call("word_close_document", document="Ivanov.docx", discard=True)
    finally:
        srv.call("word_close_document", document="open_tpl.docx", discard=True)
