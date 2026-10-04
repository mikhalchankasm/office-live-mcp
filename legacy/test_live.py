# Живой тест office-live: Excel (запись/формула/формат/autofit/save) + Word.
# Excel поднимаем сами (Dispatch + Visible), ждём регистрации в ROT,
# затем ВСЕ проверки идут через инструменты сервера (GetActiveObject-путь).
# Финал: пересоздаём тест MCP_Test через office-live и сохраняем ZCode_MCP_Test.xlsx.
import os
import time

import pythoncom
import win32com.client

import server

results = []


def step(name, fn):
    try:
        out = fn()
        results.append((name, "OK"))
        print(f"[OK]   {name}: {str(out)[:200]}")
        return out
    except Exception as e:
        results.append((name, f"ERROR: {e}"))
        print(f"[ERR]  {name}: {e}")
        return None


def wait_rot(progid, timeout=15):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            return win32com.client.GetActiveObject(progid)
        except pythoncom.com_error:
            time.sleep(0.5)
    raise RuntimeError(f"{progid} не зарегистрировался в ROT за {timeout}s")


def reattach(progid):
    """Свежий прокси после server.*-вызовов: их CoUninitialize рвёт старые прокси потока."""
    return win32com.client.GetActiveObject(progid)


MATRIX = [
    [44, 35, 100, 7, 72, 40, 78, 73, 90, 8],
    [69, 17, 22, 62, 66, 2, 14, 29, 94, 55],
    [11, 53, 54, 60, 36, 80, 33, 78, 20, 78],
    [89, 3, 91, 42, 73, 88, 57, 97, 82, 23],
    [44, 94, 28, 77, 70, 87, 31, 70, 3, 6],
    [2, 9, 33, 80, 11, 2, 27, 31, 96, 13],
    [32, 11, 32, 38, 74, 29, 6, 9, 62, 37],
    [77, 41, 59, 50, 76, 77, 51, 65, 47, 72],
    [6, 25, 52, 53, 56, 67, 86, 36, 12, 3],
    [35, 81, 3, 9, 16, 100, 21, 16, 25, 18],
]
EXPECT_SUMS = [409, 369, 474, 478, 550, 572, 404, 504, 531, 313]


def main():
    pythoncom.CoInitialize()

    # --- Excel: поднимаем и ждём ROT ---
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    for i in range(1, excel.Workbooks.Count + 1):
        if excel.Workbooks(i).Name == "scratch-test.xlsx":
            excel.Workbooks(i).Close(SaveChanges=0)

    # чистим артефакты прошлых прогонов
    for f in (r"C:\path\to\office-live-mcp\scratch-test.xlsx", r"C:\path\to\office-live-mcp\scratch-test.docx"):
        if os.path.exists(f):
            os.remove(f)

    scratch = excel.Workbooks.Add()
    scratch_name = scratch.Name
    print(f"excel up, scratch: {scratch_name}")
    wait_rot("Excel.Application")
    print("ROT ok")

    # --- прогон инструментов на черновике ---
    step("list_workbooks", lambda: server.excel_list_workbooks())
    step("add_worksheet", lambda: server.excel_add_worksheet(scratch_name, "Live_Test"))
    step("write_range", lambda: server.excel_write_range(scratch_name, "Live_Test", "B2:C3", [[11, 22], [33, 44]]))
    step("set_formula", lambda: server.excel_set_formula(scratch_name, "Live_Test", "B5", "=SUM(B2:C3)"))
    step("format_range", lambda: server.excel_format_range(scratch_name, "Live_Test", "B5", bold=True, fill_color="#FFF2CC"))
    step("autofit", lambda: server.excel_autofit(scratch_name, "Live_Test", "columns"))
    step("save_as_scratch", lambda: server.excel_save_as(scratch_name, r"C:\path\to\office-live-mcp\scratch-test.xlsx"))
    rd = step("read_back", lambda: server.excel_read_range("scratch-test.xlsx", "Live_Test", "B2:C5"))
    if rd:
        match = rd["values"] == [[11, 22], [33, 44], [None, None], [110, None]]
        print("READBACK MATCH:", match, "| got:", rd["values"])
        results.append(("readback_match", "OK" if match else "MISMATCH"))
    reattach("Excel.Application").Workbooks.Item("scratch-test.xlsx").Close(SaveChanges=0)

    # --- полный круг: MCP_Test через office-live -> ZCode_MCP_Test.xlsx ---
    excel = reattach("Excel.Application")
    book = excel.Workbooks.Add()
    book_name = book.Name
    step("MCP:add_sheet", lambda: server.excel_add_worksheet(book_name, "MCP_Test"))
    step("MCP:write_10x10", lambda: server.excel_write_range(book_name, "MCP_Test", "A1:J10", MATRIX))
    for col in "ABCDEFGHIJ":
        step(f"MCP:sum_{col}12",
             lambda c=col: server.excel_set_formula(book_name, "MCP_Test", f"{c}12", f"=SUM({c}1:{c}10)"))
    step("MCP:bold", lambda: server.excel_format_range(book_name, "MCP_Test", "A12:J12", bold=True))
    step("MCP:autofit", lambda: server.excel_autofit(book_name, "MCP_Test", "columns"))
    sums = step("MCP:read_sums", lambda: server.excel_read_range(book_name, "MCP_Test", "A12:J12"))
    if sums:
        match = sums["values"][0] == EXPECT_SUMS
        print("SUMS MATCH:", match, "| got:", sums["values"][0])
        results.append(("sums_match", "OK" if match else "MISMATCH"))
    step("save_ZCode_MCP_Test", lambda: server.excel_save_as(book_name, r"C:\path\to\work\test.xlsx"))
    # лист по умолчанию убирать не нужно; книгу оставляем ОТКРЫТОЙ для просмотра

    # --- Word ---
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    doc = word.Documents.Add()
    doc_name = doc.Name
    step("word_insert", lambda: server.word_insert_text(doc_name, "Alpha beta gamma. Alpha again."))
    rd = step("word_read", lambda: server.word_read_document(doc_name))
    if rd:
        print("word text:", repr(rd["text"][:80]))
    step("word_replace", lambda: server.word_replace_text(doc_name, "Alpha", "Omega", match_case=True))
    rd2 = server.word_read_document(doc_name)
    if rd2:
        ok = "Omega beta gamma" in rd2["text"] and "Alpha" not in rd2["text"]
        results.append(("word_replace_result", "OK" if ok else "MISMATCH"))
        print("REPLACE RESULT:", "OK" if ok else "MISMATCH", "|", repr(rd2["text"][:60]))
    step("word_save_as", lambda: server.word_save_as(doc_name, r"C:\path\to\office-live-mcp\scratch-test.docx"))
    word = reattach("Word.Application")
    word.Documents.Item(doc_name).Close(SaveChanges=0)
    word.Quit()

    print("\n=== SUMMARY ===")
    for n, s in results:
        print(f"{s:14} {n}")
    print("ZCode_MCP_Test.xlsx exists:", os.path.exists(r"C:\path\to\work\test.xlsx"))


if __name__ == "__main__":
    main()
