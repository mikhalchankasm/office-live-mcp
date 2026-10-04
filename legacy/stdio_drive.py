# Драйвер stdio-теста office-live: настоящий JSON-RPC к server.py в отдельном процессе.
# Использование: python stdio_drive.py <excel_scratch> <excel_deliv> <word_doc>
import json
import subprocess
import sys
import threading

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

EXCEL_SCRATCH, EXCEL_DELIV, WORD_DOC = sys.argv[1], sys.argv[2], sys.argv[3]

proc = subprocess.Popen(
    [sys.executable, "server.py"],
    stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
    text=True, encoding="utf-8",
)


def drain_stderr():
    for line in proc.stderr:
        if line.strip():
            print("[server-stderr]", line.rstrip())


threading.Thread(target=drain_stderr, daemon=True).start()

_next_id = [0]


def rpc(method, params=None, timeout=60):
    _next_id[0] += 1
    rid = _next_id[0]
    msg = {"jsonrpc": "2.0", "id": rid, "method": method}
    if params is not None:
        msg["params"] = params
    proc.stdin.write(json.dumps(msg) + "\n")
    proc.stdin.flush()
    # ждём ответ с нужным id
    while True:
        line = proc.stdout.readline()
        if not line:
            raise RuntimeError("server closed stdout")
        data = json.loads(line)
        if data.get("id") == rid:
            return data


def call_tool(name, args, label=None, expect_error=False):
    resp = rpc("tools/call", {"name": name, "arguments": args})
    result = resp.get("result", {})
    is_err = result.get("isError", False)
    text = result.get("content", [{}])[0].get("text", "")
    shown = label or name
    if expect_error:
        status = "OK(expected-error)" if is_err else "FAIL(no error raised)"
        print(f"[{status}] {shown}: {text[:160]}")
        return None
    if is_err:
        print(f"[FAIL] {shown}: {text[:200]}")
        return None
    print(f"[OK] {shown}")
    return json.loads(text)


results = []


def check(name, cond, extra=""):
    results.append((name, bool(cond)))
    print(f"   -> {'PASS' if cond else 'FAIL'} {name} {extra}")


# handshake
rpc("initialize", {"protocolVersion": "2024-11-05", "capabilities": {}, "clientInfo": {"name": "drive", "version": "1"}})
proc.stdin.write(json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized"}) + "\n")
proc.stdin.flush()

# --- Excel: базовые инструменты на черновике ---
call_tool("excel_list_workbooks", {})
call_tool("excel_add_worksheet", {"workbook": EXCEL_SCRATCH, "name": "Live_Test"})
call_tool("excel_write_range", {"workbook": EXCEL_SCRATCH, "sheet": "Live_Test", "cells": "B2:C3", "values": [[11, 22], [33, 44]]})
r = call_tool("excel_set_formula", {"workbook": EXCEL_SCRATCH, "sheet": "Live_Test", "cell": "B5", "formula": "=SUM(B2:C3)"})
check("formula value = 110", r and r.get("value") == 110)
call_tool("excel_format_range", {"workbook": EXCEL_SCRATCH, "sheet": "Live_Test", "cells": "B5", "bold": True, "fill_color": "#FFF2CC"})
call_tool("excel_autofit", {"workbook": EXCEL_SCRATCH, "sheet": "Live_Test", "target": "columns"})
call_tool("excel_save_as", {"workbook": EXCEL_SCRATCH, "path": r"C:\path\to\office-live-mcp\scratch-test.xlsx"})
r = call_tool("excel_read_range", {"workbook": "scratch-test.xlsx", "sheet": "Live_Test", "cells": "B2:C5"})
check("readback grid", r and r["values"] == [[11, 22], [33, 44], [None, None], [110, None]], str(r and r["values"]))

# безопасность: перезапись существующего без overwrite=True должна отказать
call_tool("excel_save_as", {"workbook": "scratch-test.xlsx", "path": r"C:\path\to\office-live-mcp\scratch-test.xlsx"}, label="excel_save_as refuse-overwrite", expect_error=True)

# --- Полный круг: MCP_Test -> ZCode_MCP_Test.xlsx ---
call_tool("excel_add_worksheet", {"workbook": EXCEL_DELIV, "name": "MCP_Test"})
r = call_tool("excel_write_range", {"workbook": EXCEL_DELIV, "sheet": "MCP_Test", "cells": "A1:J10", "values": MATRIX})
check("100 cells written", r and r["readback"] == MATRIX)
for col in "ABCDEFGHIJ":
    call_tool("excel_set_formula", {"workbook": EXCEL_DELIV, "sheet": "MCP_Test", "cell": f"{col}12", "formula": f"=SUM({col}1:{col}10)"}, label=f"excel_set_formula {col}12")
call_tool("excel_format_range", {"workbook": EXCEL_DELIV, "sheet": "MCP_Test", "cells": "A12:J12", "bold": True})
call_tool("excel_autofit", {"workbook": EXCEL_DELIV, "sheet": "MCP_Test", "target": "columns"})
r = call_tool("excel_read_range", {"workbook": EXCEL_DELIV, "sheet": "MCP_Test", "cells": "A12:J12"})
check("column sums", r and r["values"][0] == EXPECT_SUMS, str(r and r["values"][0]))
call_tool("excel_save_as", {"workbook": EXCEL_DELIV, "path": r"C:\path\to\work\test.xlsx"})

# --- Word ---
call_tool("word_list_documents", {})
call_tool("word_insert_text", {"document": WORD_DOC, "text": "Alpha beta gamma. Alpha again."})
r = call_tool("word_read_document", {"document": WORD_DOC})
check("word text", r and "Alpha beta gamma" in r["text"])
r = call_tool("word_replace_text", {"document": WORD_DOC, "search": "Alpha", "replace": "Omega", "match_case": True})
check("replace estimate 2", r and r.get("replaced_estimate") == 2, str(r))
r = call_tool("word_read_document", {"document": WORD_DOC})
check("word after replace", r and "Omega beta gamma" in r["text"] and "Alpha" not in r["text"])
call_tool("word_save_as", {"document": WORD_DOC, "path": r"C:\path\to\office-live-mcp\scratch-test.docx"})
call_tool("word_save_as", {"document": "scratch-test.docx", "path": r"C:\path\to\office-live-mcp\scratch-test.docx"}, label="word_save_as refuse-overwrite", expect_error=True)

proc.stdin.close()
proc.wait(timeout=30)

print("\n=== SUMMARY ===")
passed = sum(1 for _, ok in results if ok)
for n, ok in results:
    print(f"{'PASS' if ok else 'FAIL'}  {n}")
print(f"{passed}/{len(results)} checks passed")
sys.exit(0 if passed == len(results) else 1)
