"""Own temporary .xls compatibility regression. Run only with the owner's permission."""

from pathlib import Path

import pytest

from tests.live.mcpclient import Stdio
from tests.live.test_live_stage1 import in_excel

pytestmark = pytest.mark.live


@pytest.fixture
def srv(tmp):
    client = Stdio(env={"OFFICE_LIVE_MODE": "full", "OFFICE_LIVE_ALLOWED_DIRS": tmp,
                        "OFFICE_LIVE_JOURNAL_DIR": str(Path(tmp) / "journal"),
                        "OFFICE_LIVE_AUDIT_LOG": str(Path(tmp) / "audit.jsonl")}, timeout=30)
    yield client
    client.close()


def test_xls_save_with_modern_formula_restores_compatibility(srv, wb, tmp):
    path = str(Path(tmp) / "compatibility.xls")
    # First save the empty book, then add an unsupported formula: no setup-time compatibility dialog.
    srv.call("excel_save_as", workbook=wb, path=path)

    def prepare(app, ws):
        ws.Parent.CheckCompatibility = True
        ws.Range("B1").Formula2 = "=LET(probe_value,2,probe_value+1)"
        ws.Range("B1").Calculate()
        assert ws.Range("B1").Value2 == 3, "This probe requires Excel with LET support"
        return bool(app.DisplayAlerts)

    alerts = in_excel(path, prepare)
    srv.call("excel_write_range", workbook=path, sheet="Data", cells="A1", values=[[42]])
    assert in_excel(path, lambda app, ws: bool(ws.Parent.CheckCompatibility)) is True
    assert in_excel(path, lambda app, ws: bool(ws.Parent.Saved)) is False
    result = srv.call("excel_save", workbook=path)
    assert result["compatibility_check_skipped"] is True and result["saved"] is True
    assert "excel_save_as" in result["note"]
    assert in_excel(path, lambda app, ws: (bool(ws.Parent.CheckCompatibility), bool(ws.Parent.Saved), bool(app.DisplayAlerts))) == (True, True, alerts)
    books = srv.call("excel_list_workbooks")["workbooks"]  # same server remains responsive after Save
    assert any(book["name"] == "compatibility.xls" and book["saved"] for book in books)
    # Documented side effect: the file itself keeps the checker disabled; only the open book got its flag back.
    srv.call("excel_close_workbook", workbook=path)
    srv.call("excel_open_workbook", path=path)
    assert in_excel(path, lambda app, ws: bool(ws.Parent.CheckCompatibility)) is False
