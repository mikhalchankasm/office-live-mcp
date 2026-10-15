"""Written probes for marked temporary workbooks; run only with the owner's separate approval."""

import pytest

from tests.live.conftest import read
from tests.live.test_live_office_operations import srv as isolated_server, operations_book as isolated_book
from tests.live.test_live_stage1 import in_excel

pytestmark = pytest.mark.live
srv, operations_book = isolated_server, isolated_book


@pytest.mark.parametrize("direction", ["both", "row", "column"])
def test_probe_data_table_loan_values_array_undo(srv, operations_book, direction):
    book = operations_book
    def setup(app, ws):
        ws.Range("A1:A3").Value2 = ((0.05,), (12,), (1000,))
        ws.Range("B9").Formula = "=-PMT(A1/12,A2,A3)"
        ws.Range("C9:F9").Value2 = ((12, 24, 36, 48),)
        ws.Range("B10:B15").Value2 = tuple((rate,) for rate in (0.02, 0.04, 0.06, 0.08, 0.10, 0.12))
        if direction == "row":
            ws.Range("B10:B15").Formula = "=-PMT(A1/12,A2,A3)"
        if direction == "column":
            ws.Range("C9:F9").Formula = "=-PMT(A1/12,A2,A3)"
    in_excel(book, setup)
    result = srv.call("excel_what_if", workbook=book, sheet="Data", action="data_table", cells="B9:F15",
                      row_input="A2" if direction != "column" else "", column_input="A1" if direction != "row" else "")
    assert result["dependency_verified"] and result["results"] == "C10:F15"
    srv.call("excel_calculate", workbook=book)
    actual = read(srv, book, "C10:F15")
    for i, rate in enumerate((0.02, 0.04, 0.06, 0.08, 0.10, 0.12)):
        for j, term in enumerate((12, 24, 36, 48)):
            monthly = (0.05 if direction == "row" else rate) / 12
            months = 12 if direction == "column" else term
            assert actual[i][j] == pytest.approx(1000 * monthly / (1 - (1 + monthly) ** -months))
    srv.call("office_undo", file=book)
    assert in_excel(book, lambda app, ws: ws.Range("C10:F15").HasArray) is False
    assert all(v is None for row in read(srv, book, "C10:F15") for v in row)


def test_probe_scenario_show_delete_summary_undo(srv, operations_book):
    book = operations_book
    def setup(app, ws):
        ws.Range("A1:A2").Value2 = ((2,), (3,))
        ws.Range("A1").NumberFormat = "0%"
        ws.Range("B1").Formula = "=A1*A2"
        return ws.Range("A1:A2").Formula, ws.Range("A1").NumberFormat
    before = in_excel(book, setup)
    def action(act, **kw):
        return srv.call("excel_what_if", workbook=book, sheet="Data", action=act, **kw)
    action("add_scenario", name="Probe", changing_cells="A1,A2", values=[5, 8], comment="own temporary scenario")
    saved = action("list_scenarios")["scenarios"]
    action("show_scenario", name="Probe")
    assert read(srv, book, "A1:A2") == [[5], [8]]
    srv.call("office_undo", file=book)
    assert in_excel(book, lambda app, ws: (ws.Range("A1:A2").Formula, ws.Range("A1").NumberFormat)) == before
    action("delete_scenario", name="Probe")
    srv.call("office_undo", file=book)
    assert action("list_scenarios")["scenarios"] == saved
    sheets = in_excel(book, lambda app, ws: ws.Parent.Sheets.Count)
    assert len(action("scenario_summary", result_cells="B1")["created_sheets"]) == 1
    srv.call("office_undo", file=book)
    assert in_excel(book, lambda app, ws: ws.Parent.Sheets.Count) == sheets
    assert in_excel(book, lambda app, ws: ws.Range("A1:A2").Formula) == before[0]
    srv.call("office_undo", file=book)
    assert action("list_scenarios")["scenarios"] == []
