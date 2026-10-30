"""excel_manage_charts(export_image) never reports success for an empty PNG (issue #32)."""

import os
from types import SimpleNamespace

import pytest

from office_live import excel_charts as analysis
from office_live.errors import ToolError

PNG = b"\x89PNG\r\n\x1a\n" + b"\0" * 600


class Chart:
    def __init__(self, outputs):
        self.outputs, self.calls = list(outputs), []

    def Export(self, path, fmt):
        self.calls.append("Export")
        with open(path, "wb") as f:
            f.write(self.outputs.pop(0))

    def Refresh(self):
        self.calls.append("Refresh")


@pytest.fixture
def chart(tmp_path, monkeypatch):
    def make(outputs):
        ch = Chart(outputs)
        co = SimpleNamespace(Name="Chart1", Chart=ch)
        wb = SimpleNamespace(Name="Book.xlsx")
        monkeypatch.setattr(analysis, "pick_workbook", lambda *a, **kw: (None, wb))
        monkeypatch.setattr(analysis, "_find_chart", lambda *a, **kw: (None, co))
        monkeypatch.setattr(analysis, "check_path", lambda p, *a, **kw: str(p))
        monkeypatch.setattr(analysis.time, "sleep", lambda s: None)
        return ch

    return make


def call(**kw):
    fn = getattr(analysis.excel_manage_charts, "__wrapped__", analysis.excel_manage_charts)
    return fn(workbook="Book.xlsx", action="export_image", chart_name="Chart1", **kw)


def test_empty_first_export_is_retried_after_refresh(chart, tmp_path):
    ch = chart([b"", PNG])
    dest = tmp_path / "chart.png"
    result = call(export_path=str(dest))
    assert ch.calls == ["Export", "Refresh", "Export"]
    assert dest.read_bytes() == PNG and f'"png_bytes": {len(PNG)}' in result[0]


def test_twice_empty_export_is_an_error_and_saves_nothing(chart, tmp_path):
    chart([b"", b""])
    dest = tmp_path / "chart.png"
    with pytest.raises(ToolError, match="empty image"):
        call(export_path=str(dest))
    assert not os.path.exists(dest)


def test_good_export_needs_no_retry(chart):
    ch = chart([PNG])
    call()
    assert ch.calls == ["Export"]
