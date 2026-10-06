"""Opt-in ONLY. Own fixture windows; chat detection/movement is always disabled.

Written for a separately authorized live run. Never invokes ShellExecute or changes
the registry. Fixture cleanup closes only our documents; restore runs in finally.
"""

import pytest

pytestmark = pytest.mark.live


@pytest.fixture(autouse=True)
def _visible_test_excel(wb):
    """A bloated test Excel restarted mid-run (Quit while the test server still held references) stays alive but
    hidden, and window tests then see "a target window is hidden". Make the instance that hosts OUR workbook visible."""
    from office_live import com, xl_common

    def show():
        app, _ = xl_common.pick_workbook(wb)
        if not bool(app.Visible):
            app.Visible = True

    com.run_com(show)


@pytest.mark.parametrize("kind", ["excel", "word"])
def test_own_window_status_focus_arrange_restore(srv, wb, doc, kind):
    target = {"workbook": wb} if kind == "excel" else {"document": doc}
    before = srv.call("office_window", include_chat=False, **target)
    original = before["office_windows"][0]
    assert original["responding"] and original["monitor"]["work_area"]
    assert before["chat_window"] is None
    try:
        arranged = srv.call("office_window", action="arrange", layout="office_maximized", monitor="office", include_chat=False, **target)
        assert arranged["ok"] and arranged["chat_window"] is None
        assert arranged["applied"] == [original["hwnd"]]
        focused = srv.call("office_window", action="focus", include_chat=False, **target)
        assert isinstance(focused["foreground"], bool)
        if not focused["foreground"]:
            assert focused["reason"]  # Windows may legitimately deny stealing focus.
        status = srv.call("office_window", include_chat=False, **target)
        assert status["office_windows"][0]["state"] == "maximized"
    finally:
        restored = srv.call("office_window", action="restore", include_chat=False)
        assert restored["ok"], restored
    after = srv.call("office_window", include_chat=False, **target)["office_windows"][0]
    assert after["state"] == original["state"]
    assert after["rect"] == original["rect"]


def test_own_excel_and_word_side_by_side(srv, wb, doc):
    originals = [srv.call("office_window", include_chat=False, **target)["office_windows"][0]
                 for target in ({"workbook": wb}, {"document": doc})]
    try:
        result = srv.call("office_window", action="arrange", workbook=wb, document=doc,
                          layout="excel_word_side_by_side", monitor="office", office_share=0.5, include_chat=False)
        assert result["ok"] and result["chat_window"] is None
        assert set(result["applied"]) == {o["hwnd"] for o in originals}
        assert len(result["office_windows"]) == 2
        for actual, desired in zip(result["office_windows"], result["requested"], strict=True):
            assert actual["frame_rect"] == desired["frame_rect"], result["warnings"]
    finally:
        restored = srv.call("office_window", action="restore", include_chat=False)
        assert restored["ok"], restored
    for target, original in zip(({"workbook": wb}, {"document": doc}), originals, strict=True):
        after = srv.call("office_window", include_chat=False, **target)["office_windows"][0]
        assert (after["state"], after["rect"]) == (original["state"], original["rect"])
