"""Range PNG file output, using fake Office objects only."""

import base64
import io
import json
import random
from types import SimpleNamespace

import pytest
from PIL import Image as PILImage

from office_live import config, excel_format, registry
from office_live.errors import ToolError


def png_bytes(size=(32, 24), noisy=False):
    if noisy:
        pixels = random.Random(15).randbytes(size[0] * size[1] * 3)
        image = PILImage.frombytes("RGB", size, pixels)
    else:
        image = PILImage.new("RGB", size, "white")
    with io.BytesIO() as buf:
        image.save(buf, format="PNG")
        return buf.getvalue()


@pytest.fixture
def render(monkeypatch, tmp_path):
    data = png_bytes()
    app = SimpleNamespace()
    wb = SimpleNamespace(Name="Book.xlsx")
    ws = SimpleNamespace(Name="Data")
    rng = SimpleNamespace(Address="A1:B2")
    calls = []
    monkeypatch.setattr(config, "SETTINGS", config.load({"OFFICE_LIVE_ALLOWED_DIRS": str(tmp_path)}))
    monkeypatch.setattr(excel_format, "pick_workbook", lambda name: (app, wb))
    monkeypatch.setattr(excel_format, "get_range", lambda book, sheet, cells: (ws, rng))
    monkeypatch.setattr(excel_format, "addr_of", lambda selected: selected.Address)

    def fake_png(*args):
        calls.append(args)
        return data

    monkeypatch.setattr(excel_format, "render_range_png", fake_png)
    return SimpleNamespace(data=data, calls=calls, args=(app, wb, ws, rng))


@pytest.mark.parametrize("kwargs", [{}, {"save_to": ""}])
def test_without_save_to_keeps_inline_response(render, kwargs):
    result = excel_format.excel_render_range_image("Book.xlsx", "Data", "A1:B2", **kwargs)
    assert len(result) == 2
    assert json.loads(result[0]) == {
        "ok": True, "workbook": "Book.xlsx", "sheet": "Data", "cells": "A1:B2", "png_bytes": len(render.data),
    }
    assert isinstance(result[1], excel_format.Image)
    content = result[1].to_image_content()
    assert content.mime_type == "image/png"
    assert base64.b64decode(content.data) == render.data


@pytest.mark.parametrize("large", [False, True])
def test_save_to_writes_png_and_returns_only_file_metadata(render, monkeypatch, tmp_path, large):
    data = png_bytes((400, 300), noisy=True) if large else render.data
    if large:
        assert len(data) > 319 * 1024  # Exceeds the externally measured client limit.
        monkeypatch.setattr(excel_format, "render_range_png", lambda *args: data)
    monkeypatch.chdir(tmp_path)
    result = excel_format.excel_render_range_image("Book.xlsx", "Data", "A1:B2", max_cells=100, save_to="renders/range.png")
    dest = tmp_path / "renders" / "range.png"
    assert dest.read_bytes() == data
    with PILImage.open(dest) as image:
        assert image.format == "PNG"
        size = list(image.size)
    assert len(result) == 1 and isinstance(result[0], str)
    assert json.loads(result[0]) == {
        "ok": True, "workbook": "Book.xlsx", "sheet": "Data", "cells": "A1:B2",
        "path": str(dest), "png_bytes": len(data), "size_px": size,
    }
    if not large:
        assert render.calls == [(*render.args, 100)]


@pytest.mark.parametrize("traversal", [False, True])
def test_save_to_outside_allowed_folders_is_refused_before_render(render, tmp_path, traversal):
    allowed = tmp_path / "allowed"
    config_env = {"OFFICE_LIVE_ALLOWED_DIRS": str(allowed)}
    # Keep the attempted (refused) destination inside the executor worktree.
    dest = allowed / ".." / "outside.png" if traversal else tmp_path / "outside.png"
    with pytest.MonkeyPatch.context() as patch:
        patch.setattr(config, "SETTINGS", config.load(config_env))
        with pytest.raises(ToolError, match="outside the allowed directories"):
            excel_format.excel_render_range_image(save_to=str(dest))
    assert not (tmp_path / "outside.png").exists() and not render.calls


def test_save_to_refuses_existing_file_unless_overwrite_is_requested(render, tmp_path):
    dest = tmp_path / "range.png"
    dest.write_bytes(b"keep this file")
    with pytest.raises(ToolError, match="overwrite=true"):
        excel_format.excel_render_range_image(save_to=str(dest))
    assert dest.read_bytes() == b"keep this file" and not render.calls
    result = excel_format.excel_render_range_image(save_to=str(dest), overwrite=True)
    assert dest.read_bytes() == render.data
    assert json.loads(result[0])["path"] == str(dest)


def test_save_to_refuses_file_created_during_render(render, monkeypatch, tmp_path):
    dest = tmp_path / "range.png"

    def concurrent_file(*args):
        dest.write_bytes(b"created during render")
        return render.data

    monkeypatch.setattr(excel_format, "render_range_png", concurrent_file)
    with pytest.raises(ToolError, match="overwrite=true"):
        excel_format.excel_render_range_image(save_to=str(dest))
    assert dest.read_bytes() == b"created during render"


def test_save_to_requires_png_extension(render, tmp_path):
    dest = tmp_path / "range.jpg"
    with pytest.raises(ToolError, match="Unsupported file extension"):
        excel_format.excel_render_range_image(save_to=str(dest))
    assert not dest.exists() and not render.calls


def test_save_to_is_refused_in_readonly_but_inline_still_works(render, monkeypatch, tmp_path):
    monkeypatch.setattr(config, "SETTINGS", config.load({"OFFICE_LIVE_MODE": "readonly"}))
    dest = tmp_path / "range.png"
    with pytest.raises(ToolError, match="read-only mode"):
        excel_format.excel_render_range_image(save_to=str(dest))
    assert not dest.exists() and not render.calls
    assert len(excel_format.excel_render_range_image()) == 2


@pytest.mark.parametrize("save", [False, True])
@pytest.mark.parametrize("blank", [False, True])
def test_hidden_or_minimized_render_errors_stay_unchanged(monkeypatch, tmp_path, save, blank):
    app = SimpleNamespace(WindowState=-4143, Goto=lambda *args: None)
    wb = SimpleNamespace(Name="Book.xlsx", AutoSaveOn=True, Activate=lambda: None,
                         Windows=lambda i: SimpleNamespace(WindowState=-4140))
    ws = SimpleNamespace(Name="Data", Activate=lambda: None)
    rng = SimpleNamespace(Address="A1:B2", Width=100, Height=50, Cells=lambda *args: None)
    monkeypatch.setattr(config, "SETTINGS", config.load({"OFFICE_LIVE_ALLOWED_DIRS": str(tmp_path)}))
    monkeypatch.setattr(excel_format, "pick_workbook", lambda name: (app, wb))
    monkeypatch.setattr(excel_format, "get_range", lambda *args: (ws, rng))
    monkeypatch.setattr(excel_format, "bounds", lambda rng: (1, 1, 2, 2))
    monkeypatch.setattr(excel_format, "count_nonempty", lambda *args: 1)
    monkeypatch.setattr(excel_format, "suspend_events", lambda app: None)
    monkeypatch.setattr(excel_format, "_remember_view", lambda app: {})
    monkeypatch.setattr(excel_format, "_restore_view", lambda *args: None)
    monkeypatch.setattr(excel_format.time, "sleep", lambda seconds: None)
    monkeypatch.setattr(excel_format, "_copy_picture_png", lambda *args: png_bytes() if blank else b"")
    expected = (
        "The rendered image is blank although the range has data: the workbook window is probably hidden or minimized. Make it visible and retry."
        if blank else "Could not render the range (the workbook window must be visible and not minimized)."
    )
    dest = tmp_path / "renders" / "range.png"
    kwargs = {"save_to": str(dest)} if save else {}
    with pytest.raises(ToolError) as exc:
        excel_format.excel_render_range_image(**kwargs)
    assert str(exc.value) == expected
    assert not dest.parent.exists() and wb.AutoSaveOn is True


def test_save_to_is_declared_as_a_file_writing_argument():
    assert registry.CATALOG["excel_render_range_image"].file_args == ("save_to",)
