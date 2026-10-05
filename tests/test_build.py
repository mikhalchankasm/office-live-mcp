"""Compiler discovery and artifact wiring, without running any built executable."""

import importlib.util
from pathlib import Path
import zipfile

import pytest

SPEC = importlib.util.spec_from_file_location("office_live_build", Path(__file__).resolve().parents[1] / "packaging/build.py")
build = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(build)


def test_compiler_search_order_and_missing_hint(tmp_path, monkeypatch):
    locations = [tmp_path / "override/ISCC.exe", tmp_path / "programs/Inno Setup 6/ISCC.exe",
                 tmp_path / "local/Programs/Inno Setup 6/ISCC.exe", tmp_path / "path/ISCC.exe"]
    monkeypatch.setenv("ISCC", str(locations[0]))
    monkeypatch.setenv("ProgramFiles(x86)", str(tmp_path / "programs"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))
    monkeypatch.setattr(build.shutil, "which", lambda name: str(locations[3]))
    for path in locations:
        path.parent.mkdir(parents=True)
        path.write_bytes(b"fake compiler")
    for path in locations:
        assert build.find_iscc() == path
        path.unlink()
    with pytest.raises(FileNotFoundError, match="--no-setup"):
        build.find_iscc()


@pytest.mark.parametrize("no_setup", [True, False])
def test_build_artifacts_and_compiler_arguments(tmp_path, monkeypatch, no_setup):
    dist, work = tmp_path / "dist", tmp_path / "work"
    compiler = tmp_path / "Inno Setup 6/ISCC.exe"
    monkeypatch.setattr(build, "find_iscc", lambda: compiler if not no_setup else pytest.fail("--no-setup must not need Inno"))
    monkeypatch.setattr(build, "smoke", lambda exe: 114)
    calls = []

    def run(command, **kwargs):
        calls.append(command)
        if "PyInstaller" in command:
            app = work / "dist/office-live-mcp"
            app.mkdir(parents=True)
            (app / "office-live-mcp.exe").write_bytes(b"exe")
            (app / "office-live-link.exe").write_bytes(b"link")
        else:
            assert (work / "dist/office-live-mcp/office-live-mcp.exe").exists()
            (dist / f"office-live-mcp-{build.__version__}-setup.exe").write_bytes(b"setup")

    monkeypatch.setattr(build.subprocess, "run", run)
    assert build.build(dist, work, no_setup=no_setup) == 0
    assert len(calls) == (1 if no_setup else 2)
    archive = next(dist.glob("*.zip"))
    with zipfile.ZipFile(archive) as z:
        assert any(n.endswith("/app/office-live-mcp.exe") for n in z.namelist())
        assert any(n.endswith("/LICENSE.txt") for n in z.namelist())
    if not no_setup:
        assert calls[-1][:4] == [str(compiler), f"/DAppVersion={build.__version__}", f"/DSourceDir={work / 'dist/office-live-mcp'}", f"/O{dist}"]
