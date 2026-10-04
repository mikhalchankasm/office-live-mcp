"""Согласованность упаковки: зависимости pyproject и requirements.txt, версия, установщик."""

import re
from pathlib import Path

import pytest

from office_live import __version__

ROOT = Path(__file__).resolve().parent.parent


def _normalise(spec: str) -> str:
    """'pywin32>=306; sys_platform == "win32"' -> 'pywin32>=306' (маркеры среды не сравниваем)."""
    return re.sub(r"\s+", "", spec.split(";")[0]).lower()


def _pyproject():
    tomllib = pytest.importorskip("tomllib")  # Python 3.11+
    return tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))


def test_requirements_txt_matches_pyproject_dependencies():
    declared = {_normalise(d) for d in _pyproject()["project"]["dependencies"]}
    listed = {_normalise(line) for line in (ROOT / "requirements.txt").read_text(encoding="utf-8").splitlines() if line.strip() and not line.startswith("#")}
    assert declared == listed, "keep requirements.txt and pyproject.toml dependencies identical"


def test_every_runtime_dependency_has_a_lower_bound_and_mcp_an_upper_one():
    deps = _pyproject()["project"]["dependencies"]
    assert all(">=" in d for d in deps)
    assert any(d.startswith("mcp") and "<3" in d for d in deps), "the MCP SDK API changes between major versions: pin its upper bound"


def test_package_version_matches_pyproject():
    assert _pyproject()["project"]["version"] == __version__


def test_lint_configuration_is_not_trivially_empty():
    ruff = _pyproject()["tool"]["ruff"]["lint"]["select"]
    assert {"F", "B"} <= set(ruff)


def test_installer_stays_ascii_with_windows_line_endings():
    """install.cmd запускает cmd.exe в чужой кодовой странице: кириллица и LF в .cmd ломают метки и скобки."""
    raw = (ROOT / "install.cmd").read_bytes()
    raw.decode("ascii")
    assert b"\r\n" in raw and b"\n" not in raw.replace(b"\r\n", b"")


def test_packaged_installer_stays_ascii_with_windows_line_endings():
    raw = (ROOT / "packaging" / "install.cmd").read_bytes()
    raw.decode("ascii")
    assert b"\r\n" in raw and b"\n" not in raw.replace(b"\r\n", b"")
