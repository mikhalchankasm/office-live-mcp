"""Real Inno lifecycle ONLY on the disposable Windows CI runner.

Environment variables isolate our CLI configs, NOT HKCU or Windows known folders.
The two opt-ins below must stay in place; never run this test on a workstation.
"""

import json
import os
import subprocess
import tempfile
import time
from pathlib import Path

import pytest

from office_live import __version__
from office_live.probe import smoke

pytestmark = pytest.mark.skipif(
    os.environ.get("CI", "").lower() != "true" or os.environ.get("OFFICE_LIVE_SETUP_EXE_TEST") != "1",
    reason="setup.exe writes real HKCU: disposable CI only, with OFFICE_LIVE_SETUP_EXE_TEST=1",
)
APP_ID = "{47C3867B-C095-4C76-AACD-DA6A268CF854}"
UNINSTALL_KEY = rf"Software\Microsoft\Windows\CurrentVersion\Uninstall\{APP_ID}_is1"
REPO = Path(__file__).resolve().parents[1]


def uninstall_location():
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
            return winreg.QueryValueEx(key, "InstallLocation")[0]
    except FileNotFoundError:
        return None


def test_silent_install_upgrade_uninstall():
    assert uninstall_location() is None, "CI runner already has this product; refuse to touch it"
    setup = REPO / "dist" / f"office-live-mcp-{__version__}-setup.exe"
    assert setup.is_file()
    with tempfile.TemporaryDirectory(prefix="olsetup-") as temporary:
        test_root = Path(temporary)
        root = test_root / "Программа с пробелами"
        env = {k: v for k, v in os.environ.items() if not k.startswith("OFFICE_LIVE_")}
        for key, part in (("USERPROFILE", "home"), ("HOME", "home"), ("APPDATA", "roaming"),
                          ("LOCALAPPDATA", "local"), ("CODEX_HOME", "Особый Codex")):
            env[key] = str(test_root / part)
        env["OFFICE_LIVE_LIVE_TESTS"] = "0"
        exe, unins = root / "app/office-live-mcp.exe", root / "unins000.exe"
        state = test_root / "local/office-live-mcp/state.json"

        def run(command, *args):
            result = subprocess.run([str(command), *args], env=env, cwd=test_root, capture_output=True, timeout=240)
            assert result.returncode == 0, (result.stdout + result.stderr).decode("utf-8", "replace")
            return result.stdout

        base = ["/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", f"/DIR={root}", "/NOLINKS"]
        try:
            run(setup, *base, "/CLIENTS=none", f"/LOG={test_root / 'install.log'}")
            assert exe.is_file() and unins.is_file() and (root / "app/_internal").is_dir()
            assert (root / "LICENSE.txt").is_file() and (root / "app/office-live-mcp.files.json").is_file()
            assert json.loads((root / "office-live-mcp.install").read_text(encoding="utf-8"))["product"] == "office-live-mcp"
            assert Path(uninstall_location()) == root
            assert json.loads(state.read_text(encoding="utf-8"))["registrations"] == []
            assert smoke([str(exe)], env) > 0
            # An omitted /CLIENTS on upgrade must keep both the entry and its access mode.
            run(exe, "setup", "--yes", "--clients", "codex", "--readonly")
            config = Path(env["CODEX_HOME"]) / "config.toml"
            before = config.read_bytes()
            foreign = root / "app/user-notes.txt"
            foreign.write_text("keep this unrelated file", encoding="utf-8")
            stale = root / "app/_internal/stale-library.dll"
            stale.write_bytes(b"obsolete")
            # A CLI failure must propagate through Inno's custom exit code, with files retained.
            failed = subprocess.run([str(setup), *base, "/ACCESS=invalid", f"/LOG={test_root / 'failed-upgrade.log'}"],
                                    env=env, cwd=test_root, capture_output=True, timeout=240)
            assert failed.returncode != 0 and exe.is_file() and config.read_bytes() == before
            run(setup, *base, f"/LOG={test_root / 'upgrade.log'}")
            assert not stale.exists() and config.read_bytes() == before
            assert smoke([str(exe)], env) > 0
            assert foreign.read_text(encoding="utf-8") == "keep this unrelated file"
            run(unins, "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", f"/LOG={test_root / 'uninstall.log'}")
            deadline = time.monotonic() + 30
            while (exe.exists() or unins.exists() or uninstall_location()) and time.monotonic() < deadline:
                time.sleep(0.2)
            assert not exe.exists() and not unins.exists() and uninstall_location() is None
            assert not (root / "app/_internal").exists() and not (root / "office-live-mcp.install").exists()
            assert not (root / "app/office-live-mcp.files.json").exists()
            assert foreign.read_text(encoding="utf-8") == "keep this unrelated file"
            assert 'mcp_servers.office-live' not in config.read_text(encoding="utf-8")
            data = json.loads(state.read_text(encoding="utf-8"))
            assert data["root"] == "" and data["registrations"] == []
            assert list(config.parent.glob("*.bak-*"))
        finally:
            # Only our temporary installation, never a pre-existing product.
            if unins.exists() and uninstall_location() and Path(uninstall_location()) == root:
                run(unins, "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART")
