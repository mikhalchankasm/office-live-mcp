import copy
import json
import sys
from pathlib import Path

import pytest

from office_live import cli, config, install, links, protocol, registrations, state

NATIVE_REGISTRY = protocol.Registry


@pytest.fixture
def link_install(tmp_path, monkeypatch):
    root = tmp_path / "own install"
    (root / "app").mkdir(parents=True)
    exe = root / "app" / "office-live-mcp.exe"
    exe.write_text("fixture")
    exe.with_name(protocol.GUI_EXE).write_text("fixture GUI")
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(exe))
    monkeypatch.setattr(cli, "_detect_clients", lambda: dict.fromkeys(cli.CLIENT_LABELS, False))
    return root, exe


def test_native_registry_writes_are_blocked_even_if_wrapper_is_bypassed():
    import winreg

    for method, args in (("CreateKeyEx", (winreg.HKEY_CURRENT_USER, protocol.KEY)),
                         ("SetValueEx", (None, "", 0, 1, "bad")), ("DeleteKey", (winreg.HKEY_CURRENT_USER, protocol.KEY))):
        with pytest.raises(AssertionError, match="never write"):
            getattr(winreg, method)(*args)
    with pytest.raises(AssertionError, match="never write"):
        NATIVE_REGISTRY().write("bad")


def test_setup_links_and_state_roundtrip_without_client_side_effects(link_install, fake_registry, capsys):
    root, exe = link_install
    assert cli.setup_cmd(["--links", "--allowed-dirs", str(root)]) == 0
    command = f'"{exe.with_name(protocol.GUI_EXE)}" open-link "%1"'
    assert fake_registry.snapshot == protocol.tree(command)
    record, = state.read_state()["registrations"]
    assert record["kind"] == "protocol" and record["status"] == "confirmed" and record["root"] == str(root)
    assert record["allowed_dirs"] == [str(root)]
    assert not Path(registrations.target("codex")["path"]).exists()
    assert cli.setup_cmd(["--no-links"]) == 0
    assert fake_registry.snapshot is None and not state.read_state()["registrations"]
    assert "Удалить свою схему" in capsys.readouterr().out


def test_dry_run_does_not_write_registry_profile_or_state(link_install, fake_registry, capsys):
    before = sorted(Path.home().parent.rglob("*")) if Path.home().parent.exists() else []
    assert cli.setup_cmd(["--links", "--dry-run"]) == 0
    assert "write" not in fake_registry.calls and not state.state_path().exists()
    after = sorted(Path.home().parent.rglob("*")) if Path.home().parent.exists() else []
    assert after == before
    assert "ПЛАН:" in capsys.readouterr().out


def test_foreign_protocol_never_overwritten_or_deleted(link_install, fake_registry):
    fake_registry.snapshot = protocol.tree('"C:\\foreign.exe" "%1"')
    before = copy.deepcopy(fake_registry.snapshot)
    assert cli.setup_cmd(["--links", "--force"]) == 1
    assert cli.setup_cmd(["--no-links"]) == 1
    assert fake_registry.snapshot == before and set(fake_registry.calls) == {"read"}
    assert not state.state_path().exists()


def test_modified_owned_tree_is_preserved(link_install, fake_registry):
    assert cli.setup_cmd(["--links"]) == 0
    fake_registry.snapshot["values"]["custom"] = ("keep", 1)
    assert cli.setup_cmd(["--no-links"]) == 1
    assert fake_registry.snapshot["values"]["custom"] == ("keep", 1)
    assert state.read_state()["registrations"]


def test_update_migrates_tracked_path_preserving_policy(link_install, tmp_path, fake_registry):
    root, exe = link_install
    assert cli.setup_cmd(["--links", "--allowed-dirs", str(root)]) == 0
    moved = tmp_path / "new" / "app" / exe.name
    moved.parent.mkdir(parents=True)
    moved.with_name(protocol.GUI_EXE).write_text("fixture GUI")
    assert protocol.configure(server=(str(moved), []))
    assert str(moved.with_name(protocol.GUI_EXE)) in protocol.command_of(fake_registry.snapshot)
    record, = state.read_state()["registrations"]
    assert record["allowed_dirs"] == [str(root)] and record["root"] == str(moved.parent.parent)
    assert not protocol.configure(False, server=(str(exe), []))  # old copy cannot unregister the new one
    assert protocol.configure(False, server=(str(moved), []))


def test_state_pending_survives_registry_write_failure(link_install, fake_registry, monkeypatch):
    def failed(command):
        raise OSError("fixture denied")

    monkeypatch.setattr(fake_registry, "write", failed)
    assert cli.setup_cmd(["--links"]) == 1
    assert state.read_state()["registrations"][0]["status"] == "pending"


def test_source_registration_uses_pythonw_module(tmp_path, monkeypatch, fake_registry):
    py = tmp_path / "venv with spaces" / "python.exe"
    py.parent.mkdir()
    py.with_name("pythonw.exe").write_text("fixture")
    monkeypatch.setattr(registrations, "identity", lambda: (str(py), ["server.py"], str(tmp_path)))
    assert protocol.configure()
    assert protocol.command_of(fake_registry.snapshot) == f'"{py.with_name("pythonw.exe")}" -m office_live open-link "%1"'


def test_uninstall_protocol_dry_run_then_removal(link_install, fake_registry, monkeypatch):
    root, _ = link_install
    install._write_marker(root)
    install.write_manifest(root / "app")
    assert cli.setup_cmd(["--links"]) == 0
    before = state.state_path().read_bytes()
    assert install.uninstall_cmd(["--target", str(root), "--keep-files", "--dry-run"]) == 0
    assert state.state_path().read_bytes() == before and fake_registry.snapshot
    monkeypatch.setattr(install, "frozen", lambda: False)
    assert install.uninstall_cmd(["--target", str(root)]) == 0
    assert fake_registry.snapshot is None and not state.read_state()["registrations"] and not root.exists()


def test_gui_identity_loads_saved_policy_without_client_environment(link_install, monkeypatch):
    root, exe = link_install
    assert cli.setup_cmd(["--links", "--allowed-dirs", str(root)]) == 0
    monkeypatch.setattr(sys, "executable", str(exe.with_name(protocol.GUI_EXE)))
    monkeypatch.setattr(config, "SETTINGS", config.load({}))
    assert protocol.allowed_settings().allowed_dirs == (root,)


def test_saved_policy_is_enforced_in_handler(link_install, monkeypatch):
    from office_live import com, navigation

    root, _ = link_install
    assert cli.setup_cmd(["--links", "--allowed-dirs", str(root)]) == 0
    monkeypatch.setattr(config, "SETTINGS", config.load({}))
    monkeypatch.setenv("OFFICE_LIVE_LINK_NO_DIALOG", "1")
    from types import SimpleNamespace as NS

    doc = NS(FullName=r"C:\outside\Doc.docx", Path=r"C:\outside", Name="Doc.docx", Activate=lambda: pytest.fail("No navigation outside policy"))
    monkeypatch.setattr(navigation, "resolve", lambda *a: (None, doc, None, None))
    monkeypatch.setattr(com, "run_com", lambda fn, *a, **kw: fn())
    assert protocol.open_link([links.build("word", doc=doc.FullName)]) == 1


def test_corrupt_protocol_state_refuses_without_registry_write(link_install, fake_registry):
    assert cli.setup_cmd(["--links"]) == 0
    data = state.read_state()
    data["registrations"][0]["allowed_dirs"] = ["relative"]
    state.state_path().write_text(json.dumps(data), encoding="utf-8")
    fake_registry.calls.clear()
    assert cli.setup_cmd(["--links"]) == 1
    assert not fake_registry.calls


def test_native_adapter_uses_only_fixed_hkcu_paths_with_fake_winreg(monkeypatch):
    """Exercise the actual thin adapter against an in-memory winreg implementation."""
    from contextlib import contextmanager
    from types import SimpleNamespace

    values, calls = {}, []

    @contextmanager
    def create(hive, path, reserved, access):
        assert hive == "FAKE_HKCU" and reserved == 0 and access == 2
        parts = path.split("\\")
        for i in range(1, len(parts) + 1):
            values.setdefault("\\".join(parts[:i]), {})
        calls.append(("create", path))
        yield path

    @contextmanager
    def opened(hive, path):
        full = path if hive == "FAKE_HKCU" else hive + "\\" + path
        if full not in values:
            raise FileNotFoundError(full)
        yield full

    def children(path):
        return [k[len(path) + 1:] for k in values if k.startswith(path + "\\") and "\\" not in k[len(path) + 1:]]

    def set_value(path, name, reserved, kind, value):
        assert reserved == 0 and kind == 1
        values[path][name] = (value, kind)

    def delete(hive, path):
        assert hive == "FAKE_HKCU" and not children(path)
        calls.append(("delete", path))
        del values[path]

    fake = SimpleNamespace(HKEY_CURRENT_USER="FAKE_HKCU", REG_SZ=1, KEY_WRITE=2, CreateKeyEx=create, OpenKey=opened,
                           SetValueEx=set_value, QueryInfoKey=lambda p: (len(children(p)), len(values[p]), 0),
                           EnumKey=lambda p, i: children(p)[i],
                           EnumValue=lambda p, i: (list(values[p])[i], *list(values[p].values())[i]), DeleteKey=delete)
    monkeypatch.setitem(sys.modules, "winreg", fake)
    native = NATIVE_REGISTRY()
    assert native.read() is None
    native.write('"fixture.exe" open-link "%1"')
    assert native.read() == protocol.tree('"fixture.exe" open-link "%1"')
    native.delete()
    assert native.read() is None
    assert all(path == protocol.KEY or path.startswith(protocol.KEY + "\\") for _, path in calls)


def test_concurrent_foreign_registration_is_not_overwritten(link_install, fake_registry, monkeypatch):
    original = state.write_state

    def race(data):
        original(data)
        fake_registry.snapshot = protocol.tree('"foreign.exe" "%1"')

    monkeypatch.setattr(protocol, "write_state", race)
    assert cli.setup_cmd(["--links"]) == 1
    assert protocol.command_of(fake_registry.snapshot) == '"foreign.exe" "%1"'
    assert "write" not in fake_registry.calls
