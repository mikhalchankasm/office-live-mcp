"""Release checks use temporary files and a closed fake command surface; no git writes/network."""

import importlib.util
import json
import subprocess
import zipfile
from pathlib import Path

import pytest

SPEC = importlib.util.spec_from_file_location("release", Path(__file__).resolve().parents[1] / "scripts/release.py")
r = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(r)
BASE, HEAD, MERGE = "a" * 40, "b" * 40, "c" * 40
TEXTS = ['[project]\nversion = "0.4.6"\n', '__version__ = "0.4.6"\n', '# Changelog\n\n## Unreleased\n\n- Fix.\n\n## 0.4.6\n\nOld.\n']
ARGS = ["0.4.7", "--summary", "Release summary", "--execute"]


@pytest.fixture
def fake(tmp_path, monkeypatch):
    for name, text in zip(r.FILES, TEXTS, strict=True):
        path = tmp_path / name
        path.parent.mkdir(exist_ok=True)
        path.write_text(text, encoding="utf-8")
    state = {"calls": [], "dirty": "", "local_tag": "", "remote_tag": "", "ci": "success", "ci_sha": BASE,
             "checks": "pass", "tag_ci": "success", "branch": "main", "head": HEAD, "pushed": False, "merged": False}

    def run(cmd, **kwargs):
        assert isinstance(cmd, list) and kwargs["cwd"] == tmp_path
        state["calls"].append(cmd)
        key = tuple(cmd[:3])
        output = ""
        if key == ("git", "status", "--porcelain"):
            output = state["dirty"]
        elif cmd[:2] == ["git", "show"]:
            output = (tmp_path / cmd[2].split(":", 1)[1]).read_text(encoding="utf-8")
        elif cmd[:2] == ["git", "rev-parse"]:
            output = {"origin/main": BASE, "HEAD": state["head"], "v0.4.7^{}": MERGE, "v0.4.7": "d" * 40}[cmd[2]]
        elif key == ("git", "branch", "--show-current"):
            output = state["branch"]
        elif key == ("git", "tag", "--list"):
            output = state["local_tag"]
        elif key == ("git", "ls-remote", "--tags"):
            output = state["remote_tag"]
        elif key == ("git", "ls-remote", "--heads"):
            output = f"{BASE}\trefs/heads/main" if cmd[-1] == "main" else (f"{HEAD}\trefs/heads/rel-0.4.7" if state["pushed"] else "")
        elif key == ("git", "switch", "-c"):
            state["branch"] = cmd[3]
        elif cmd[:2] == ["git", "push"]:
            if cmd[-1].startswith("refs/tags/"):
                state["remote_tag"] = f"{'d' * 40}\trefs/tags/v0.4.7\n{MERGE}\trefs/tags/v0.4.7^{{}}"
            else:
                state["pushed"] = True
        elif key == ("git", "tag", "-a"):
            state["local_tag"] = cmd[3]
        elif key == ("git", "cat-file", "-t"):
            output = "tag"
        elif key == ("gh", "run", "list"):
            output = [{"conclusion": state["ci"], "headSha": state["ci_sha"]}] if cmd[4] == "main" else [
                {"databaseId": 42, "conclusion": state["tag_ci"], "headSha": MERGE, "status": "completed", "event": "push"}]
        elif key == ("gh", "pr", "create"):
            output = "https://github.com/example/repo/pull/34"
        elif key == ("gh", "pr", "view"):
            output = {"number": 34, "headRefName": "rel-0.4.7", "headRefOid": HEAD, "baseRefName": "main",
                      "state": "MERGED" if state["merged"] else "OPEN", "mergeCommit": {"oid": MERGE}}
        elif key == ("gh", "pr", "checks"):
            output = [{"name": "CI", "state": "SUCCESS", "bucket": state["checks"]}]
        elif key == ("gh", "pr", "merge"):
            state["merged"] = True
        elif key == ("gh", "release", "view"):
            output = {"assets": [{"name": f"office-live-mcp-0.4.7-{suffix}"} for suffix in ("setup.exe", "win64.zip")]}
        else:
            assert cmd[:2] in (["git", "fetch"], ["git", "add"], ["git", "commit"], ["git", "branch"], ["git", "merge-base"]), cmd
        return subprocess.CompletedProcess(cmd, 1 if key == ("gh", "pr", "checks") and state["checks"] == "fail" else 0,
                                           output if isinstance(output, str) else json.dumps(output), "")

    monkeypatch.setattr(r, "run", run)
    monkeypatch.setattr(r.time, "sleep", lambda _: pytest.fail("unexpected wait"))
    return tmp_path, state


def test_plan(fake, capsys):
    root, state = fake
    assert r.main(["9.9.9"], root=root) == 0
    assert state["calls"] == [] and not (root / "artifacts").exists()
    assert "9. evidence" in capsys.readouterr().out


@pytest.mark.parametrize(("key", "value"), [("dirty", " M file"), ("local_tag", "v0.4.7"), ("remote_tag", "tag"),
                                         ("ci", "failure"), ("ci_sha", HEAD), ("version", "0.4.6"), ("notes", "## Unreleased\n\n## Old\n")])
def test_preflight_refuses(fake, capsys, key, value):
    root, state = fake
    args = ARGS.copy()
    if key == "version":
        args[0] = value
    elif key == "notes":
        (root / "CHANGELOG.md").write_text(value, encoding="utf-8")
    else:
        state[key] = value
    assert r.main(args, root=root) == 1
    assert "RESULT STOPPED step 1:" in capsys.readouterr().out
    assert not any(cmd[:2] == ["git", "switch"] for cmd in state["calls"])


def test_sequence_and_edits(fake, capsys):
    root, state = fake
    assert r.main(ARGS, root=root) == 0
    assert capsys.readouterr().out.rstrip().endswith(f"RESULT RELEASED v0.4.7 {MERGE}")
    assert 'version = "0.4.7"' in (root / r.FILES[0]).read_text()
    assert '__version__ = "0.4.7"' in (root / r.FILES[1]).read_text()
    assert f"## 0.4.7 — {r.date.today()}\n\nRelease summary\n\n- Fix." in (root / r.FILES[2]).read_text(encoding="utf-8")
    calls = state["calls"]
    assert ["gh", "pr", "merge", "34", "--squash", "--match-head-commit", HEAD] in calls
    tag = ["git", "tag", "-a", "v0.4.7", MERGE, "-m", "Office Live MCP 0.4.7"]
    assert calls.index(tag) > next(i for i, cmd in enumerate(calls) if cmd[:3] == ["gh", "pr", "merge"])
    directory = root / "artifacts/release/v0.4.7"
    assert len(list(directory.glob("*.log"))) == 9
    assert (directory / "pr-body.md").read_text(encoding="utf-8").rstrip().endswith(r.FOOTER)
    assert len(json.loads((directory / "evidence.json").read_text(encoding="utf-8"))["steps"]) == 9


@pytest.mark.parametrize("mismatch", [False, True])
def test_red_ci_and_resume(fake, capsys, mismatch):
    root, state = fake
    state["checks"] = "fail"
    assert r.main(ARGS, root=root) == 1
    assert "RESULT STOPPED step 4: red PR CI" in capsys.readouterr().out
    assert not state["merged"]
    state["checks"] = "pass"
    if mismatch:
        state["head"] = BASE
    assert r.main([*ARGS, "--from-step", "4"], root=root) == int(mismatch)
    assert state["merged"] != mismatch


def test_red_tag_ci(fake, capsys):
    root, state = fake
    state["tag_ci"] = "failure"
    assert r.main(ARGS, root=root) == 1
    assert "RESULT STOPPED step 7: red tag CI" in capsys.readouterr().out


@pytest.mark.parametrize("bad", [None, "setup.exe", "office-live-mcp.exe", "office-live-link.exe"])
def test_asset_versions(tmp_path, monkeypatch, bad):
    (tmp_path / "setup.exe").write_bytes(b"setup")
    with zipfile.ZipFile(tmp_path / "package.zip", "w") as archive:
        for name in ("office-live-mcp.exe", "office-live-link.exe"):
            archive.writestr(f"root/app/{name}", b"exe")
    monkeypatch.setattr(r, "file_version", lambda path: (0, 4, 6, 0) if path.name == bad else (0, 4, 7, 0))
    if bad:
        with pytest.raises(r.Stop, match="VersionInfo mismatch"):
            r.verify_assets(tmp_path, ["setup.exe", "package.zip"], "0.4.7")
    else:
        hashes = r.verify_assets(tmp_path, ["setup.exe", "package.zip"], "0.4.7")
        assert len(hashes) == 2 and all(len(value) == 64 for value in hashes.values())


def test_poll_pending_and_timeout(fake, monkeypatch):
    root, _ = fake
    release = r.Release(type("Args", (), {"version": "0.4.7", "summary": "s", "verify_assets": False, "from_step": 1})(), root)
    waits = []
    monkeypatch.setattr(r.time, "sleep", waits.append)
    ticks = iter([0, 1, 61])
    monkeypatch.setattr(r.time, "monotonic", lambda: next(ticks))
    with pytest.raises(r.Stop, match="timeout"):
        release.poll(lambda: False, 1)
    assert waits == [30]
    monkeypatch.setattr(r.time, "monotonic", lambda: 0)
    ready = iter([False, True])
    release.poll(lambda: next(ready), 30)
    assert waits == [30, 30]


def test_partial_push_cannot_be_replayed(fake, monkeypatch, capsys):
    root, state = fake
    original = r.run

    def fail_push(cmd, **kwargs):
        result = original(cmd, **kwargs)
        if cmd[:2] == ["git", "push"]:
            result.returncode = 1
        return result

    monkeypatch.setattr(r, "run", fail_push)
    assert r.main(ARGS, root=root) == 1
    assert "RESULT STOPPED step 3: partial change" in capsys.readouterr().out
    before = len(state["calls"])
    assert r.main([*ARGS, "--from-step", "3"], root=root) == 1
    assert len(state["calls"]) == before


def test_versioninfo_reader(monkeypatch, tmp_path):
    import sys
    from types import SimpleNamespace

    monkeypatch.setitem(sys.modules, "win32api", SimpleNamespace(GetFileVersionInfo=lambda path, key: {"FileVersionMS": 4, "FileVersionLS": 7 << 16}))
    assert r.file_version(tmp_path / "unused.exe") == (0, 4, 7, 0)
