"""Release coordinator. Planning never runs commands or writes evidence."""

import argparse
import hashlib
import json
import re
import subprocess
import tempfile
import time
import zipfile
from datetime import date
from pathlib import Path

STEPS = ("preflight", "version-pr", "push-pr", "pr-checks", "merge", "tag", "tag-ci", "assets", "evidence")
PLAN = ("preflight: clean tree, fetch origin --tags, unused version/tag, Unreleased, green main CI",
        "version-pr: branch from origin/main, update versions and CHANGELOG, commit with co-author",
        "push-pr: push release branch and create PR with changelog notes",
        "pr-checks: poll every 30 seconds, stop on red checks or after 30 minutes",
        "merge: squash PR with --match-head-commit and record merge SHA",
        "tag: create annotated tag on merge SHA and push it",
        "tag-ci: poll every 30 seconds, stop on red CI or after 40 minutes",
        "assets: check setup.exe and zip; --verify-assets downloads, hashes and checks VersionInfo",
        "evidence: finalize logs and evidence.json")
FILES = ("pyproject.toml", "office_live/__init__.py", "CHANGELOG.md")
PATTERNS = (r'(?m)^version = "([^"]+)"$', r'(?m)^__version__ = "([^"]+)"$')
FOOTER = "🤖 Generated with [Claude Code](https://claude.com/claude-code)"


class Stop(RuntimeError):
    pass


class PartialChangeError(Stop):
    """A failed step may already have changed local or remote state; inspect its log."""


def require(condition, message):
    if not condition:
        raise Stop(message)


def run(cmd, *, cwd, timeout=120):
    return subprocess.run(cmd, cwd=cwd, capture_output=True, text=True, encoding="utf-8", timeout=timeout, shell=False)


def version(value):
    require(re.fullmatch(r"(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)", value), "expected X.Y.Z")
    return tuple(map(int, value.split(".")))


def release_files(texts, target, summary):
    current = [re.search(pattern, text) for pattern, text in zip(PATTERNS, texts[:2], strict=True)]
    require(all(current), "version fields missing")
    require(current[0][1] == current[1][1], "version fields disagree")
    require(version(target) > version(current[0][1]), "version is already occupied")
    section = re.search(r"(?ms)^## Unreleased\s*\n(.*?)(?=^## |\Z)", texts[2])
    require(section and section[1].strip(), "empty or missing Unreleased")
    notes = section[1].strip()
    updated = [re.sub(pattern, lambda m: m[0].replace(m[1], target), text) for pattern, text in zip(PATTERNS, texts[:2], strict=True)]
    updated.append(texts[2][:section.start()] + f"## {target} — {date.today().isoformat()}\n\n{summary}\n\n{notes}\n\n" + texts[2][section.end():])
    return updated, notes


def file_version(path):
    import win32api

    info = win32api.GetFileVersionInfo(str(path), "\\")
    return tuple(part for key in ("FileVersionMS", "FileVersionLS") for part in (info[key] >> 16, info[key] & 65535))


def verify_assets(directory, names, target):
    hashes = {name: hashlib.sha256((directory / name).read_bytes()).hexdigest() for name in names}
    expected = (*version(target), 0)
    require(file_version(directory / names[0]) == expected, "setup VersionInfo mismatch")
    with zipfile.ZipFile(directory / names[1]) as archive, tempfile.TemporaryDirectory() as tmp:
        for name in ("office-live-mcp.exe", "office-live-link.exe"):
            matches = [entry for entry in archive.infolist() if Path(entry.filename).name == name]
            require(len(matches) == 1, f"missing or ambiguous {name} in zip")
            path = Path(tmp) / name
            path.write_bytes(archive.read(matches[0]))
            require(file_version(path) == expected, f"{name} VersionInfo mismatch")
    return hashes


class Release:
    def __init__(self, args, root):
        self.args, self.root = args, root
        self.tag, self.branch = f"v{args.version}", f"rel-{args.version}"
        self.directory = root / "artifacts" / "release" / self.tag
        self.path = self.directory / "evidence.json"
        self.data = {"version": args.version, "summary": args.summary, "verify_assets": args.verify_assets, "steps": {}}
        self.step, self.changed = args.from_step, False

    def log(self, text):
        with (self.directory / f"{self.step:02d}-{STEPS[self.step - 1]}.log").open("a", encoding="utf-8") as stream:
            stream.write(text + "\n")

    def save(self):
        temporary = self.path.with_suffix(".tmp")
        temporary.write_text(json.dumps(self.data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        temporary.replace(self.path)

    def cmd(self, *cmd, mutate=False, allowed=(0,)):
        self.log(json.dumps(cmd, ensure_ascii=False))
        self.changed |= mutate
        result = run(list(cmd), cwd=self.root)
        self.log(f"exit={result.returncode}\n{result.stdout}\n{result.stderr}")
        require(result.returncode in allowed, f"command failed ({result.returncode}): {' '.join(cmd[:3])}")
        return result.stdout.strip()

    def query(self, *cmd, **kwargs):
        return json.loads(self.cmd(*cmd, **kwargs))

    def clean(self):
        require(not self.cmd("git", "status", "--porcelain", "--untracked-files=all"), "working tree is dirty")

    def pr(self):
        return self.query("gh", "pr", "view", str(self.data["pr"]), "--json", "number,headRefName,headRefOid,baseRefName,state,mergeCommit")

    def check_pr(self, merged=False):
        pr = self.pr()
        require((pr["number"], pr["headRefName"], pr["headRefOid"], pr["baseRefName"]) ==
                (self.data["pr"], self.branch, self.data["head_sha"], "main"), "PR evidence mismatch")
        require(pr["state"] == ("MERGED" if merged else "OPEN"), "unexpected PR state")
        if merged:
            require(pr["mergeCommit"]["oid"] == self.data["merge_sha"], "merge SHA mismatch")
        return pr

    def remote_tag(self):
        return self.cmd("git", "ls-remote", "--tags", "origin", f"refs/tags/{self.tag}", f"refs/tags/{self.tag}^{{}}")

    def validate(self, step):
        require(all(str(n) in self.data["steps"] for n in range(1, step)), "previous step evidence missing")
        self.clean()
        if step == 2:
            require(self.cmd("git", "rev-parse", "origin/main") == self.data["base_sha"], "base SHA mismatch")
            require(self.cmd("git", "ls-remote", "--heads", "origin", "main").split() ==
                    [self.data["base_sha"], "refs/heads/main"], "origin/main changed since preflight")
        if step >= 3:
            require(self.cmd("git", "branch", "--show-current") == self.branch, "branch mismatch")
            require(self.cmd("git", "rev-parse", "HEAD") == self.data["head_sha"], "HEAD SHA mismatch")
        if step >= 4:
            self.check_pr(merged=step >= 6)
            if step <= 5:
                require(self.cmd("git", "ls-remote", "--heads", "origin", self.branch).split() ==
                        [self.data["head_sha"], f"refs/heads/{self.branch}"], "origin branch SHA mismatch")
        if step >= 5:
            require(self.data["checked_sha"] == self.data["head_sha"], "checked SHA mismatch")
        if step >= 7:
            require(self.cmd("git", "cat-file", "-t", self.tag) == "tag", "tag is not annotated")
            require(self.cmd("git", "rev-parse", f"{self.tag}^{{}}") == self.data["merge_sha"], "local tag mismatch")
            remote = dict(line.split()[::-1] for line in self.remote_tag().splitlines())
            require(remote.get(f"refs/tags/{self.tag}^{{}}") == self.data["merge_sha"] and
                    remote.get(f"refs/tags/{self.tag}") == self.cmd("git", "rev-parse", self.tag), "origin tag mismatch")

    def poll(self, check, minutes):
        deadline = time.monotonic() + minutes * 60
        while True:
            if check():
                return
            require(time.monotonic() < deadline, f"CI timeout after {minutes} minutes")
            time.sleep(30)

    def preflight(self):
        self.clean()
        self.cmd("git", "fetch", "origin", "--tags", mutate=True)
        self.data["base_sha"] = self.cmd("git", "rev-parse", "origin/main")
        texts = [self.cmd("git", "show", f"origin/main:{name}") for name in FILES]
        release_files(texts, self.args.version, self.args.summary)
        require(not self.cmd("git", "tag", "--list", self.tag) and not self.remote_tag(), "tag already exists")
        require(not self.cmd("git", "branch", "--list", self.branch) and
                not self.cmd("git", "ls-remote", "--heads", "origin", self.branch), "release branch already exists")
        ci = self.query("gh", "run", "list", "--branch", "main", "--workflow", "CI", "-L", "1", "--json", "conclusion,headSha")
        require(ci and ci[0]["conclusion"] == "success" and ci[0]["headSha"] == self.data["base_sha"], "main CI is not green on origin/main")

    def prepare(self):
        self.cmd("git", "switch", "-c", self.branch, "origin/main", mutate=True)
        texts = [(self.root / name).read_text(encoding="utf-8") for name in FILES]
        updated, self.data["notes"] = release_files(texts, self.args.version, self.args.summary)
        for name, text in zip(FILES, updated, strict=True):
            (self.root / name).write_text(text, encoding="utf-8")
        self.cmd("git", "add", "--", *FILES, mutate=True)
        self.cmd("git", "commit", "-m", f"chore(release): {self.args.version}", "-m",
                 "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>", mutate=True)
        self.data["head_sha"] = self.cmd("git", "rev-parse", "HEAD")

    def publish_pr(self):
        self.cmd("git", "push", "-u", "origin", self.branch, mutate=True)
        body = self.directory / "pr-body.md"
        body.write_text(self.data["notes"] + "\n\n" + FOOTER + "\n", encoding="utf-8")
        url = self.cmd("gh", "pr", "create", "--base", "main", "--head", self.branch, "--title",
                       f"chore(release): {self.args.version}", "--body-file", str(body), mutate=True)
        pr = self.query("gh", "pr", "view", url, "--json", "number")
        self.data["pr"] = pr["number"]
        self.check_pr()

    def checks(self):
        self.check_pr()
        checks = self.query("gh", "pr", "checks", str(self.data["pr"]), "--json", "name,state,bucket", allowed=(0, 1, 8))
        require(not any(c["bucket"] in ("fail", "cancel") for c in checks), "red PR CI")
        return bool(checks) and all(c["bucket"] == "pass" for c in checks)

    def wait_pr(self):
        self.poll(self.checks, 30)
        self.data["checked_sha"] = self.data["head_sha"]

    def merge(self):
        require(self.checks(), "PR checks no longer green")
        self.cmd("gh", "pr", "merge", str(self.data["pr"]), "--squash", "--match-head-commit", self.data["head_sha"], mutate=True)
        pr = self.pr()
        require(pr["state"] == "MERGED" and pr["mergeCommit"], "PR not merged")
        self.data["merge_sha"] = pr["mergeCommit"]["oid"]

    def create_tag(self):
        require(not self.cmd("git", "tag", "--list", self.tag) and not self.remote_tag(), "tag already exists")
        self.cmd("git", "fetch", "origin", "main", mutate=True)
        self.cmd("git", "merge-base", "--is-ancestor", self.data["merge_sha"], "origin/main")
        self.cmd("git", "tag", "-a", self.tag, self.data["merge_sha"], "-m", f"Office Live MCP {self.args.version}", mutate=True)
        self.cmd("git", "push", "origin", f"refs/tags/{self.tag}", mutate=True)

    def tag_ci(self):
        runs = self.query("gh", "run", "list", "--branch", self.tag, "--workflow", "CI", "-L", "1", "--json", "databaseId,status,conclusion,headSha,event")
        if not runs:
            return False
        ci = runs[0]
        require(ci["headSha"] == self.data["merge_sha"] and ci["event"] == "push", "tag CI SHA/event mismatch")
        if ci["status"] != "completed":
            return False
        require(ci["conclusion"] == "success", "red tag CI")
        self.data["tag_ci"] = ci
        return True

    def assets(self):
        names = [f"office-live-mcp-{self.args.version}-{suffix}" for suffix in ("setup.exe", "win64.zip")]
        assets = self.query("gh", "release", "view", self.tag, "--json", "assets")["assets"]
        require(all(sum(a["name"] == name for a in assets) == 1 for name in names), "release assets missing")
        self.data["assets"] = assets
        if self.args.verify_assets:
            directory = self.directory / "assets"
            directory.mkdir(exist_ok=True)
            self.cmd("gh", "release", "download", self.tag, "--dir", str(directory), "--pattern", names[0], "--pattern", names[1], mutate=True)
            self.data["sha256"] = verify_assets(directory, names, self.args.version)

    def execute(self):
        self.directory.mkdir(parents=True, exist_ok=True)
        self.logging = True
        actions = (self.preflight, self.prepare, self.publish_pr, self.wait_pr, self.merge, self.create_tag,
                   lambda: self.poll(self.tag_ci, 40), self.assets, lambda: self.data.update(result="RELEASED"))
        started = False
        try:
            if self.args.from_step > 1:
                previous = json.loads(self.path.read_text(encoding="utf-8"))
                require(all(previous[key] == self.data[key] for key in ("version", "summary", "verify_assets")), "evidence options mismatch")
                require(previous.get("in_progress") is None, "incomplete step: inspect and reconcile evidence before resuming")
                require(not any(int(n) >= self.step for n in previous["steps"]), "step already completed")
                self.data = previous
            else:
                require(not self.path.exists(), "evidence already exists; use --from-step")
            for step in range(self.args.from_step, 10):
                self.step = step
                self.changed = False
                if self.step > 1:
                    self.validate(self.step)
                self.data["in_progress"] = self.step
                self.save()
                started = True
                actions[self.step - 1]()
                self.data["steps"][str(self.step)] = STEPS[self.step - 1]
                self.data["in_progress"] = None
                self.data.pop("stopped", None)
                self.save()
                started = False
                self.log("OK")
        except (Exception, KeyboardInterrupt) as error:
            if started:
                self.data["stopped"] = {"step": self.step, "reason": str(error)}
                if not self.changed:
                    self.data["in_progress"] = None
                self.save()
            if self.changed:
                raise PartialChangeError(f"partial change; inspect log: {error}") from error
            raise


def main(argv=None, *, root=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("version")
    parser.add_argument("--summary")
    parser.add_argument("--execute", action="store_true")
    parser.add_argument("--from-step", type=int, choices=range(1, 10), default=1)
    parser.add_argument("--verify-assets", action="store_true")
    args = parser.parse_args(argv)
    release = Release(args, Path(root) if root else Path(__file__).resolve().parents[1])
    try:
        version(args.version)
        if not args.execute:
            for n, name in enumerate(PLAN, 1):
                print(f"{n}. {name}")
            print(f"RESULT STOPPED step {args.from_step}: plan only; use --execute")
            return 0
        require(args.summary and args.summary.strip() and len(args.summary.splitlines()) == 1, "--summary must be one nonempty line")
        release.execute()
    except (Exception, KeyboardInterrupt) as error:
        reason = " ".join(str(error).splitlines()) or type(error).__name__
        if getattr(release, "logging", False):
            release.log(f"STOPPED: {reason}")
        print(f"RESULT STOPPED step {release.step}: {reason}")
        return 1
    print(f"RESULT RELEASED {release.tag} {release.data['merge_sha']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
