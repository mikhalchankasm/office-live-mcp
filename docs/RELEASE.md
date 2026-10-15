# Release gate

A release is a squash-merged `chore(release): X.Y.Z` PR plus the tag `vX.Y.Z` on it; CI (`checks.yml`, tag build)
builds setup.exe and the zip and creates the GitHub release with generated notes. `scripts/release.py` (#26) runs the
steps below: it prints the plan by default, `--execute` runs it and stops at the first failure, `--from-step N`
resumes; every step writes its log under `artifacts/release/<tag>/` (git-ignored). Until it lands, the same steps by hand.

## 0. Preconditions

- `git fetch`: the version is not taken — another computer may have released meanwhile.
- `main` CI is green; every PR in the batch merged; `## Unreleased` in `CHANGELOG.md` lists them.
- Open Issues labelled for this release are closed or moved.

## 1. Release candidate (L2 + L3)

- `python packaging/build.py` on the candidate SHA; packaged tests green.
- Full live suite in the owner's window (`scripts/live_run.py` once #25 lands); targeted probes of the batch included.
  Any failure: fix in a PR, new candidate, rerun only the affected tests plus the smoke set. A failure that also
  happens on the previous release's code (A/B run from `git archive`) is environment: record it in an Issue, it does
  not block the release.

## 2. Tech-debt sweep

- Review the batch's changed areas; `ruff`, skipped/xfail tests, files over 60 KB, TODOs added in the batch.
- Each finding: fixed in the batch, or a `tech-debt` Issue.

## 3. Version PR

- Branch `rel-X.Y.Z`: version in `pyproject.toml` and `office_live/__init__.py`; `## Unreleased` → `## X.Y.Z — date`
  with a one-line summary; tool count in README/DETAILS if it changed.
- PR `chore(release): X.Y.Z`, green CI, squash merge.

## 4. Tag and publish

- `git tag vX.Y.Z <merge SHA>`, push the tag; wait for the tag CI run.
- The release has `office-live-mcp-X.Y.Z-setup.exe` and `office-live-mcp-X.Y.Z-win64.zip`; their VersionInfo is X.Y.Z.

## 5. After the release

- Owner: installs or updates and accepts it live — the only owner step.
- winget: new manifests prepared for the owner to submit (the owner submits).
- Delete merged branches, prune worktrees, archive executor reports in `docs/dev/`.
