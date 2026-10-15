"""The agent rules stay short, single-sourced and complete (AGENTS.md is the only mandatory document)."""

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def read(name: str) -> str:
    return (ROOT / name).read_text(encoding="utf-8")


def test_agents_md_is_short_and_has_its_sections():
    text = read("AGENTS.md")
    assert len(text.encode("utf-8")) <= 9 * 1024, "AGENTS.md is read before every task: keep it under 9 KB"
    for heading in ("## 1. Safety boundary", "## 2. The loop", "## 3. What to read", "## 4. Running a task",
                    "## 5. Verification levels", "## 6. Where things live", "## 7. Roles"):
        assert heading in text


def test_owner_office_rule_is_pinned():
    # Owner decision 2026-10-15: only a hung Office started by the run itself may be closed, never theirs.
    text = read("AGENTS.md")
    assert "Never close, quit or kill an Excel/Word the owner started" in text
    assert "Never select a process by image name" in text


def test_claude_md_is_only_a_pointer():
    text = read("CLAUDE.md")
    assert len(text.encode("utf-8")) <= 1536 and "AGENTS.md" in text
    assert "## " not in text, "rules live in AGENTS.md; CLAUDE.md must not grow its own sections"


def test_brief_and_pr_templates_carry_their_sections():
    brief = read("docs/TASK_BRIEF.md")
    for section in ("## Outcome", "## Boundaries", "## Acceptance", "**Measured on the unchanged tree:**", "## Files",
                    "## Verify", "## Links"):
        assert section in brief
    pr = read(".github/pull_request_template.md")
    for section in ("Fixes #", "## What changed", "## Verification", "## Findings"):
        assert section in pr


def test_release_gate_lists_every_step():
    gate = read("docs/RELEASE.md")
    for step in ("## 0. Preconditions", "## 1. Release candidate", "## 2. Tech-debt sweep", "## 3. Version PR",
                 "## 4. Tag and publish", "## 5. After the release"):
        assert step in gate
