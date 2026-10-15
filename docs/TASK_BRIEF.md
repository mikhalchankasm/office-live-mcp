# Task brief

Every Issue that is worked on carries one brief in its body, at most 2 KB. An executor must be able to start from
the brief alone. A brief that does not fit in 2 KB is a task that should be split.

```markdown
## Outcome
<One sentence: what exists after this task that does not exist now.>

## Boundaries
- In scope: <the narrowest useful end-to-end slice; files/functions>
- Out of scope: <what must not be touched or generalised>

## Acceptance
<One observable fact the owner can check in under a minute. Technical-only, or needs live Office (L3)?>
**Measured on the unchanged tree:** <what happens today and why that proves the task is needed>

## Files
<Expected files, or the entry point.>

## Verify
<The exact L0 command for this task, e.g. python -m pytest tests/test_core.py -q>

## Links
<At most three: owning doc section, blocking Issue, prior PR.>
```

Rules

- Run the acceptance check on the unchanged tree before handing the brief over. It must fail for the reason the
  task exists. A good check also rejects a plausible wrong implementation.
- The acceptance asserts an exact value or behaviour, never mere presence, and does not depend on timing, UI
  language or window size (Office here runs in Russian and English).
- The brief is the contract. An executor that finds it wrong stops and says so; it does not reinterpret it.
- Edit the brief in place; its history is the Issue's edit history.
