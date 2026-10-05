# Changelog

## Unreleased

- Add `excel_compare_ranges`: position/key matching across workbooks and instances, values/types, formulas or direct formats,
  duplicate-key/header refusal, bounded differences and explicit unmatched ranges/headers. Available in readonly mode.
- Add `excel_clean_text`: preview by default without change logging, fixed cleanup order, conservative numeric conversion,
  formula/merge/pivot/protection guards, changed-cell runs, literal text read-back and number-format restoration.
- Add `word_compare_documents`: a new unsaved revision document, same-instance sources, hidden read-only temporary files,
  macro suppression and guaranteed cleanup attempts. Source revisions produce warnings; discard by closing the result.
- Add `word_sort_table`: up to three text/number/date keys, full preflight checks, locale-aware typed-key validation and native undo.
- Clean/sort prepare undo only after preflight succeeds and refuse to write when the undo snapshot cannot be made
  (with `OFFICE_LIVE_UNDO=0` they write without one); partial COM failures keep the history.
- Packaged `install.cmd` warns when the extracted `app\_internal` path exceeds 200 characters, then continues installation.
  Unit tests isolate the installer profile and use temporary folders, including the exact 200/201 boundary.
- Add unit/refusal/undo regressions, Office type-library signature checks and opt-in live scenarios for all four tools.
  Update the generated catalog and EN/RU documentation to 109 tools.
- CI can be started by hand (`workflow_dispatch`).
- Known limitation documented: the PyInstaller build does not start when the full path of its libraries exceeds
  260 characters (very deep extraction folder, long `--target` or user profile).

## 0.3.1 — 2026-10-05

- Installer: the informational `doctor` step is limited to 120 s, so a hung Excel/Word no longer stops the
  installation before the program check and the agent connection.
- README: language buttons, step-by-step install and update, a subjective rating column in the comparison and a
  project status note.

## 0.3.0 — 2026-10-05

- **Installer for a bare machine**: `office-live-mcp-<version>-win64.zip` with a self-contained `office-live-mcp.exe`
  (Python inside) and `install.cmd` that checks Windows and Office, installs per user without admin rights, verifies the
  installed copy and connects the agents. Safe updates (the old copy is kept until the new one starts) and an
  `uninstall` that removes only its own files.
- **Journal and undo**: a Markdown journal of every change per document, an optional `Лог` log sheet in Excel, and
  `office_undo` for the agent's changes (Word's own undo history; Excel snapshots and inverse operations), refusing when
  the user edited the same content afterwards unless `force=true`.
- Hardening after four independent reviews: readonly bypasses closed, slicer and pivot fixes, installer config handling
  (real TOML parsing, no lost restrictions), UTF-8 CLI output, CI on Python 3.10 and 3.13 with archive builds.

## 0.2.0 — 2026-10-04

- Package with ~100 typed tools for Excel, Word and Word↔Excel bridges, toolsets, read-only mode, allowed folders,
  audit log, strict targeting, AutoSave guard, one-click `install.cmd` from source and a `setup` dialog for agents.

## 0.1.0

- First single-file prototype with 15 tools.
