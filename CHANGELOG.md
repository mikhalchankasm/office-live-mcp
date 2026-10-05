# Changelog

## 0.4.0

- Add clickable `officelive://` locations and read-only `office_link` (113 tools). Return bounded links for Excel writes,
  supported Word locations, search, comparison, issue reports, formula tracing and document headings; link Markdown journals
  and the optional Excel log sheet without including log hyperlinks in undo snapshots.
- Add strict local/open-target navigation, exact matching across instances, saved folder policy, malformed/malicious URI
  refusal and short Windows error dialogs. Package `office-live-link.exe` without a console alongside the MCP executable.
- Register the per-user protocol on install, support `install --no-links`, `setup --links/--no-links`, dry-run, state tracking,
  ownership checks, update and uninstall. Unit tests use a fake registry and block native writes; EXE acceptance skips
  protocol registration on isolated profiles. Add opt-in live link tests (not run) and a timed invalid-link GUI build smoke.

- Add `excel_split_column`: Python delimiter/fixed-width parsing, preview, full-output preflight and undo, literal text
  verification and conservative numeric conversion. Refuse occupied output, formulas/spills, merges, tables and pivots.
- Extend `excel_clean_text` with strict `text_to_date`, explicit formats/time, Excel 1900/1904 serial dates, invalid-date
  reporting and reversible number formats.
- Add read-only `excel_trace_formula`: bounded local COM references plus formula tokenization, cross-sheet dependents,
  range names and simple structured columns. Report incomplete graphs for dynamic/external/unsupported references,
  inactive-sheet COM, ambiguous COM failures and limits; never activate sheets or draw arrows.
- Add `excel_protection`: readonly status, sheet permissions, editable Locked ranges and workbook structure protection.
  Password-free actions support inverse undo; password actions create explicit barriers. Redact password-named arguments
  in audit (including positional calls), both journals, undo history and errors.
- Add fake/refusal/privacy/undo tests, type-library protection/reference checks and opt-in stage 2 live tests (not run).
  Update documentation and generated catalog to 112 tools. No release publication is performed by this change.

- Track installation root/version and client registrations in per-user `state.json`; custom-path updates reuse the existing
  installation, preserve all client settings and never auto-connect newly detected clients.
- Add project/custom scopes, `CODEX_HOME`, conflict protection (`--force`), config locks, atomic writes and read-back.
  Preserve additional MCP settings; independent setup calls support different access modes for Codex and ZCode.
- Add fingerprint-aware disconnect/uninstall across all recorded scopes, file manifests, reparse/boundary checks,
  preservation of user files, and a background deletion report. Keep the EXE when dependent registrations cannot be removed.
- Add side-effect-free `--dry-run` to install/setup/uninstall. Check staged and installed copies over MCP; preserve rollback
  copies until verification succeeds. Installer checks no longer call `doctor` or access running Office.
- Normalize stderr to UTF-8 in source/frozen entry points. Drain diagnostic bytes independently; strictly validate stdout
  as UTF-8 JSON-RPC with bounded RPC/EOF waits. Build smoke checks full/readonly modes on temporary profiles.
- Isolate every unit-test profile variable and add lifecycle acceptance, conflict, concurrent-change, rollback, timeout,
  Unicode/long-path and dry-run regressions. Build in a short temporary directory; document unsupported portable/all-users modes.

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
  Update the generated catalog and EN/RU documentation for stage 1.
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
