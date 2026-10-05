# Changelog

## Unreleased

- Installer: the informational `doctor` step is limited to 120 s, so a hung Excel/Word no longer stops the
  installation before the program check and the agent connection.

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
