# Changelog

## Unreleased

- `bridge_excel_to_word_documents` no longer flashes Word windows over the user's work. In the demo (5 rows,
  `export_pdf=True`) every generated document opened in a visible Word window on top of Excel, first as the template with
  unfilled `{{Manager}}`/`{{Region}}` placeholders, then as the filled letter, for the whole ~8.5 s call. Documents are
  now generated in a separate hidden Word instance (`DispatchEx`) as windowless documents. That instance quits when the
  call ends, also after a failure, so no hidden documents or `~$` lock files are left behind. The user's Word, its
  documents, the active window and the foreground window (usually Excel) stay as they were. If Word still takes the
  foreground, it is handed back, but only when Word took it. The result contract (`files`, `pdf_files`,
  `unfilled_placeholders`, …) is unchanged.
- Why a separate instance: windowless documents in the user's Word are unreliable. After the user closes a visible
  document, `Close` silently fails for every other hidden document, which stay open with their lock files (checked
  live). Opening documents visibly and hiding them at once (the 0.4.1 `word_compare_documents` pattern) still flashes a
  window. Cost: about 2 s to start that Word per call (5 rows with PDF took ~5.8 s instead of ~3.5 s on the test machine).
- Unit regressions with COM fakes (`tests/test_demo_fixes.py`), plus live scenarios in `tests/live/test_live_demo_fixes.py`:
  no Word window appears during the merge, the foreground and the user's active document are unchanged, and a template
  that is open in the user's Word works.

## 0.4.2 — 2026-10-06

A normal Windows installer instead of unpacking a zip and running install.cmd.

- Установщик setup.exe (Inno Setup): установка без администратора, русский и английский мастер, выбор агентов после
  копирования, ярлыки подключения/диагностики, обновление поверх и удаление через «Приложения». ZIP остаётся запасным вариантом.
- `clients --json` и `postinstall` для установщика: проверка MCP, учёт установки, добавление/снятие клиентов с сохранением
  настроек подключённых агентов. `setup` и `doctor` поддерживают `--pause`; `uninstall` делегирует удаление Inno и очищает
  корень в state при `--keep-files`, сохраняя состояние и резервные копии.
- Сборка и CI выпускают setup.exe и ZIP; реальная тихая установка/обновление/удаление проверяются только на CI runner.
- Исправлено: корень установки сравнивается и в короткой (8.3), и в длинной форме пути — `uninstall --keep-files` и
  самоудаление установки из ZIP больше не спотыкаются о пути вида `C:\Users\RUNNER~1\…`.

## 0.4.1 — 2026-10-05

Fixes for defects reproduced on real Excel/Word (Russian-locale Office 365) while filming the demo on 2026-10-05.

- `office_undo` really undoes `excel_conditional_format`: an added rule is deleted instead of relying on a cross-workbook
  PasteSpecial that leaves it in place; `clear` snapshots every affected rule whole, so a partly cleared rule comes back
  unsplit. Rules are checked after the restore, and undo fails with an explicit error instead of reporting `undone`
  when they cannot be matched.
- Number formats are written in English notation with LCID 1033 (as VBA does) by `excel_format_range`, `excel_clean_text`
  (`date_number_format`, `General` for converted numbers), `excel_split_column` and pivot value fields. In Russian Excel
  `dd.mm.yyyy`, `DD.MM.YYYY`, `dd.mm`, the `date`/`time`/`general` shortcuts, `[Red]` and `[h]:mm` used to be written as
  literal text or rejected. Formats already in local notation (`ДД.ММ.ГГГГ`) keep working. `excel_get_format`, template
  profiles and `excel_pivot_info` report formats in the same English notation. Chart formats keep the local-separator
  path (Excel stores them verbatim; checked live).
- `excel_create_chart` with a pivot source (its name, or a range inside it) binds to that pivot even when the active cell
  is in another pivot: the source's top-left cell is selected while the chart is added, then the user's view is restored.
  A chart that fails or binds to the wrong pivot is deleted. Undo removes a pivot chart from the pivot's sheet.
- `excel_manage_slicers` `add`/`connect` check before any change that all pivots share one PivotCache and explain the
  limitation; a failure after the slicer cache was created deletes it, so `add` is atomic.
- `word_select` without `target` follows the arguments (`find_text` → find, `table_index` → table, otherwise paragraphs)
  and refuses arguments the chosen target would ignore. Previously `find_text` alone silently selected paragraph 1.
- `bridge_excel_range_to_word_table` with `keep_formatting` reads and applies uniformly formatted rows in one call each and
  builds the table with Word screen updating off. The cell-look step of a 44x5 banded range takes ~4.6 s instead of
  ~17.5 s; mixed rows and hidden columns still go cell by cell.
- Unit regressions with COM fakes and live scenarios (`tests/live/test_live_demo_fixes.py`, `test_live_undo.py`).
- Fix `word_compare_documents` changing the user's active document (usually the original source) when a closed file is
  opened as a temporary copy: a hidden `Documents.Open` makes Word add footnote/endnote separators and header/footer styles
  to the active document without clearing its saved flag. The copy now opens visible with screen updating off, its window
  is hidden at once and the previously active window is restored; a copy whose window cannot be hidden is closed.
- Add `office_window` (`window`, `ui`, available in readonly and all presets): status, verified foreground activation,
  Office/chat or Excel/Word layouts, maximize and session-local restore. Use exact open COM targets, physical pixels,
  temporary thread DPI awareness v2, monitor work areas and DWM frame compensation with actual rectangle read-back.
- Detect chat conservatively through process ancestry, inherited consoles, unique Windows Terminal candidates or an
  explicit hint. Exclude Office, shell, hidden, cloaked, owned and tool windows; report ambiguity without moving chat.
  Report hung/modal/busy windows, focus denial and constrained layouts. Preserve placements for rollback/restore.
- Remember successful Excel/Word Save As aliases only in the current server session, return links using the new path,
  and give the separate link handler a clear renamed/closed error without guessing similar filenames.
- Add fake Win32/DLL tests, native window/COM guards and Office type-library checks. Add opt-in live window tests using
  only fixture documents with chat disabled and restoration in finally. Full live run: 119/119 on real Excel and Word.
- Installer: folder renames during install/update retry when Windows briefly refuses access right after the MCP check
  (antivirus or the just-started exe holding the folder); a copy held by a running agent still gives up quickly.

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
