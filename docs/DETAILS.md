# Office Live MCP — full documentation

[← Back to the overview](../README.md) · [Русская версия](DETAILS.ru.md) · [All tools](TOOLS.md) · [Changelog](../CHANGELOG.md)

- [Status](#status)
- [Install, update, uninstall](#install-update-uninstall)
- [What it can do](#what-it-can-do)
- [Safety](#safety)
- [How it compares](#how-it-compares)
- [Development](#development)

## Status

**Actively developed.** 0.x releases — the tools and the installer are tested on real Excel and Word, but the project is
still being improved. Power Query and Power Pivot/DAX are planned for later releases. Feedback and issues are welcome.

## Install, update, uninstall

**Requirements:** Windows 10/11 (64-bit) and desktop Excel and/or Word. No Python, git, winget or admin rights.

| Step | What to do |
|:---:|---|
| **1** | Download **`office-live-mcp-<version>-setup.exe`** from [**Releases**](https://github.com/mikhalchankasm/office-live-mcp/releases/latest). |
| **2** | Run it: **Next → Next → Finish**. The wizard checks for Office, installs the program and lets you choose agents and **full** or **read-only** access. |
| **3** | **Restart your agent** and ask: *"show me which workbooks are open in Excel"*. |

### Update

| Step | What to do |
|:---:|---|
| **1** | Download the new **setup.exe** from [Releases](https://github.com/mikhalchankasm/office-live-mcp/releases/latest). |
| **2** | Run it over the existing installation. Connected agents keep their settings; the access choice applies to newly added agents. If prompted, close agents using the server. Uncheck an agent to disconnect its recorded registrations. |
| **3** | Restart your agent. |

**Uninstall:** Windows **Settings → Apps → Office Live MCP → Uninstall**. State, logs and configuration backups are retained.

The program lives in `%LOCALAPPDATA%\Programs\office-live-mcp\app`. Start menu shortcuts **Connect to agents** and
**Diagnostics (doctor)** let you connect later or check the installation. VS Code / project configurations use the console shortcut.

### Silent installation

```bat
office-live-mcp-<version>-setup.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /CLIENTS=codex,zcode /ACCESS=readonly
```

Without `/CLIENTS`, silent installation adds no agents and an update keeps existing registrations.
Use `setup --config PATH` for a custom file,
`--scope project --project DIR` for a project, or `setup --remove --clients zcode` to disconnect one client.
Registrations are tracked across scopes; conflicts preserve user changes. Configuration read-back confirms the entry,
but **restart the client** to load it. Portable mode is planned; all-users installation is unsupported.
See the [detailed guide and examples](GUIDE.ru.md#установка-подключения-и-обновление).

The executable is not code-signed yet: SmartScreen may warn (*More info* → *Run anyway*), and Smart App Control or a
corporate antivirus may block it. From source: `install.cmd` in a clone, or `pip install .` and
`python -m office_live setup` — see the [guide](GUIDE.ru.md).

### For administrators: fallback ZIP

The release also includes **`office-live-mcp-<version>-win64.zip`**. Extract the entire archive and run `install.cmd`
(or `install.cmd --yes --clients none`). Updates use the remembered directory; `--target` selects a custom root.
Existing 0.4.x ZIP installations can be upgraded with setup.exe without changing the `app\office-live-mcp.exe` path.
For installations managed by setup.exe, prefer setup.exe for future updates.

## What it can do

123 tools in groups you can switch on and off (`OFFICE_LIVE_TOOLSETS`, including the optional Python tool) — catalog in [TOOLS.md](TOOLS.md).

| Group | Highlights |
|---|---|
| Excel | read/write ranges, **compare ranges by position or key**, **clean text, parse explicit dates and split columns with preview/undo**, **trace formula inputs/consumers with completeness reporting**, **sheet/workbook protection, goal seek, subtotals and sparklines**, formulas (A1, R1C1, dynamic arrays), formatting, conditional formatting, validation, sort/filter, hide/group rows, names, tables, **pivot tables with filters, slicers and timelines**, charts, data profile, issue finder, sheet blueprint, range snapshot as PNG |
| Word | **fragment preview/apply as tracked changes and threaded comment metadata**, **compare documents into a new revision document**, **sort table rows by up to three keys with undo**, **insert saved documents, restrict editing and convert tables ↔ text**, structure, reading by pages, exact find/replace, inserting at a bookmark or next to a table, styles, lists, tables, headers/footers, TOC, comments, track changes, footnotes, page snapshot as PNG, `{{placeholder}}` templates |
| Bridges | Word table → Excel (numbers stay numbers, `007` stays text), Excel range/chart → Word, mail merge Excel → Word/PDF with a read-only plan, inspect a file without opening it |
| History | per-document change journal, optional `Лог` log sheet in Excel, `office_undo` for the agent's changes |

### Reproducible examples

[Register → Word/PDF](examples/register-to-acts.md) · [Comments → tracked edits](examples/comments-to-tracked-edits.md) · [Compare registers](examples/compare-registers.md) · [Range → report table](examples/range-to-report-table.md)

## Safety

- **Read-only mode** (`OFFICE_LIVE_MODE=readonly`): writing tools are not registered; multi-action tools keep only
  their read actions. The range snapshot (PNG) is the one read tool with a side effect: a temporary chart on the sheet
  marks the workbook as modified.
- **Allowed folders** (`OFFICE_LIVE_ALLOWED_DIRS`): files elsewhere are invisible to the agent; nothing is saved
  automatically and existing files are never overwritten without `overwrite=true`.
- Workbooks with **AutoSave** (OneDrive/SharePoint) are refused for changes; Excel events and macros stay off while the agent writes.
- **Journal and undo**: every change is logged per document; `office_undo` reverts the agent's steps (Word's own undo
  history; Excel snapshots) and refuses when you edited the same content afterwards unless `force=true`.
- Optional audit log (`OFFICE_LIVE_AUDIT_LOG`) and strict targeting (`OFFICE_LIVE_STRICT_TARGET`).
- Protection passwords are masked in logs/errors/history. Password operations create an undo barrier; password-free protection restores flags and Locked properties.

The server runs locally and opens no ports. What the agent reads is sent to its model provider — limit the folders
for confidential files. Details and limits: [guide](GUIDE.ru.md) · [SECURITY.md](../SECURITY.md).

## How it compares

Facts below come from each product's public documentation, **checked on 2026-10-06**. "Not confirmed" means the
documentation does not mention the feature — not that it is missing. Corrections are welcome.

### Office AI assistants (add-ins)

| | **Office Live MCP** | [Claude for Excel](https://claude.com/docs/office-agents/excel) / [Word](https://claude.com/docs/office-agents/word) | [ChatGPT for Excel](https://help.openai.com/en/articles/20001063-chatgpt-for-excel) / [Word](https://help.openai.com/en/articles/20001526-chatgpt-for-word) |
|---|---|---|---|
| Form | MCP server, local process, COM; no Office add-in | Office add-in (sidebar) | Office add-in (sidebar) |
| Agent and model | any MCP client and the model it uses | Claude | ChatGPT; Codex in the ChatGPT desktop app can work with an open Excel workbook through the add-in |
| Apps | Excel, Word | Excel, Word, PowerPoint; Outlook in beta ([overview](https://claude.com/claude-for-microsoft-365)) | Excel, Google Sheets, Word, PowerPoint |
| Platforms and versions | Windows, desktop Office. Tested on Microsoft 365 Apps (version 16), Windows 11, Russian UI; other versions not verified | web, Windows (Microsoft 365 builds listed in the docs), Mac; not Office 2016/2019 perpetual | Excel desktop and web; Word: not specified |
| Word tracked changes | yes: track-changes control, fragment edits recorded as revisions | yes: tracked changes mode | not confirmed |
| Word comment threads | list threads with replies, reply, resolve | works through threads, edits the anchored text and replies | not confirmed |
| Excel pivot tables | create, fields, filters, slicers, timelines | edits pivot tables | not confirmed |
| Macros / VBA | VBA source is read from the file, never run | not supported | "may not be fully supported" |
| Price | server is free (MIT); the agent/model is billed by its provider | paid Claude plans | all ChatGPT plans, Free with limited usage |
| Where data goes | the server runs locally and opens no ports; what the agent reads goes to that agent's model provider | processed by Anthropic; inputs/outputs deleted within 30 days per the docs | per ChatGPT plan; some logs may be kept 30 days per the docs |
| Source code | open, MIT | proprietary | proprietary |

Working in the open document, cell references, pivot tables, Word revisions and context shared between Excel and Word
exist in these assistants too. Office Live MCP is for people who want **their own agent** (Claude Code, Cursor, Codex
CLI, ZCode, VS Code…) to work in Office, with server-side limits they control.

### MCP servers for Office

| Project | How it works | Office documents | Notes from its README |
|---|---|---|---|
| **Office Live MCP** | COM to running Excel and Word | the files you have open | Word↔Excel bridges, journal, undo of supported actions, read-only mode, allowed folders |
| [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel) (MIT) | COM to Excel (.NET), MCP server or CLI | asks to close open workbooks first (exclusive access) | Excel only; Power Query, DAX, VBA, 326 operations; Excel 2016 or later |
| [haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) (MIT) | files via openpyxl, no Excel needed | files on disk | cross-platform; formulas stored, not calculated; summary tables instead of real PivotTables; folder confinement, read-only mode |
| [negokaz/excel-mcp-server](https://github.com/negokaz/excel-mcp-server) (MIT) | files; live editing on Windows | files; open workbooks on Windows | 7 tools listed; screen capture on Windows |
| [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) (MIT) | files via python-docx | files on disk | repository archived (read-only) |
| [OfficeMCP/OfficeMCP](https://github.com/OfficeMCP/OfficeMCP) | COM via a generic `RunPython` tool | running Office apps | many Office apps; the agent runs arbitrary Python, no license file |

File-based servers are the better choice for headless or cross-platform pipelines; COM servers need Windows and
desktop Office.

## Development

```bat
python -m venv .venv && .venv\Scripts\pip install -r requirements.txt pytest ruff
.venv\Scripts\python -m pytest tests -q          :: unit tests, no Office needed
set OFFICE_LIVE_LIVE_TESTS=1 && .venv\Scripts\python -m pytest tests\live -q   :: real Excel/Word
.venv\Scripts\python packaging\build.py          :: build setup.exe + ZIP (needs pyinstaller and Inno Setup 6.7+)
```

Architecture, COM pitfalls and all settings: [GUIDE.ru.md](GUIDE.ru.md) (Russian). Changes: [CHANGELOG.md](../CHANGELOG.md).
License: [MIT](../LICENSE).
