# Office Live MCP

[![CI](https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/checks.yml/badge.svg)](https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/checks.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![Python](https://img.shields.io/badge/python-3.10%E2%80%933.13-blue)

[![English](https://img.shields.io/badge/English-README-0969da?style=for-the-badge)](README.md)
[![Русский](https://img.shields.io/badge/%D0%A0%D1%83%D1%81%D1%81%D0%BA%D0%B8%D0%B9-README-cf222e?style=for-the-badge)](README.ru.md)

> **Status: actively developed.** 0.x releases — the tools and the installer are tested on real Excel and Word, but the
> project is still being improved (the Excel + Word side-by-side layout is confirmed on real windows; chat layouts and
> restore are still covered only by fake tests; Power Query and Power Pivot/DAX later). Feedback and issues are welcome.

An MCP server that lets any AI agent (Claude, Cursor, Codex, ZCode, VS Code…) work in the Excel and Word documents
**you already have open** — the way Office's built-in assistants do. Changes appear on screen immediately, and you keep
working next to the agent.

It attaches to running Excel/Word through COM, so it sees exactly what you see: unsaved edits, recalculated formulas,
pivot tables, slicers, charts, the current selection.

https://github.com/user-attachments/assets/b5911b0d-fa66-4d23-a888-ed7adeeeea07

<sub>30-second tour: a Word table moved into Excel, an unsaved edit the agent sees, clean-up, a pivot with a chart, a Word
report from a template and mail merge — real tool calls on real Excel and Word. All data is fictional.</sub>

## Install

**Requirements:** Windows 10/11 (64-bit) and desktop Excel and/or Word. No Python, git, winget or admin rights.

| Step | What to do |
|:---:|---|
| **1** | Download **`office-live-mcp-<version>-setup.exe`** from [**Releases**](https://github.com/mikhalchankasm/office-live-mcp/releases/latest). |
| **2** | Run it: **Next → Next → Finish**. The wizard checks for Office, installs the program and lets you choose agents and **full** or **read-only** access. |
| **3** | **Restart your agent** and ask: *"show me which workbooks are open in Excel"*. |

## Update

| Step | What to do |
|:---:|---|
| **1** | Download the new **setup.exe** from [Releases](https://github.com/mikhalchankasm/office-live-mcp/releases/latest). |
| **2** | Run it over the existing installation. Connected agents keep their settings; the access choice applies to newly added agents. If prompted, close agents using the server. Uncheck an agent to disconnect its recorded registrations. |
| **3** | Restart your agent. |

**Uninstall:** Windows **Settings → Apps → Office Live MCP → Uninstall**. State, logs and configuration backups are retained.

The program lives in `%LOCALAPPDATA%\Programs\office-live-mcp\app`. Start menu shortcuts **Connect to agents** and
**Diagnostics (doctor)** let you connect later or check the installation. VS Code / project configurations use the console shortcut.

**Silent installation:**

```bat
office-live-mcp-<version>-setup.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /CLIENTS=codex,zcode /ACCESS=readonly
```

Without `/CLIENTS`, silent installation adds no agents and an update keeps existing registrations. `/NOLINKS` skips
`officelive://` registration. Use `setup --config PATH` for a custom file,
`--scope project --project DIR` for a project, or `setup --remove --clients zcode` to disconnect one client.
Registrations are tracked across scopes; conflicts preserve user changes. Configuration read-back confirms the entry,
but **restart the client** to load it. Portable mode is planned; all-users installation is unsupported.
See the [detailed guide and examples](docs/GUIDE.ru.md#установка-подключения-и-обновление).

The executable is not code-signed yet: SmartScreen may warn (*More info* → *Run anyway*), and Smart App Control or a
corporate antivirus may block it. From source: `install.cmd` in a clone, or `pip install .` and
`python -m office_live setup` — see the [guide](docs/GUIDE.ru.md).

### For administrators: fallback ZIP

The release also includes **`office-live-mcp-<version>-win64.zip`**. Extract the entire archive and run `install.cmd`
(or `install.cmd --yes --clients none --no-links`). Updates use the remembered directory; `--target` selects a custom root.
Existing 0.4.x ZIP installations can be upgraded with setup.exe without changing the `app\office-live-mcp.exe` path.
For installations managed by setup.exe, prefer setup.exe for future updates.

## What it can do

114 tools in groups you can switch on and off (`OFFICE_LIVE_TOOLSETS`, including the optional Python tool) — full list in [docs/TOOLS.md](docs/TOOLS.md).

| Group | Highlights |
|---|---|
| Excel | read/write ranges, **compare ranges by position or key**, **clean text, parse explicit dates and split columns with preview/undo**, **trace formula inputs/consumers with completeness reporting**, **sheet/workbook protection**, formulas (A1, R1C1, dynamic arrays), formatting, conditional formatting, validation, sort/filter, hide/group rows, names, tables, **pivot tables with filters, slicers and timelines**, charts, data profile, issue finder, sheet blueprint, range snapshot as PNG |
| Word | **compare documents into a new revision document**, **sort table rows by up to three keys with undo**, structure, reading by pages, exact find/replace, inserting at a bookmark or next to a table, styles, lists, tables, headers/footers, TOC, comments, track changes, footnotes, page snapshot as PNG, `{{placeholder}}` templates |
| Bridges | Word table → Excel (numbers stay numbers, `007` stays text), Excel range/chart → Word, mail merge Excel → Word/PDF, inspect a file without opening it |
| History | per-document change journal, optional `Лог` log sheet in Excel, `office_undo` for the agent's changes, **clickable links** to plans and results (`office_link`) |
| Windows | `office_window`: status, verified focus, Office/chat layouts with monitor/DPI/frame handling, restore the last layout; available in readonly |

Ask *“Arrange the windows: Excel left, chat right”* or *“Restore the previous window layout”*. Layouts run only on
request/consent; an ambiguous chat is left out. The `window` group is included in all presets; explicit group lists
can omit it. After Save As, use the returned new links: old-name aliases work only inside the same server session,
not in the separate browser link handler. [Window details and limitations](docs/GUIDE.ru.md#окна-office-и-чата).

Clickable `officelive://` links select a range or Word location in an **already open local file**. Installation registers
the handler for your Windows account; skip with setup.exe `/NOLINKS`, disable with `setup --no-links`, enable with
`setup --links`. The handler never opens files or edits their contents. For restricted folders use
`setup --links --allowed-dirs "D:\Work"`; the browser handler has its own saved policy, separate from client settings.
See [link formats and safety](docs/GUIDE.ru.md#кликабельные-ссылки).

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

Everything runs locally; the server opens no ports. What the agent reads ends up in the model's context — limit the folders
for confidential files. Details and limits: [guide](docs/GUIDE.ru.md) · [SECURITY.md](SECURITY.md).

## How it compares

The popular Office MCP servers edit **files on disk**; a few drive the **running** application.

| Project | Approach | Pros | Cons | Rating* |
|---|---|---|---|:---:|
| **Office Live MCP** (this) | Live COM, attaches to the files you have open; Excel **and** Word | sees unsaved edits and recalculated values; works next to you; pivots, slicers, charts; Word↔Excel bridges; read-only mode, allowed folders, journal, undo; clickable links to changed cells; one-file installer | Windows + desktop Office only; COM is slower than file editing; Excel's own Ctrl+Z history is cleared by automation; no Power Query/DAX yet | ★★★★★ |
| [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel) | Live COM (.NET), Excel only | very broad Excel coverage: Power Query, DAX, Power Pivot, VBA, 300+ operations | asks you to close your open workbooks first; no Word; Windows only | ★★★★☆ |
| [haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) | Files via openpyxl | cross-platform, no Excel needed, `uvx` install, HTTP transport, read-only and folder limits | no live workbooks; formulas are not recalculated; pivot "tables" are static summaries; no Word | ★★★☆☆ |
| [negokaz/excel-mcp-server](https://github.com/negokaz/excel-mcp-server) | Files (Go), live editing on Windows | `npx` install, cross-platform file mode, screenshots on Windows | small toolset (8 tools); no Word | ★★☆☆☆ |
| [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) | Files via python-docx | no Word needed, rich document creation, PDF export | no live documents, no track changes; archived (read-only) since March 2026 | ★★☆☆☆ |
| [OfficeMCP/OfficeMCP](https://github.com/OfficeMCP/OfficeMCP) | Live COM via generic `RunPython` | many Office apps (Outlook, PowerPoint, Access…) | no typed tools, no safety limits — the agent runs arbitrary code | ★★☆☆☆ |

\* *Subjective rating by the author of this project for one use case: an AI agent assisting a person in the Excel and
Word documents that are open right now. For headless or cross-platform file pipelines the order would be different —
there a file-based server is the better choice.*

## Development

```bat
python -m venv .venv && .venv\Scripts\pip install -r requirements.txt pytest ruff
.venv\Scripts\python -m pytest tests -q          :: unit tests, no Office needed
set OFFICE_LIVE_LIVE_TESTS=1 && .venv\Scripts\python -m pytest tests\live -q   :: real Excel/Word
.venv\Scripts\python packaging\build.py          :: build setup.exe + ZIP (needs pyinstaller and Inno Setup 6.7+)
```

Architecture, COM pitfalls and all settings: [docs/GUIDE.ru.md](docs/GUIDE.ru.md) (Russian). Changes: [CHANGELOG.md](CHANGELOG.md).

## License

[MIT](LICENSE)
