# Office Live MCP

[![CI](https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![Python](https://img.shields.io/badge/python-3.10%E2%80%933.13-blue)

**English** · [Русский](README.ru.md)

An MCP server that lets any AI agent (Claude, Cursor, Codex, ZCode, VS Code…) work in the Excel and Word documents
**you already have open** — the way Office's built-in assistants do. Changes appear on screen immediately, and you keep
working next to the agent.

It attaches to running Excel/Word through COM, so it sees exactly what you see: unsaved edits, recalculated formulas,
pivot tables, slicers, charts, the current selection.

## Install

Requirements: Windows 10/11 (64-bit) and desktop Excel and/or Word. No Python, git, winget or admin rights needed.

1. Download `office-live-mcp-<version>-win64.zip` from [Releases](https://github.com/mikhalchankasm/office-live-mcp/releases/latest)
   and **extract the whole archive**.
2. Double-click **`install.cmd`**. It checks Windows and Office, copies the program to
   `%LOCALAPPDATA%\Programs\office-live-mcp`, verifies it and asks which access to grant (full or read-only) and which
   agents to connect (detected ones are preselected).
3. Restart your agent and ask: *"show me which workbooks are open in Excel"*.

Update: run `install.cmd` from a newer archive (settings are kept). Uninstall:
`"%LOCALAPPDATA%\Programs\office-live-mcp\app\office-live-mcp.exe" uninstall`.
Unattended: `install.cmd --yes --clients claude-code,cursor --readonly`.

The executable is not code-signed yet: SmartScreen may warn ("More info" → "Run anyway"), and Smart App Control or a
corporate antivirus may block it. From source: `install.cmd` in a clone, or `pip install .` and
`python -m office_live setup` — see the [guide](docs/GUIDE.ru.md).

## What it can do

105 tools in groups you can switch on and off (`OFFICE_LIVE_TOOLSETS`) — full list in [docs/TOOLS.md](docs/TOOLS.md).

| Group | Highlights |
|---|---|
| Excel | read/write ranges, formulas (A1, R1C1, dynamic arrays), formatting, conditional formatting, validation, sort/filter, hide/group rows, names, tables, **pivot tables with filters, slicers and timelines**, charts, data profile, issue finder, sheet blueprint, range snapshot as PNG |
| Word | structure, reading by pages, exact find/replace, inserting at a bookmark or next to a table, styles, lists, tables, headers/footers, TOC, comments, track changes, footnotes, page snapshot as PNG, `{{placeholder}}` templates |
| Bridges | Word table → Excel (numbers stay numbers, `007` stays text), Excel range/chart → Word, mail merge Excel → Word/PDF, inspect a file without opening it |
| History | per-document change journal, optional `Лог` log sheet in Excel, `office_undo` for the agent's changes |

## Safety

- **Read-only mode** (`OFFICE_LIVE_MODE=readonly`): writing tools are not even registered.
- **Allowed folders** (`OFFICE_LIVE_ALLOWED_DIRS`): files elsewhere are invisible to the agent; nothing is saved
  automatically and existing files are never overwritten without `overwrite=true`.
- Workbooks with **AutoSave** (OneDrive/SharePoint) are refused for changes; Excel events and macros stay off while the agent writes.
- **Journal and undo**: every change is logged per document; `office_undo` reverts the agent's steps (Word's own undo
  history; Excel snapshots) and refuses when you edited the same content afterwards unless `force=true`.
- Optional audit log (`OFFICE_LIVE_AUDIT_LOG`) and strict targeting (`OFFICE_LIVE_STRICT_TARGET`).

Everything runs locally; the server opens no ports. What the agent reads ends up in the model's context — limit the folders
for confidential files. Details and limits: [guide](docs/GUIDE.ru.md) · [SECURITY.md](SECURITY.md).

## How it compares

The popular Office MCP servers edit **files on disk**; a few drive the **running** application.

| Project | Approach | Pros | Cons |
|---|---|---|---|
| **Office Live MCP** (this) | Live COM, attaches to the files you have open; Excel **and** Word | sees unsaved edits and recalculated values; works next to you; pivots, slicers, charts; Word↔Excel bridges; read-only mode, allowed folders, journal, undo; one-file installer | Windows + desktop Office only; COM is slower than file editing; Excel's own Ctrl+Z history is cleared by automation |
| [haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) | Files via openpyxl | cross-platform, no Excel needed, `uvx` install, HTTP transport, read-only and folder limits | no live workbooks; formulas are not recalculated; pivot "tables" are static summaries; no Word |
| [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel) | Live COM (.NET), Excel only | very broad Excel coverage: Power Query, DAX, Power Pivot, VBA, 300+ operations | asks you to close your open workbooks first; no Word; Windows only |
| [negokaz/excel-mcp-server](https://github.com/negokaz/excel-mcp-server) | Files (Go), live editing on Windows | `npx` install, cross-platform file mode, screenshots on Windows | small toolset (8 tools); no Word |
| [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) | Files via python-docx | no Word needed, rich document creation, PDF export | no live documents, no track changes; archived (read-only) since March 2026 |
| [OfficeMCP/OfficeMCP](https://github.com/OfficeMCP/OfficeMCP) | Live COM via generic `RunPython` | many Office apps (Outlook, PowerPoint, Access…) | no typed tools, no safety limits — the agent runs arbitrary code |

Pick a file-based server for headless or cross-platform pipelines; pick Office Live MCP when the agent should assist a
person in documents that are open right now.

## Development

```bat
python -m venv .venv && .venv\Scripts\pip install -r requirements.txt pytest ruff
.venv\Scripts\python -m pytest tests -q          :: unit tests, no Office needed
set OFFICE_LIVE_LIVE_TESTS=1 && .venv\Scripts\python -m pytest tests\live -q   :: real Excel/Word
.venv\Scripts\python packaging\build.py          :: build the installer archive (needs pyinstaller)
```

Architecture, COM pitfalls and all settings: [docs/GUIDE.ru.md](docs/GUIDE.ru.md) (Russian). Changes: [CHANGELOG.md](CHANGELOG.md).

## License

[MIT](LICENSE)
