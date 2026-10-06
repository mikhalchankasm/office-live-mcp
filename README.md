# Office Live MCP

[![CI](https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/checks.yml/badge.svg)](https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/checks.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![Python](https://img.shields.io/badge/python-3.10%E2%80%933.13-blue)

[![English](https://img.shields.io/badge/English-README-0969da?style=for-the-badge)](README.md)
[![Русский](https://img.shields.io/badge/%D0%A0%D1%83%D1%81%D1%81%D0%BA%D0%B8%D0%B9-README-cf222e?style=for-the-badge)](README.ru.md)

An MCP server that lets any AI agent (Claude, Cursor, Codex, ZCode, VS Code…) work in the Excel and Word documents
**you already have open**. Changes appear on screen immediately, and you keep working next to the agent.

https://github.com/user-attachments/assets/b5911b0d-fa66-4d23-a888-ed7adeeeea07

<sub>30-second tour: real tool calls on real Excel and Word. All data is fictional.</sub>

## Highlights

- **Works in what you have open** — sees unsaved edits, recalculated formulas, pivots, charts and your current selection.
- **Excel and Word, 114 tools** — ranges, formulas, clean-up, pivots with slicers, charts; Word structure, templates, track changes, document compare.
- **Bridges** — Word tables → Excel with real numbers, Excel ranges and charts → Word, mail merge to Word and PDF.
- **Safe by default** — read-only mode, allowed folders, nothing is saved automatically, per-document journal and undo.
- **Any MCP agent, one installer** — no Python, git or admin rights; everything runs locally.

## Install

**Requirements:** Windows 10/11 (64-bit) and desktop Excel and/or Word.

1. Download **`office-live-mcp-<version>-setup.exe`** from [**Releases**](https://github.com/mikhalchankasm/office-live-mcp/releases/latest).
2. Run it: **Next → Next → Finish** — pick your agents and **full** or **read-only** access.
3. **Restart your agent** and ask: *"show me which workbooks are open in Excel"*.

**Update:** run the new setup.exe over the existing installation, then restart the agent.
**Uninstall:** Windows **Settings → Apps → Office Live MCP → Uninstall**.

## Learn more

[![Full documentation](https://img.shields.io/badge/Full_documentation-%E2%86%92-0969da?style=for-the-badge)](docs/DETAILS.md)

Silent install and setup options, the fallback ZIP, every feature, windows and clickable links, safety details,
comparison with other Office MCP servers, development — [docs/DETAILS.md](docs/DETAILS.md).
All 114 tools: [docs/TOOLS.md](docs/TOOLS.md) · Changes: [CHANGELOG.md](CHANGELOG.md) · License: [MIT](LICENSE)
