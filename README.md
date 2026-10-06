# Office Live MCP

[![CI](https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/checks.yml/badge.svg)](https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/checks.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![Python](https://img.shields.io/badge/python-3.10%E2%80%933.13-blue)

[![English](https://img.shields.io/badge/English-README-0969da?style=for-the-badge)](README.md)
[![Русский](https://img.shields.io/badge/%D0%A0%D1%83%D1%81%D1%81%D0%BA%D0%B8%D0%B9-README-cf222e?style=for-the-badge)](README.ru.md)

An open-source (MIT) MCP server that lets **your MCP agent** — Claude Code, Claude Desktop, Cursor, Codex CLI, ZCode,
VS Code… — work in the Excel and Word documents **you already have open**. Changes appear on screen immediately, and you
keep working next to the agent. It connects to running Excel and Word through COM — no Office add-in to install.

## Why choose Office Live MCP?

[Claude for Excel/Word](https://claude.com/claude-for-microsoft-365) and
[ChatGPT for Excel](https://help.openai.com/en/articles/20001063-chatgpt-for-excel) / [Word](https://help.openai.com/en/articles/20001526-chatgpt-for-word)
bring their assistants into Office. Office Live MCP gives **your existing agent** access to Office, with these advantages:

- **Choose your agent and model.** Use a compatible MCP client and the models it supports; keep your existing development tools, instructions and workflow when working in Excel and Word.
- **Own and extend the integration.** The source is open under MIT: inspect how documents are accessed, add tools and adapt the server to your team's workflow. There is no server subscription fee.
- **Set access limits on the server.** Choose read-only mode, allowed folders, enabled tool groups and explicit document targeting. These settings are enforced by the server rather than depending only on instructions to the model.
- **Automate Word ↔ Excel document workflows.** Transfer tables, ranges and charts; generate a Word/PDF document for each Excel row from your template, with a read-only preview of file names, placeholders and collisions. Text identifiers such as `007` stay text.
- **Review and undo supported changes.** Preview targeted Word edits and apply them as tracked revisions; inspect the local change journal and undo supported agent actions. Undo refuses conflicting later edits instead of silently replacing them.

**Simple setup:** one Windows installer, without Python, git, administrator rights or an Office add-in.
The server runs locally; data read by your agent goes to that agent's model provider, and agent/model costs follow its provider's terms.
[Detailed comparison with Claude, ChatGPT and other MCP servers](docs/DETAILS.md#how-it-compares).

https://github.com/user-attachments/assets/b5911b0d-fa66-4d23-a888-ed7adeeeea07

<sub>30-second tour: real tool calls on real Excel and Word. All data is fictional.</sub>

## Highlights

- **Works in what you have open** — sees unsaved edits, recalculated formulas, pivots, charts and your current selection.
- **Excel and Word, 123 tools** — ranges, formulas, clean-up, goal seek, subtotals, pivots with slicers, charts; Word structure, templates, tracked edits, comment threads, document compare.
- **Bridges** — Word tables → Excel with real numbers, Excel ranges and charts → Word, one Word/PDF document per Excel row with a read-only preview first.

## Install

**Requirements:** Windows 10/11 (64-bit) and desktop Excel and/or Word.

1. Download the **Windows installer (the file ending in `-setup.exe`)** from the [**latest release**](https://github.com/mikhalchankasm/office-live-mcp/releases/latest).
2. Run it: **Next → Next → Finish** — on a fresh installation, detected agents are already selected. Keep them selected to connect all detected agents, and choose **full** or **read-only** access.
3. **Restart your agent** and ask: *"show me which workbooks are open in Excel"*.

**Update:** run the new setup.exe over the existing installation, then restart the agent.
**Uninstall:** Windows **Settings → Apps → Office Live MCP → Uninstall**.

## Learn more

[![Full documentation](https://img.shields.io/badge/Full_documentation-%E2%86%92-0969da?style=for-the-badge)](docs/DETAILS.md)

Silent install and setup options, the fallback ZIP, every feature, safety details,
comparison with Office AI add-ins and other MCP servers, development — [docs/DETAILS.md](docs/DETAILS.md).
Tool catalog: [docs/TOOLS.md](docs/TOOLS.md) · Changes: [CHANGELOG.md](CHANGELOG.md) · License: [MIT](LICENSE)
