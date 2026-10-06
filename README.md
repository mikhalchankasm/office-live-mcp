# Office Live MCP

[![CI](https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/checks.yml/badge.svg)](https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/checks.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)

[![English](https://img.shields.io/badge/English-README-0969da?style=for-the-badge)](README.md)
[![Русский](https://img.shields.io/badge/%D0%A0%D1%83%D1%81%D1%81%D0%BA%D0%B8%D0%B9-README-cf222e?style=for-the-badge)](README.ru.md)

## Your AI agent. Your open Excel and Word. Less manual work.

Turn spreadsheet rows into Word/PDF documents, compare registers and review Word edits — from the agent you already use.

> [!TIP]
> **Keep your agent. Choose your model. See changes in Office as they happen.**
> Office Live MCP connects a compatible MCP client directly to your running desktop Excel and Word.

## What you gain

| **🤖 YOUR AGENT & MODEL** | **📄 EXCEL → WORD → PDF** | **🛡️ ACCESS YOU CONTROL** |
|:---|:---|:---|
| Keep your existing instructions and workflow. Choose any model supported by your MCP client. | Generate a document per spreadsheet row from your template. Preview file names and conflicts before creating files. | Set read-only access, allowed folders and enabled tools. Limits are enforced by the server. |
| **👀 CHANGES VISIBLE IN OFFICE** | **↩️ REVIEW & UNDO** | **🔓 FREE SERVER · MIT SOURCE** |
| Work in documents already open, including unsaved edits. Results appear on screen immediately. | Preview targeted Word edits and record them as tracked revisions. Journal and undo for supported actions, with conflict checks. | No server subscription. Inspect, modify and extend the source. One installer; no Python or admin rights. |

**Works with:** Claude Code · Claude Desktop · Cursor · Codex CLI · ZCode · VS Code and other compatible MCP clients.

[**Download the Windows installer →**](https://github.com/mikhalchankasm/office-live-mcp/releases/latest) · [Installation](#install)

https://github.com/user-attachments/assets/b5911b0d-fa66-4d23-a888-ed7adeeeea07

<sub>30-second tour: real tool calls on real Excel and Word. All data is fictional.</sub>

## Try it on a real task

| Ask your agent | Get |
|---|---|
| “Generate Word documents and PDFs from this register using my template. Show the plan first.” | [A document per row, with a preview before generation](docs/examples/register-to-acts.md). Text codes such as `007` stay text. |
| “Compare these two registers by code and show what changed.” | [Differences matched by key](docs/examples/compare-registers.md). |
| “Read the comments in Word and propose edits. Apply them after I approve.” | [Targeted edits recorded as tracked revisions](docs/examples/comments-to-tracked-edits.md). |

**123 tools** for Excel, Word and document workflows: formulas, data cleanup, pivots, charts, templates, comment threads and more.

## Install

**Requirements:** Windows 10/11 (64-bit) and desktop Excel and/or Word.

1. Download the **Windows installer (the file ending in `-setup.exe`)** from the [**latest release**](https://github.com/mikhalchankasm/office-live-mcp/releases/latest).
2. Run it: **Next → Next → Finish** — on a fresh installation, detected agents are already selected. Keep them selected to connect all detected agents, and choose **full** or **read-only** access.
3. **Restart your agent** and ask: *"show me which workbooks are open in Excel"* or *"show me the open Word documents"*.

**Update:** run the new setup.exe over the existing installation, then restart the agent.
**Uninstall:** Windows **Settings → Apps → Office Live MCP → Uninstall**.

**Data and costs:** the server runs locally, without an Office add-in. Data read by your agent goes to its model provider;
the agent and model are billed under that provider's terms.

## Learn more

[![Full documentation](https://img.shields.io/badge/Full_documentation-%E2%86%92-0969da?style=for-the-badge)](docs/DETAILS.md)

[Tool catalog](docs/TOOLS.md) · [Changelog](CHANGELOG.md) · [MIT license](LICENSE) ·
[Comparison with Claude, ChatGPT and other MCP servers](docs/DETAILS.md#how-it-compares)
