<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="docs/assets/hero-dark.svg">
    <img alt="Office Live MCP — your AI agent in the Excel and Word files you already have open" src="docs/assets/hero-light.svg" width="100%">
  </picture>
</p>

<p align="center">
  <a href="https://github.com/mikhalchankasm/office-live-mcp/releases/latest"><img alt="Latest release" src="https://img.shields.io/github/v/release/mikhalchankasm/office-live-mcp?label=release&color=217346"></a>
  <a href="https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/checks.yml"><img alt="CI" src="https://github.com/mikhalchankasm/office-live-mcp/actions/workflows/checks.yml/badge.svg"></a>
  <a href="LICENSE"><img alt="License: MIT" src="https://img.shields.io/badge/license-MIT-2b579a"></a>
  <img alt="Windows 10/11" src="https://img.shields.io/badge/Windows-10%20%7C%2011-555">
</p>

<p align="center">
  <a href="https://github.com/mikhalchankasm/office-live-mcp/releases/latest"><img alt="Download the Windows installer" src="https://img.shields.io/badge/Download_for_Windows-setup.exe-217346?style=for-the-badge"></a>
  &nbsp;
  <a href="docs/DETAILS.md"><img alt="Documentation" src="https://img.shields.io/badge/Documentation-%E2%86%92-2b579a?style=for-the-badge"></a>
</p>

<p align="center"><b>English</b> · <a href="README.ru.md">Русский</a></p>

<h3 align="center">Your AI agent. Your open Excel and Word. Less manual work.</h3>

<p align="center">
Turn spreadsheet rows into Word/PDF documents, compare registers and review Word edits —<br>
from the agent you already use, in the files you already have open.
</p>

https://github.com/user-attachments/assets/b5911b0d-fa66-4d23-a888-ed7adeeeea07

<p align="center"><sub>30-second tour: real tool calls on real Excel and Word. All data is fictional.</sub></p>

## Why Office Live MCP

<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="docs/assets/features-dark.svg">
    <img alt="Your agent and model; Excel to Word to PDF; access you control; changes you can see; review and undo; free and open source" src="docs/assets/features-light.svg" width="100%">
  </picture>
</p>

<p align="center">
  <b>Works with</b>&nbsp; Claude Code · Claude Desktop · Cursor · Codex CLI · ZCode · VS Code<br>
  <sub>and other MCP clients</sub>
</p>

<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="docs/assets/how-dark.svg">
    <img alt="How it works: your AI agent → Office Live MCP, running locally with read-only mode, allowed folders, journal and undo → the Excel and Word files you already have open" src="docs/assets/how-light.svg" width="100%">
  </picture>
</p>

## Try it on a real task

| Ask your agent | You get |
|---|---|
| *“Generate Word documents and PDFs from this register using my template. Show the plan first.”* | [A document per row, previewed before generation](docs/examples/register-to-acts.md). Codes such as `007` stay text. |
| *“Compare these two registers by code and show what changed.”* | [Differences matched by key](docs/examples/compare-registers.md). |
| *“Read the comments in Word and propose edits. Apply them after I approve.”* | [Targeted edits recorded as tracked revisions](docs/examples/comments-to-tracked-edits.md). |

**123 tools** for Excel, Word and document workflows — formulas, data cleanup, goal seek, pivots, charts, templates,
comment threads and more. [Full catalog →](docs/TOOLS.md)

## Install

> **Requirements:** Windows 10/11 (64-bit) and desktop Excel and/or Word.

1. **Download** the installer (`…-setup.exe`) from the [latest release](https://github.com/mikhalchankasm/office-live-mcp/releases/latest).
2. **Run it:** Next → Next → Finish. On a fresh install, detected agents are preselected; choose **full** or **read-only** access.
3. **Restart your agent** and ask: *“show me which workbooks are open in Excel”*.

<sub>**Update:** run the new setup.exe over the old one. **Uninstall:** Settings → Apps → Office Live MCP.
**Data and costs:** the server runs locally, without an Office add-in; what your agent reads goes to its model provider,
and the agent and model are billed under that provider's terms.</sub>

## Learn more

[Full documentation](docs/DETAILS.md) · [Tool catalog](docs/TOOLS.md) · [Changelog](CHANGELOG.md) ·
[Comparison with Claude, ChatGPT and other MCP servers](docs/DETAILS.md#how-it-compares) · [MIT license](LICENSE)
