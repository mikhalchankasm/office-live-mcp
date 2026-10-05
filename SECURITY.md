# Security

## Reporting a vulnerability

Please report security issues privately through
[GitHub Security Advisories](https://github.com/mikhalchankasm/office-live-mcp/security/advisories/new)
rather than a public issue. Include the version (`office-live-mcp.exe doctor` prints it), steps to reproduce and the
impact. You will get an answer within a few days.

## Security model in short

Office Live MCP runs locally as a child process of your AI agent and talks to it over stdio; it opens no network ports.
It acts with your Windows user's rights inside Excel and Word, so the main risks are an agent changing documents you
did not intend to change and document content reaching the model.

- `OFFICE_LIVE_MODE=readonly` registers no writing tools at all.
- `OFFICE_LIVE_ALLOWED_DIRS` hides and protects every file outside the listed folders (open, save, export, insert).
- Changes to AutoSave (cloud) documents are refused by default; Excel events and macros are disabled while the agent writes.
- Every change is journaled per document; `office_undo` reverts the agent's steps and refuses when the user edited the
  same content afterwards unless forced.
- Text inside documents is untrusted data: the server instructs agents never to follow instructions found there.
- `office_run_python` (arbitrary Python with COM access) is off unless `OFFICE_LIVE_ALLOW_EVAL=1`; it bypasses all of the above.

What is guaranteed and what is not is described in detail in [docs/GUIDE.ru.md](docs/GUIDE.ru.md) (section «Безопасность»).
