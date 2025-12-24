# Repository Guidelines

This repository hosts **m365cmd**, a portable PowerShell Core REPL for Microsoft 365 administration. Keep changes modular, testable, and aligned with the existing core/handler architecture.

## Project Structure & Module Organization

- Entry points: `m365cmd.cmd` (Windows), `m365cmd.sh` (macOS/Linux), `m365cmd-main.ps1` (cross-platform).
- `lib/`: all runtime code.
  - `core-*.ps1`: shared utilities (config, parsing, Graph, auth, metadata, files).
  - `handlers-*.ps1`: service-specific command handlers.
  - `core-loader.ps1`, `handlers-loader.ps1`: load modules in manifest order.
  - `core.manifest.json`, `handlers.manifest.json`: deterministic load order.
  - `apps.catalog.json`: app portal/CLI shortcuts.
- `tests/`: Pester suites + `coverage.manifest.json`.
- `modules/`, `data/`, `logs/`: generated at runtime; do not commit secrets.

## Build, Test, and Development Commands

- Run REPL: `m365cmd.cmd` or `./m365cmd.sh`.
- Basic tests: `Powershell Core -NoProfile -File .\tests\Run-Tests.ps1`.
- Module checks: `Powershell Core -NoProfile -File .\tests\Run-Tests.ps1 -Modules`.
- Integration tests: `Powershell Core -NoProfile -File .\tests\Run-Tests.ps1 -Integration`.
- Write tests (creates/deletes data): `Powershell Core -NoProfile -File .\tests\Run-Tests.ps1 -Write`.

## Coding Style & Naming Conventions

- PowerShell Core 7+ syntax; 4-space indentation; ASCII by default unless a file already uses Unicode.
- Name new utilities as `core-XX-<topic>.ps1` and handlers as `handlers-XX-<area>.ps1`.
- Keep handlers single-responsibility and update `/help` for new commands.
- After adding/removing modules, run `manifest sync` and update tests.

## Testing Guidelines

- Framework: Pester. Test files are `*.Tests.ps1` under `tests/`.
- Add tests for every new core/handler function.
- Integration tests must be gated behind `-Integration`; write tests behind `-Write`.
- Update `tests/coverage.manifest.json` when new scripts are added.

## Commit & Pull Request Guidelines

- No established commit history yet; use clear, imperative messages (e.g., `Add Graph metadata cache`).
- PRs should include: summary, rationale, tests run, and any breaking changes.
- Include screenshots only for CLI output changes or new help flows.

## Security & Configuration Tips

- Secrets live in `m365cmd.config.json`; never commit real credentials.
- Prefer delegated auth; use app-only creds only when required.
- Add confirmation prompts for destructive actions and document any side effects.

## Agent Workflow Rule

- When a user request results in file updates, finish all work first, then create a **single** meaningful English `git commit` as the final step before replying with a summary.
