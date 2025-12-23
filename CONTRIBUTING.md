# Contributing Guide

Thanks for your interest in contributing to **m365cmd**. This project is a PowerShell Core REPL for Microsoft 365 administration, and contributions that improve reliability, usability, and coverage are welcome.

## How to Contribute

1. **Open an issue first** (for new features or changes in behavior).
2. **Fork the repository** and create a feature branch.
3. **Keep changes small** and focused on a single topic.
4. **Add tests** for any new core or handler functionality.
5. **Update docs** when you change commands, config, or behaviors.

## Development Setup

- PowerShell Core 7+ is required.
- Optional: Node.js 20+ LTS for CLI for Microsoft 365 integration.

Run the REPL:

```bash
./m365cmd.sh
```

Or on Windows:

```powershell
m365cmd.cmd
```

## Testing

Use the Pester test runner:

```powershell
pwsh -NoProfile -File .\tests\Run-Tests.ps1
```

Other test modes:

```powershell
pwsh -NoProfile -File .\tests\Run-Tests.ps1 -Modules
pwsh -NoProfile -File .\tests\Run-Tests.ps1 -Integration
pwsh -NoProfile -File .\tests\Run-Tests.ps1 -Write
```

## Coding Conventions

- PowerShell Core 7+ syntax; 4-space indentation.
- ASCII by default unless a file already uses Unicode.
- New core utilities: `lib/core-XX-<topic>.ps1`.
- New handlers: `lib/handlers-XX-<area>.ps1`.
- Update `lib/core.manifest.json` and `lib/handlers.manifest.json` after adding modules (`manifest sync`).

## Documentation Expectations

- Update `README.md` when adding commands, services, or workflows.
- Update `/help` text in handlers when adding new commands.

## Commit Messages

Use clear, imperative messages:

- `feat(handlers): add Teams message purge command`
- `fix(core): handle null Graph context`
- `docs: explain auth scopes`

## Pull Request Checklist

- [ ] Summary and rationale included
- [ ] Tests added/updated
- [ ] Tests executed (list commands)
- [ ] Docs updated (README/help)
- [ ] No secrets in config or logs

## Security

- Do not commit real credentials or tokens.
- Keep destructive operations behind confirmation prompts.
- Prefer delegated auth unless app-only is required.

## Questions / Discussion

Open a GitHub issue for questions or proposals.
