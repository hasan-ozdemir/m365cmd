# m365cmd

A portable, modular PowerShell Core REPL for Microsoft 365 administration and workload management.

**Community contribution:** This project is a community contribution by **Senior lead software engineer Hasan Ozdemir**. Sponsorships and donations are welcome; your support helps sustain development and gives extra motivation to keep expanding the project.

---

## Table of contents

- [Why m365cmd](#why-m365cmd)
- [What it does](#what-it-does)
- [Design goals](#design-goals)
- [Architecture overview](#architecture-overview)
- [Folder structure](#folder-structure)
- [Core components](#core-components)
- [Handler components](#handler-components)
- [Runtime flow](#runtime-flow)
- [Supported areas](#supported-areas)
- [Requirements](#requirements)
- [Supported platforms](#supported-platforms)
- [Quick start](#quick-start)
- [Installation](#installation)
- [First login](#first-login)
- [Authentication models](#authentication-models)
- [Command model](#command-model)
- [Commands by area](#commands-by-area)
- [Apps catalog](#apps-catalog)
- [Configuration](#configuration)
- [Modules and dependencies](#modules-and-dependencies)
- [Microsoft Graph v1 vs beta](#microsoft-graph-v1-vs-beta)
- [Data, logs, and portability](#data-logs-and-portability)
- [Testing](#testing)
- [Development](#development)
- [Contributing](#contributing)
- [Sponsor / donate](#sponsor--donate)
- [References](#references)
- [License](#license)
- [Disclaimer](#disclaimer)

---

## Why m365cmd

If you manage Microsoft 365 daily, you likely jump between multiple portals, modules, and scripts. **m365cmd** centralizes common admin and workload tasks in a single, consistent CLI with discoverable commands, reusable presets, and a portable folder structure.

---

## What it does

m365cmd is a **PowerShell Core 7+ REPL** that unifies:

- Microsoft Graph (v1.0 + beta) operations
- Admin center workflows (users, groups, roles, domains, licensing, etc.)
- Exchange Online, Teams, and SharePoint Online connectivity
- Files, OneDrive, SharePoint, Teams messages, mail, calendar, tasks, OneNote, Planner
- Power Platform admin helpers
- Viva, Purview, Defender, and compliance areas (as far as public APIs allow)

---

## Design goals

- **Portable:** everything stays inside this folder (config, data, logs, modules)
- **Consistent UX:** global `/` commands, local commands, named args
- **Graph-first:** default to v1, fallback to beta when needed
- **Modular:** core utilities and handlers are separated and easy to extend
- **MFA-friendly:** interactive or device-code login
- **Testable:** Pester-backed tests for core, handlers, and integrations

---

## Architecture overview

m365cmd is organized as a small runtime (REPL + loaders) and a set of modular building blocks:

- **Entry points**: `m365cmd.cmd` (Windows), `m365cmd.sh` (macOS/Linux), and `m365cmd-main.ps1` (cross-platform)
- **Loaders**: `core-loader.ps1` and `handlers-loader.ps1` load modules based on manifest order
- **Core utilities**: shared primitives for config, parsing, Graph, modules, metadata, files, auth
- **Handlers**: feature-focused command handlers grouped by service area
- **REPL**: `repl.ps1` drives input, parsing, dispatch, and output
- **Manifests**: `core.manifest.json`, `handlers.manifest.json`, and `tests/coverage.manifest.json`
- **Catalogs and config**: `apps.catalog.json` and `m365cmd.config.json`

---

## Folder structure

```
./
  m365cmd.cmd              # Windows launcher
  m365cmd.sh               # macOS/Linux launcher
  m365cmd-main.ps1          # Cross-platform entrypoint
  m365cmd.config.json       # Runtime config (generated)
  README.md                 # Documentation
  LICENSE                   # License

  lib/
    core.ps1                # Core bootstrap (loads core modules)
    handlers.ps1            # Handler bootstrap (loads handler modules)
    repl.ps1                # REPL loop
    core-loader.ps1         # Core loader (manifest + order)
    handlers-loader.ps1     # Handler loader (manifest + order)
    core-*.ps1              # Core utilities (see below)
    handlers-*.ps1          # Command handlers (see below)
    core.manifest.json      # Core module load order
    handlers.manifest.json  # Handler module load order
    apps.catalog.json       # App portal/command catalog

  tests/
    Run-Tests.ps1            # Test runner (Pester)
    coverage.manifest.json   # Coverage map for all scripts
    *.Tests.ps1              # Unit/integration test suites

  data/                      # Runtime data cache (created at runtime)
  logs/                      # Logs (created at runtime)
  modules/                   # PowerShell module cache (auto-installed)
```

---

## Core components

Each core file provides shared utilities for the handlers and the REPL:

| File | Purpose |
| --- | --- |
| `lib/core-00-base.ps1` | Base helpers, logging, paths, console formatting, common utilities. |
| `lib/core-10-config.ps1` | Config defaults, load/save, normalization, config getters/setters. |
| `lib/core-20-modules.ps1` | Module discovery, install/update, import, auto-install logic. |
| `lib/core-30-parse.ps1` | Command parsing, tokenization, named-arg parsing. |
| `lib/core-31-query.ps1` | Query helpers for list/filter/paging patterns. |
| `lib/core-32-format.ps1` | Output formatting helpers (tables, JSON, normalized objects). |
| `lib/core-33-alias.ps1` | Alias and preset expansion logic. |
| `lib/core-34-mail.ps1` | Mail message helpers and mail-specific utilities. |
| `lib/core-40-graph-context.ps1` | Graph context discovery, login state helpers. |
| `lib/core-41-graph-requests.ps1` | Graph request wrappers with v1/beta fallback. |
| `lib/core-42-graph-org.ps1` | Organization/tenant helpers (IDs, metadata). |
| `lib/core-50-resolve-userdrive.ps1` | Resolve user IDs and OneDrive/drive IDs. |
| `lib/core-51-resolve-appgroup.ps1` | Resolve app, group, and related IDs. |
| `lib/core-52-resolve-roleperm.ps1` | Resolve roles/permissions and lookups. |
| `lib/core-60-auth.ps1` | Auth helpers (device code, app-only tokens). |
| `lib/core-70-metadata.ps1` | Graph metadata cache and refresh logic. |
| `lib/core-80-files.ps1` | Filesystem and OneDrive path helpers. |
| `lib/core-90-pp.ps1` | Power Platform API helpers. |
| `lib/core-loader.ps1` | Loads core modules in a deterministic order. |
| `lib/core.ps1` | Core bootstrap; wires globals and shared state. |

---

## Handler components

Each handler file contains commands for a specific service area:

| File | Area |
| --- | --- |
| `lib/handlers-00-help.ps1` | Help system and command catalog. |
| `lib/handlers-01-common.ps1` | Shared handler utilities and common commands. |
| `lib/handlers-02-admin.ps1` | Admin tasks (users, groups, roles, licensing). |
| `lib/handlers-03-adminportal.ps1` | Admin portal shortcuts and helpers. |
| `lib/handlers-04-exo.ps1` | Exchange Online connectivity and commands. |
| `lib/handlers-05-teams.ps1` | Teams connectivity and commands. |
| `lib/handlers-06-spo.ps1` | SharePoint Online connectivity and SPO admin tasks. |
| `lib/handlers-07-files.ps1` | Files/OneDrive/SharePoint file operations. |
| `lib/handlers-08-sharepoint.ps1` | SharePoint sites, lists, pages, content types. |
| `lib/handlers-09-graph.ps1` | Graph generic commands and utilities. |
| `lib/handlers-10-webhook.ps1` | Webhook management and Graph subscriptions. |
| `lib/handlers-11-forms.ps1` | Microsoft Forms helpers and shortcuts. |
| `lib/handlers-12-stream.ps1` | Microsoft Stream helpers and shortcuts. |
| `lib/handlers-13-copilot.ps1` | Copilot Graph beta commands. |
| `lib/handlers-14-auth.ps1` | Login/logout and auth-related commands. |
| `lib/handlers-15-risk.ps1` | Identity risk and protection commands. |
| `lib/handlers-16-bookings.ps1` | Microsoft Bookings commands. |
| `lib/handlers-17-orgx.ps1` | Org Explorer helpers. |
| `lib/handlers-18-whiteboard.ps1` | Whiteboard helpers and links. |
| `lib/handlers-19-apps-common.ps1` | App catalog utilities and shared app helpers. |
| `lib/handlers-20-apps.ps1` | Apps catalog commands (open, list, cli). |
| `lib/handlers-21-clipchamp.ps1` | Clipchamp helpers and shortcuts. |
| `lib/handlers-22-connections.ps1` | Connections/Viva Connections helpers. |
| `lib/handlers-23-engage.ps1` | Engage/Yammer helpers. |
| `lib/handlers-24-insights.ps1` | Viva Insights helpers. |
| `lib/handlers-25-kaizala.ps1` | Kaizala helpers. |
| `lib/handlers-26-learning.ps1` | Viva Learning helpers. |
| `lib/handlers-27-lists.ps1` | Microsoft Lists helpers. |
| `lib/handlers-28-loop.ps1` | Loop helpers. |
| `lib/handlers-29-outlook.ps1` | Outlook mail/calendar/contact helpers. |
| `lib/handlers-30-powerapps.ps1` | Power Apps commands. |
| `lib/handlers-31-powerautomate.ps1` | Power Automate commands. |
| `lib/handlers-32-powerpages.ps1` | Power Pages commands. |
| `lib/handlers-33-powerpoint.ps1` | PowerPoint helpers. |
| `lib/handlers-34-sway.ps1` | Sway helpers. |
| `lib/handlers-35-visio.ps1` | Visio helpers. |
| `lib/handlers-36-word.ps1` | Word helpers. |
| `lib/handlers-37-tools.ps1` | Utility commands (diagnostics, environment). |
| `lib/handlers-38-addins.ps1` | Add-in deployment and management helpers. |
| `lib/handlers-39-alias.ps1` | Alias management commands. |
| `lib/handlers-40-manifest.ps1` | Manifest tooling and synchronization. |
| `lib/handlers-41-common.ps1` | Shared helpers for advanced/security/compliance commands. |
| `lib/handlers-42-accessreview.ps1` | Access reviews. |
| `lib/handlers-43-audit.ps1` | Audit log queries. |
| `lib/handlers-44-auditfeed.ps1` | Office 365 Management Activity API feed. |
| `lib/handlers-45-billing.ps1` | Billing and subscription helpers. |
| `lib/handlers-46-ca.ps1` | Conditional access helpers. |
| `lib/handlers-47-compliance.ps1` | Compliance center helpers. |
| `lib/handlers-48-defender.ps1` | Defender for Microsoft 365 helpers. |
| `lib/handlers-49-device.ps1` | Device inventory and management helpers. |
| `lib/handlers-50-health.ps1` | Service health and messages. |
| `lib/handlers-51-intune.ps1` | Intune helpers. |
| `lib/handlers-52-label.ps1` | Sensitivity label helpers. |
| `lib/handlers-53-message.ps1` | Service message center helpers. |
| `lib/handlers-54-pp.ps1` | Power Platform base commands. |
| `lib/handlers-55-purview.ps1` | Purview helpers. |
| `lib/handlers-56-report.ps1` | Reports and usage analytics. |
| `lib/handlers-57-security.ps1` | Security center helpers. |
| `lib/handlers-58-viva.ps1` | Viva suite helpers. |
| `lib/handlers-99-dispatch.ps1` | Command dispatch router. |
| `lib/handlers-loader.ps1` | Loads handler modules in a deterministic order. |
| `lib/handlers.ps1` | Handler bootstrap; wires dispatch and help data. |
| `lib/repl.ps1` | REPL loop (input, parse, dispatch, output). |

---

## Runtime flow

1. `m365cmd.cmd` (Windows), `m365cmd.sh` (macOS/Linux), or `m365cmd-main.ps1` starts the app
2. `core.ps1` loads core modules in manifest order
3. `handlers.ps1` loads handlers in manifest order
4. `repl.ps1` starts a loop: read -> parse -> dispatch -> output
5. Handlers call core utilities (Graph, auth, parsing, formatting)
6. Results are printed or saved to data/logs as needed

---

## Supported areas

Use `/help` for the live command list. Highlights include:

- **Admin:** users, groups, roles, domains, licensing, directory settings
- **Mail/Calendar:** mail folders/messages, events, contacts
- **Files:** OneDrive + SharePoint files, sharing permissions, links
- **Teams:** chats, messages, tabs, app installs
- **SharePoint:** sites, lists, pages, columns, content types, permissions
- **Security/Compliance:** alerts, incidents, audit, Purview, Defender
- **Power Platform:** environments, flows, Power Apps/Power Pages
- **Copilot:** chat/search/retrieve (Graph beta)

---

## Requirements

- PowerShell Core 7+
- Internet access for Graph and module installs
- Optional: Node.js 18+ (LTS recommended) for SPFx project packaging commands

---

## Supported platforms

- Windows (primary target; includes `m365cmd.cmd`)
- macOS (PowerShell Core)
- Linux (PowerShell Core)

---

## Quick start

### Windows

```powershell
m365cmd.cmd
```

### macOS / Linux

```bash
chmod +x ./m365cmd.sh
```

```bash
./m365cmd.sh
```

If needed, you can also run directly with PowerShell Core:

```bash
Powershell Core -NoProfile -File ./m365cmd-main.ps1
```

Inside the REPL:

```text
/help
/login
/status
```

---

## Installation

This project is **portable**. The simplest install is:

1. Download/clone the repository
2. Keep the entire folder intact
3. Run `m365cmd.cmd` (Windows) or `./m365cmd.sh` (macOS/Linux)

On macOS/Linux, ensure the launcher is executable: `chmod +x m365cmd.sh`.

No global installation is required.

---

## First login

```text
/login
```

If your environment blocks interactive auth:

```text
/login device
```

---

## Authentication models

m365cmd supports:

1. **Delegated user login** (interactive or device code)
2. **App-only tokens** (client credentials) for specific APIs

App-only credentials live in config:

```text
/config set auth.app.clientId <id>
/config set auth.app.clientSecret <secret>
/config set auth.app.tenantId <tenantId>
```

---

## Command model

### Global commands (start with `/`)

- `/help` - full help
- `/status` - connection + config status
- `/login` - connect to Microsoft Graph
- `/logout` - disconnect from Microsoft Graph
- `/tenant` - update defaults (prefix/domain/tenantId)
- `/config` - show/update config values

### Local commands

Local commands do not use `/`. Examples:

```text
user list
user get info@contoso.com
license list
role list
file list --path /Documents
spo connect --prefix contoso
teams connect
```

### Aliases and presets

You can define custom aliases and multi-step presets:

```text
/config set aliases.local.u "user {args}"
/config set presets.user-hard-reset "user disable {args}; user session revoke {args}; user mfa reset {args}"
```

---

## Commands by area

Use `/help` or `/help <topic>` for full command lists.

- **Admin:** users, groups, roles, domains, licensing, directory settings
- **Mail/Calendar:** mail folders/messages, events, contacts
- **Files:** OneDrive + SharePoint files, sharing permissions, links
- **Teams:** chats, messages, tabs, app installs
- **SharePoint:** sites, lists, pages, columns, content types, permissions
- **Security/Compliance:** alerts, incidents, audit, Purview, Defender
- **Power Platform:** environments, flows, Power Apps/Power Pages
- **Copilot:** chat/search/retrieve (Graph beta)

---

## Apps catalog

`apps` is a lightweight catalog for one-liners that open app portals or route to command mappings.

```text
apps list
apps get copilot
apps open sharepoint
apps map set powerapps --cmd "powerapps app list"
apps run powerapps
apps cli powerapps --cmd "powerapps app list"
```

Catalog file:

```
lib/apps.catalog.json
```

Mappings are stored in:

```
data/appmap.json
```

---

## Configuration

Config is stored here:

```
./m365cmd.config.json
```

Common settings:

- `tenant.defaultPrefix`
- `tenant.defaultDomain`
- `tenant.tenantId`
- `admin.defaultUpn`
- `auth.scopes`
- `auth.loginMode`
- `auth.app.clientId` / `clientSecret` / `tenantId`
- `graph.defaultApi` (v1 | beta)
- `graph.fallbackToBeta` (true/false)
- `graph.autoSyncMetadata` (true/false)
- `graph.metadataRefreshHours`
- `modules.autoInstall`

Example:

```text
/config set tenant.defaultDomain contoso.onmicrosoft.com
/config set admin.defaultUpn admin@contoso.com
```

---

## Modules and dependencies

Modules are auto-installed on demand into `./modules/`:

- Microsoft.Graph
- Microsoft.Graph.Beta
- ExchangeOnlineManagement
- MicrosoftTeams
- Microsoft.Online.SharePoint.PowerShell
- MSAL.PS
- O365CentralizedAddInDeployment

You can also manually install or update:

```text
module list
module install Microsoft.Graph
module update Microsoft.Graph
```

---

## Microsoft Graph v1 vs beta

- v1.0 is used by default for stability.
- Beta is supported for features that do not exist in v1.0.
- Beta endpoints can change without notice and are not recommended for production-only flows.

If you need beta functionality, install the beta module:

```powershell
Install-Module Microsoft.Graph.Beta
```

---

## Data, logs, and portability

This project is designed to be **portable**:

- Entire folder can be copied to another machine
- No global install required
- All artifacts remain in `./data`, `./logs`, and `./modules`

---

## Testing

All PowerShell scripts are covered by tests and a coverage manifest.
Pester is used as the test framework.

### Basic tests

```powershell
Powershell Core -NoProfile -File .\tests\Run-Tests.ps1
```

### Include module checks

```powershell
Powershell Core -NoProfile -File .\tests\Run-Tests.ps1 -Modules
```

### Integration tests (requires /login)

```powershell
Powershell Core -NoProfile -File .\tests\Run-Tests.ps1 -Integration
```

### Write tests (creates + deletes data)

```powershell
Powershell Core -NoProfile -File .\tests\Run-Tests.ps1 -Write
```

Write tests create temporary objects (groups, apps, folders, todo lists, events) and delete them automatically.

---

## Development

### Structure

```
lib/
  core-*.ps1       # core utilities
  handlers-*.ps1   # feature handlers
  repl.ps1         # REPL loop
  core-loader.ps1
  handlers-loader.ps1

m365cmd-main.ps1   # entrypoint
m365cmd.cmd        # Windows launcher
m365cmd.sh         # macOS/Linux launcher
```

### Adding a command

1. Add handler code to the correct `handlers-*.ps1`
2. Update manifests using:

```text
manifest sync
```

3. Add /help entries
4. Add tests

### Adding a new core utility

1. Create a new `core-XX-*.ps1` file
2. Add it to `core.manifest.json` (or run `manifest sync`)
3. Add tests to `tests/` for new helpers

### Updating the apps catalog

- Edit `lib/apps.catalog.json`
- Use `apps list` to verify entries

---

## Contributing

- Open issues for feature requests
- Fork and submit PRs
- Keep functions testable and add Pester tests
- Maintain the naming pattern for core and handler files

---

## Sponsor / donate

This project is a community contribution by **Senior lead software engineer Hasan Ozdemir**. If this tool helps you, consider supporting its growth. Sponsorships and donations make continued development sustainable and provide extra motivation for new features, integrations, and documentation improvements.

---

## References

- Microsoft Graph PowerShell SDK installation: https://learn.microsoft.com/en-us/graph/sdks/sdk-installation
- Microsoft Graph beta guidance: https://learn.microsoft.com/en-us/graph/sdks/use-beta
- Pester documentation: https://pester.dev/docs/quick-start

---

## License

See `LICENSE`.

---

## Disclaimer

This project is not affiliated with Microsoft. Use at your own risk. Ensure you comply with your organization security and compliance policies.
