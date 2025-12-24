# m365cmd tests

## How to run

- Basic sanity tests:
  - `Powershell Core -NoProfile -File .\tests\Run-Tests.ps1`

- Include module availability checks:
  - `Powershell Core -NoProfile -File .\tests\Run-Tests.ps1 -Modules`

- Include integration tests (requires `/login` in m365cmd):
  - `Powershell Core -NoProfile -File .\tests\Run-Tests.ps1 -Integration`

- Include integration + write tests (creates and deletes test data):
  - `Powershell Core -NoProfile -File .\tests\Run-Tests.ps1 -Write`

- Install Pester automatically if missing:
  - `Powershell Core -NoProfile -File .\tests\Run-Tests.ps1 -InstallPester`

## Notes
- Module and integration tests are tagged and skipped by default.
- Tests isolate config/data under a temporary test root.
- Integration suite includes app smoke tests and API endpoint probes.
- Graph endpoint coverage test calls all GET endpoints found in code and may require broad permissions.
- Write suite will create and delete test objects (groups, apps, folders, todo lists, events).
