Describe "Environment" {
    BeforeAll {
        $repoRoot = (Get-Location).Path
        if (-not (Test-Path (Join-Path $repoRoot "lib"))) {
            $repoRoot = Split-Path -Parent $repoRoot
        }
        if (-not (Test-Path (Join-Path $repoRoot "lib"))) {
            throw "Repo root not found. Run tests from repo root."
        }
        $lib = Join-Path $repoRoot "lib"
        $tempRoot = [System.IO.Path]::GetTempPath()
        $env:M365CMD_ROOT = Join-Path $tempRoot ("m365cmd-tests-" + [guid]::NewGuid().ToString("n"))
        New-Item -ItemType Directory -Path $env:M365CMD_ROOT -Force | Out-Null
        . (Join-Path $lib "core.ps1")
        . (Join-Path $lib "handlers.ps1")
        . (Join-Path $lib "repl.ps1")
    }
    It "runs on PowerShell 7+" {
        $PSVersionTable.PSVersion.Major | Should -BeGreaterThan 6
    }

    It "creates config in test root" {
        $cfg = Load-Config
        $cfg | Should -Not -BeNullOrEmpty
        Test-Path $Paths.Config | Should -BeTrue
    }

    It "creates required directories" {
        Ensure-Directories
        Test-Path $Paths.Data | Should -BeTrue
        Test-Path $Paths.Logs | Should -BeTrue
        Test-Path $Paths.Modules | Should -BeTrue
        Test-Path $Paths.Tools | Should -BeTrue
    }
}

Describe "Module availability" -Tag "modules" {
    It "Microsoft.Graph is available" {
        (Test-ModuleAvailable "Microsoft.Graph") | Should -BeTrue
    }
}

Describe "Integration" -Tag "integration" {
    It "can call /me when connected" {
        $ctx = Get-MgContextSafe
        if (-not $ctx) {
            Set-ItResult -Skipped -Because "Not connected to Microsoft Graph. Run /login first."
            return
        }
        $me = Invoke-GraphRequest -Method "GET" -Uri "/me"
        $me | Should -Not -BeNullOrEmpty
    }
}
