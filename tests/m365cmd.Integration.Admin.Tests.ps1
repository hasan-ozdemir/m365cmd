function Get-TestRoot {
    $repoRoot = (Get-Location).Path
    if (-not (Test-Path (Join-Path $repoRoot "lib"))) {
        $repoRoot = Split-Path -Parent $repoRoot
    }
    if (-not (Test-Path (Join-Path $repoRoot "lib"))) {
        throw "Repo root not found. Run tests from repo root."
    }
    return $repoRoot
}

Describe "Admin module integration" -Tag "integration" {
    BeforeAll {
        $repoRoot = Get-TestRoot
        $env:M365CMD_ROOT = Join-Path ([System.IO.Path]::GetTempPath()) ("m365cmd-tests-" + [guid]::NewGuid().ToString("n"))
        New-Item -ItemType Directory -Path $env:M365CMD_ROOT -Force | Out-Null
        . (Join-Path $repoRoot "lib\core.ps1")
        . (Join-Path $repoRoot "lib\handlers.ps1")
        . (Join-Path $repoRoot "lib\repl.ps1")
        $global:Config = Normalize-Config (Get-DefaultConfig)
    }

    It "Exchange Online" {
        if (-not (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "ExchangeOnlineManagement module not loaded or unavailable."
            return
        }
        try {
            $info = Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1
            if (-not $info) {
                Set-ItResult -Skipped -Because "Not connected to Exchange Online."
                return
            }
            $info | Should -Not -BeNullOrEmpty
        } catch {
            throw $_
        }
    }

    It "SharePoint Online" {
        if (-not (Get-Command Get-SPOTenant -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "SPO module not loaded or unavailable."
            return
        }
        try {
            $tenant = Get-SPOTenant -ErrorAction Stop
            $tenant | Should -Not -BeNullOrEmpty
        } catch {
            Set-ItResult -Skipped -Because "Not connected to SharePoint Online."
        }
    }

    It "Teams" {
        if (-not (Get-Command Get-CsTenant -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "MicrosoftTeams module not loaded or unavailable."
            return
        }
        try {
            $tenant = Get-CsTenant -ErrorAction Stop
            $tenant | Should -Not -BeNullOrEmpty
        } catch {
            Set-ItResult -Skipped -Because "Not connected to Microsoft Teams."
        }
    }

    It "Compliance (IPPSSession)" {
        if (-not (Get-Command Get-ComplianceSearch -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because "Compliance PowerShell not connected."
            return
        }
        try {
            $searches = Get-ComplianceSearch -ErrorAction Stop
            $searches | Should -Not -BeNullOrEmpty
        } catch {
            Set-ItResult -Skipped -Because "Compliance PowerShell not connected or no permissions."
        }
    }
}
