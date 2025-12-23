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

function Invoke-UrlProbe {
    param([string]$Url)
    try {
        $resp = Invoke-WebRequest -Uri $Url -Method Head -SkipHttpErrorCheck -TimeoutSec 15
        return [int]$resp.StatusCode
    } catch {
        try {
            $resp = Invoke-WebRequest -Uri $Url -Method Get -SkipHttpErrorCheck -TimeoutSec 15
            return [int]$resp.StatusCode
        } catch {
            return 0
        }
    }
}

function Get-UrlListFromFiles {
    param([string[]]$Paths)
    $urls = @()
    foreach ($p in $Paths) {
        if (-not (Test-Path $p)) { continue }
        $raw = Get-Content -Raw -Path $p
        foreach ($m in [regex]::Matches($raw, 'https?://[^\s''")]+')) {
            $u = $m.Value.TrimEnd(".",",",";")
            if ($u) { $urls += $u }
        }
    }
    return ($urls | Sort-Object -Unique)
}

Describe "API endpoints" -Tag "integration" {
    BeforeAll {
        $repoRoot = Get-TestRoot
        $env:M365CMD_ROOT = Join-Path ([System.IO.Path]::GetTempPath()) ("m365cmd-tests-" + [guid]::NewGuid().ToString("n"))
        New-Item -ItemType Directory -Path $env:M365CMD_ROOT -Force | Out-Null
        . (Join-Path $repoRoot "lib\core.ps1")
        . (Join-Path $repoRoot "lib\handlers.ps1")
        . (Join-Path $repoRoot "lib\repl.ps1")
        $global:Config = Normalize-Config (Get-DefaultConfig)
    }

    It "Graph metadata v1 is reachable" {
        (Invoke-UrlProbe "https://graph.microsoft.com/v1.0/`$metadata") | Should -BeGreaterThan 0
    }

    It "Graph metadata beta is reachable" {
        (Invoke-UrlProbe "https://graph.microsoft.com/beta/`$metadata") | Should -BeGreaterThan 0
    }

    It "Forms base URL is reachable" {
        $url = $global:Config.forms.baseUrl
        if (-not $url) { $url = "https://forms.office.com" }
        (Invoke-UrlProbe $url) | Should -BeGreaterThan 0
    }

    It "O365 Management API base is reachable" {
        $url = $global:Config.o365.manageApiBase
        if (-not $url) { $url = "https://manage.office.com" }
        (Invoke-UrlProbe $url) | Should -BeGreaterThan 0
    }

    It "Defender base is reachable" {
        $url = $global:Config.defender.baseUrl
        if (-not $url) { $url = "https://api.security.microsoft.com" }
        (Invoke-UrlProbe $url) | Should -BeGreaterThan 0
    }

    It "Power Platform base is reachable" {
        $url = $global:Config.pp.baseUrl
        if (-not $url) { $url = "https://api.powerplatform.com" }
        (Invoke-UrlProbe $url) | Should -BeGreaterThan 0
    }

    It "Portal URLs from app catalog respond" {
        $catalog = Join-Path (Get-TestRoot) "lib\apps.catalog.json"
        if (-not (Test-Path $catalog)) { Set-ItResult -Skipped -Because "App catalog missing."; return }
        $apps = Get-Content -Raw -Path $catalog | ConvertFrom-Json
        foreach ($app in @($apps)) {
            if (-not $app.portal) { continue }
            $code = Invoke-UrlProbe $app.portal
            $code | Should -BeGreaterThan 0
        }
    }

    It "All hard-coded URLs in scripts respond" {
        $repoRoot = Get-TestRoot
        $paths = @()
        $paths += (Get-ChildItem -Path $repoRoot -File -Filter "*.ps1" | Select-Object -ExpandProperty FullName)
        $paths += (Get-ChildItem -Path (Join-Path $repoRoot "lib") -File -Filter "*.ps1" | Select-Object -ExpandProperty FullName)
        $urls = Get-UrlListFromFiles -Paths $paths
        foreach ($u in $urls) {
            $code = Invoke-UrlProbe $u
            $code | Should -BeGreaterThan 0
        }
    }
}
