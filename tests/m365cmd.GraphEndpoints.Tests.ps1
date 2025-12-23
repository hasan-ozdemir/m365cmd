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

function Ensure-GraphOrSkip {
    $ctx = Get-MgContextSafe
    if (-not $ctx) {
        Set-ItResult -Skipped -Because "Not connected to Microsoft Graph. Run /login first."
        return $false
    }
    return $true
}

function Get-GraphCallSpecs {
    param([string[]]$Paths)
    $specs = @()
    foreach ($p in $Paths) {
        if (-not (Test-Path $p)) { continue }
        $raw = Get-Content -Raw -Path $p
        $matches = [regex]::Matches($raw, '(?s)Invoke-GraphRequest(?:Auto)?\s+.*?-Uri\s+("[^"]+"|''[^'']+'')')
        foreach ($m in $matches) {
            $block = $m.Value
            $uriMatch = [regex]::Match($block, '-Uri\s+("[^"]+"|''[^'']+'')')
            if (-not $uriMatch.Success) { continue }
            $uriRaw = $uriMatch.Groups[1].Value
            $uri = $uriRaw.Trim('"','''')
            if (-not $uri.StartsWith("/")) { continue }
            $methodMatch = [regex]::Match($block, '-Method\s+("[A-Za-z]+"|''[A-Za-z]+'')')
            $method = if ($methodMatch.Success) { $methodMatch.Groups[1].Value.Trim('"','''') } else { "GET" }
            $specs += [pscustomobject]@{ Method = $method.ToUpperInvariant(); Uri = $uri }
        }
    }
    return $specs
}

function Invoke-GraphStrict {
    param([string]$Method,[string]$Uri)
    $resp = Invoke-GraphRequestAuto -Method $Method -Uri $Uri -AllowFallback
    if ($resp -eq $null) {
        throw ("Graph request failed (" + $Method + " " + $Uri + ")")
    }
    return $resp
}

Describe "Graph endpoint coverage" -Tag "integration" {
    BeforeAll {
        $repoRoot = Get-TestRoot
        $env:M365CMD_ROOT = Join-Path ([System.IO.Path]::GetTempPath()) ("m365cmd-tests-" + [guid]::NewGuid().ToString("n"))
        New-Item -ItemType Directory -Path $env:M365CMD_ROOT -Force | Out-Null
        . (Join-Path $repoRoot "lib\core.ps1")
        . (Join-Path $repoRoot "lib\handlers.ps1")
        . (Join-Path $repoRoot "lib\repl.ps1")
        $global:Config = Normalize-Config (Get-DefaultConfig)
    }

    It "All GET Graph endpoints respond" {
        if (-not (Ensure-GraphOrSkip)) { return }
        $repoRoot = Get-TestRoot
        $paths = @()
        $paths += (Get-ChildItem -Path (Join-Path $repoRoot "lib") -File -Filter "*.ps1" | Select-Object -ExpandProperty FullName)
        $specs = Get-GraphCallSpecs -Paths $paths
        $getSpecs = $specs | Where-Object { $_.Method -eq "GET" } | Sort-Object Uri -Unique
        foreach ($s in $getSpecs) {
            Invoke-GraphStrict -Method $s.Method -Uri $s.Uri | Out-Null
        }
    }
}
