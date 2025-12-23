Describe "Manifests" {
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
    It "core.manifest.json matches core-*.ps1 files" {
        $manifest = Get-Content -Raw -Path (Join-Path $lib "core.manifest.json") | ConvertFrom-Json
        $files = Get-ChildItem -Path $lib -Filter "core-*.ps1" | Where-Object { $_.Name -notin @("core.ps1","core-loader.ps1") } | Sort-Object Name | Select-Object -ExpandProperty Name
        $missing = @($files | Where-Object { $manifest -notcontains $_ })
        $extra = @($manifest | Where-Object { $files -notcontains $_ })
        $missing.Count | Should -Be 0
        $extra.Count | Should -Be 0
    }

    It "handlers.manifest.json matches handlers-*.ps1 files" {
        $manifest = Get-Content -Raw -Path (Join-Path $lib "handlers.manifest.json") | ConvertFrom-Json
        $files = Get-ChildItem -Path $lib -Filter "handlers-*.ps1" | Where-Object { $_.Name -notin @("handlers.ps1","handlers-loader.ps1") } | Sort-Object Name | Select-Object -ExpandProperty Name
        $missing = @($files | Where-Object { $manifest -notcontains $_ })
        $extra = @($manifest | Where-Object { $files -notcontains $_ })
        $missing.Count | Should -Be 0
        $extra.Count | Should -Be 0
    }
}
