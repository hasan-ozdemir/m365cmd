Describe "Functions" {
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
        function script:Get-FunctionsInFile {
            param([string]$Path)
            $raw = Get-Content -Raw -Path $Path
            $matches = [regex]::Matches($raw, "(?m)^function\s+([A-Za-z0-9_-]+)\b")
            return ($matches | ForEach-Object { $_.Groups[1].Value })
        }
    }
    It "all core functions are defined after load" {
        $coreFiles = Get-ChildItem -Path $lib -Filter "core-*.ps1" | Where-Object { $_.Name -notin @("core.ps1","core-loader.ps1") }
        $names = @()
        foreach ($f in $coreFiles) { $names += Get-FunctionsInFile $f.FullName }
        $names = $names | Sort-Object -Unique
        foreach ($name in $names) {
            (Get-Command $name -ErrorAction SilentlyContinue) | Should -Not -BeNullOrEmpty
        }
    }

    It "all handler functions are defined after load" {
        $handlerFiles = Get-ChildItem -Path $lib -Filter "handlers-*.ps1" | Where-Object { $_.Name -notin @("handlers.ps1","handlers-loader.ps1") }
        $names = @()
        foreach ($f in $handlerFiles) { $names += Get-FunctionsInFile $f.FullName }
        $names = $names | Sort-Object -Unique
        foreach ($name in $names) {
            (Get-Command $name -ErrorAction SilentlyContinue) | Should -Not -BeNullOrEmpty
        }
    }
}
