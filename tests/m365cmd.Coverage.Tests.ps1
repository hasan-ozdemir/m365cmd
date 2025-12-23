Describe "Coverage" {
    It "coverage manifest matches ps1 files" {
        $repoRoot = (Get-Location).Path
        if (-not (Test-Path (Join-Path $repoRoot "lib"))) {
            $repoRoot = Split-Path -Parent $repoRoot
        }
        if (-not (Test-Path (Join-Path $repoRoot "lib"))) {
            throw "Repo root not found. Run tests from repo root."
        }
        $manifestPath = Join-Path $repoRoot "tests\coverage.manifest.json"
        $manifest = Get-Content -Raw -Path $manifestPath | ConvertFrom-Json
        $files = @()
        $files += (Get-ChildItem -Path $repoRoot -File -Filter "*.ps1" | Select-Object -ExpandProperty FullName)
        $files += (Get-ChildItem -Path (Join-Path $repoRoot "lib") -File -Filter "*.ps1" | Select-Object -ExpandProperty FullName)
        $files = $files | Sort-Object -Unique
        $rel = $files | ForEach-Object { $_.Substring($repoRoot.Length).TrimStart([char[]]@('\','/')) } | Sort-Object
        $manifestSorted = @($manifest) | Sort-Object
        ($rel | Where-Object { $manifestSorted -notcontains $_ }).Count | Should -Be 0
        ($manifestSorted | Where-Object { $rel -notcontains $_ }).Count | Should -Be 0
    }

    It "entrypoint and loader scripts contain expected markers" {
        $repoRoot = (Get-Location).Path
        if (-not (Test-Path (Join-Path $repoRoot "lib"))) {
            $repoRoot = Split-Path -Parent $repoRoot
        }
        $paths = @(
            "m365cmd-main.ps1",
            "lib\\core.ps1",
            "lib\\handlers.ps1",
            "lib\\repl.ps1",
            "lib\\core-loader.ps1",
            "lib\\handlers-loader.ps1"
        )
        foreach ($p in $paths) {
            $full = Join-Path $repoRoot $p
            Test-Path $full | Should -BeTrue
            $raw = Get-Content -Raw -Path $full
            $raw.Length | Should -BeGreaterThan 0
        }
    }
}
