Describe "REPL and parsing" {
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
        $global:Config = Normalize-Config (Get-DefaultConfig)
    }
    It "splits command sequences" {
        $parts = Split-CommandSequence "user list; group list ;  role list"
        $parts.Count | Should -Be 3
    }

    It "parses named args" {
        $parsed = Parse-NamedArgs @("--name","Test","--flag","true","pos1")
        $parsed.Map["name"] | Should -Be "Test"
        $parsed.Positionals[0] | Should -Be "pos1"
    }

    It "splits args into tokens" {
        $parts = Split-Args "user list"
        $parts.Count | Should -Be 2
        $parts[0] | Should -Be "user"
        $parts[1] | Should -Be "list"
    }

    It "expands aliases" {
        $expanded = Expand-AliasCommand -Cmd "u" -InputArgs @("list") -IsGlobal:$false
        $expanded.Count | Should -BeGreaterThan 0
    }
}
