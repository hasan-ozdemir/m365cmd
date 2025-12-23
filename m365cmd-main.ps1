$ErrorActionPreference = "Stop"

$global:ScriptRoot = if ($PSScriptRoot) {
    $PSScriptRoot
} else {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not $env:M365CMD_ROOT) {
    $env:M365CMD_ROOT = $global:ScriptRoot
}

. (Join-Path $global:ScriptRoot "lib" "core.ps1")
. (Join-Path $global:ScriptRoot "lib" "handlers.ps1")
. (Join-Path $global:ScriptRoot "lib" "repl.ps1")

Start-M365Cmd
