# Core: Base
# Purpose: Base shared utilities.
function Write-Info {
    param([string]$Message)
    Write-Host $Message
}



function Write-Warn {
    param([string]$Message)
    Write-Host ("WARN: " + $Message) -ForegroundColor Yellow
}



function Write-Err {
    param([string]$Message)
    Write-Host ("ERROR: " + $Message) -ForegroundColor Red
}

$ScriptRoot = if ($global:ScriptRoot -and (Test-Path $global:ScriptRoot)) {
    $global:ScriptRoot
} elseif ($env:M365CMD_ROOT -and (Test-Path $env:M365CMD_ROOT)) {
    $env:M365CMD_ROOT
} else {
    Split-Path -Parent $PSCommandPath
}

$Paths = [ordered]@{
    Root    = $ScriptRoot
    Config  = Join-Path $ScriptRoot "m365cmd.config.json"
    Logs    = Join-Path $ScriptRoot "logs"
    Data    = Join-Path $ScriptRoot "data"
    Modules = Join-Path $ScriptRoot "modules"
    Tools   = Join-Path $ScriptRoot "tools"
}



function Ensure-Directories {
    foreach ($p in @($Paths.Logs, $Paths.Data, $Paths.Modules, $Paths.Tools)) {
        if (-not (Test-Path $p)) {
            New-Item -ItemType Directory -Path $p -Force | Out-Null
        }
    }
    $ppDir = Join-Path $Paths.Data "pp"
    if (-not (Test-Path $ppDir)) {
        New-Item -ItemType Directory -Path $ppDir -Force | Out-Null
    }
}



