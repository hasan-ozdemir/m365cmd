# m365cmd test runner
[CmdletBinding()]
param(
    [switch]$Modules,
    [switch]$Integration,
    [switch]$Write,
    [switch]$Strict,
    [switch]$InstallPester,
    [switch]$PassThru
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$tests = $PSScriptRoot

if (-not (Get-Command Invoke-Pester -ErrorAction SilentlyContinue)) {
    if ($InstallPester) {
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        } catch {}
        if (-not (Get-Command -Name Save-Module -ErrorAction SilentlyContinue)) {
            Write-Host "ERROR: PowerShellGet not available. Install Pester manually." -ForegroundColor Red
            exit 1
        }
        try {
            if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                Install-PackageProvider -Name NuGet -Force -Scope CurrentUser | Out-Null
            }
            Save-Module -Name Pester -Path (Join-Path $root "modules") -Force -ErrorAction Stop
            Import-Module Pester -ErrorAction Stop | Out-Null
        } catch {
            Write-Host ("ERROR: Failed to install Pester. " + $_.Exception.Message) -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host "ERROR: Pester not found. Re-run with -InstallPester." -ForegroundColor Red
        exit 1
    }
}

$exclude = @("integration","modules","write")
if ($Modules) { $exclude = $exclude | Where-Object { $_ -ne "modules" } }
if ($Integration) { $exclude = $exclude | Where-Object { $_ -ne "integration" } }
if ($Write) {
    $exclude = $exclude | Where-Object { $_ -notin @("write","integration") }
}

$config = New-PesterConfiguration
$config.Run.Path = $tests
if ($exclude.Count -gt 0) { $config.Filter.ExcludeTag = $exclude }
if ($Strict) { $config.Run.Exit = $true }
$config.Output.Verbosity = "Detailed"
if ($PassThru) { $config.Run.PassThru = $true }

Invoke-Pester -Configuration $config
