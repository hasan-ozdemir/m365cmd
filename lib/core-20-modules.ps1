# Core: Modules
# Purpose: Modules shared utilities.
function Set-LocalModulePath {
    $modulesPath = $Paths.Modules
    if (-not (Test-Path $modulesPath)) {
        New-Item -ItemType Directory -Path $modulesPath -Force | Out-Null
    }
    $sep = [System.IO.Path]::PathSeparator
    $current = $env:PSModulePath -split [regex]::Escape($sep)
    if ($current -notcontains $modulesPath) {
        $env:PSModulePath = ($modulesPath + $sep + $env:PSModulePath)
    }
}



function Test-ModuleAvailable {
    param([string]$Name)
    $module = Get-Module -ListAvailable -Name $Name -ErrorAction SilentlyContinue | Select-Object -First 1
    return $null -ne $module
}


function Install-ModuleIfMissing {
    param([string]$Name)
    if (Test-ModuleAvailable $Name) { return $true }
    $auto = Parse-Bool $global:Config.modules.autoInstall $false
    if (-not $auto) {
        Write-Warn ($Name + " module not found. Use: module install " + $Name)
        return $false
    }
    Write-Warn ($Name + " module not found. Installing...")
    try {
        Set-LocalModulePath
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    } catch {}
    if (-not (Get-Command -Name Save-Module -ErrorAction SilentlyContinue)) {
        Write-Err "Save-Module is not available. Install PowerShellGet."
        return $false
    }
    try {
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -Force -Scope CurrentUser | Out-Null
        }
        Save-Module -Name $Name -Path $Paths.Modules -Force -ErrorAction Stop
        Write-Info ("Installed module: " + $Name)
    } catch {
        Write-Err ("Module install failed: " + $_.Exception.Message)
        return $false
    }
    return (Test-ModuleAvailable $Name)
}



function Ensure-GraphModule {
    if (-not (Install-ModuleIfMissing "Microsoft.Graph")) {
        return $false
    }
    Import-Module Microsoft.Graph -ErrorAction SilentlyContinue | Out-Null
    return $true
}



function Ensure-ModuleLoaded {
    param(
        [string]$Name,
        [switch]$UseWindowsPowerShell
    )
    if (-not (Install-ModuleIfMissing $Name)) {
        return $false
    }
    try {
        if ($UseWindowsPowerShell -and $PSVersionTable.PSVersion.Major -ge 7 -and $IsWindows) {
            Import-Module $Name -UseWindowsPowerShell -ErrorAction Stop | Out-Null
        } else {
            Import-Module $Name -ErrorAction Stop | Out-Null
        }
        return $true
    } catch {
        Write-Err ("Failed to import module: " + $Name + ". " + $_.Exception.Message)
        return $false
    }
}



function Ensure-MsalModule {
    if (-not (Install-ModuleIfMissing "MSAL.PS")) {
        return $false
    }
    try {
        Import-Module MSAL.PS -ErrorAction Stop | Out-Null
        return $true
    } catch {
        Write-Err ("Failed to import module: MSAL.PS. " + $_.Exception.Message)
        return $false
    }
}

$global:AppTokenCache = @{}


