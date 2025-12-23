# Handler: Tools
# Purpose: Tools command handlers.
function Get-M365CliInstallRoot {
    return (Join-Path $Paths.Tools "m365")
}


function Get-M365CliPath {
    $root = Get-M365CliInstallRoot
    $bin = Join-Path $root "node_modules/.bin"
    $cmd = if ($IsWindows) { "m365.cmd" } else { "m365" }
    $local = Join-Path $bin $cmd
    if (Test-Path $local) { return $local }
    $found = Get-Command -Name "m365" -ErrorAction SilentlyContinue
    if ($found) { return $found.Source }
    return $null
}


function Handle-M365CliCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: m365cli status|install|path|app|run <args...> OR m365cli <m365 args...>"
        return
    }
    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    switch ($sub) {
        "status" {
            $path = Get-M365CliPath
            if (-not $path) {
                Write-Warn "CLI for Microsoft 365 not found. Use: m365cli install"
                return
            }
            Write-Host ("m365 path: " + $path)
            try {
                & $path "--version"
            } catch {
                Write-Warn "Unable to run m365 --version."
            }
        }
        "path" {
            $path = Get-M365CliPath
            if ($path) { Write-Host $path } else { Write-Warn "CLI for Microsoft 365 not found." }
        }
        "install" {
            $npm = Get-Command -Name "npm" -ErrorAction SilentlyContinue
            if (-not $npm) {
                Write-Warn "npm not found. Install Node.js first."
                return
            }
            Ensure-Directories
            $root = Get-M365CliInstallRoot
            if (-not (Test-Path $root)) {
                New-Item -ItemType Directory -Path $root -Force | Out-Null
            }
            Write-Info "Installing CLI for Microsoft 365 locally..."
            try {
                & $npm.Source "install" "--prefix" $root "@pnp/cli-microsoft365"
                Write-Info "Install completed."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "app" {
            Handle-M365CliAppCommand $rest
        }
        "run" {
            $path = Get-M365CliPath
            if (-not $path) {
                Write-Warn "CLI for Microsoft 365 not found. Use: m365cli install"
                return
            }
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: m365cli run <m365 args...>"
                return
            }
            & $path @rest
        }
        default {
            $path = Get-M365CliPath
            if (-not $path) {
                Write-Warn "CLI for Microsoft 365 not found. Use: m365cli install"
                return
            }
            & $path @Args
        }
    }
}


function Get-M365CliAppMapPath {
    return (Join-Path $Paths.Data "m365cli.appmap.json")
}


function Load-M365CliAppMap {
    $path = Get-M365CliAppMapPath
    if (-not (Test-Path $path)) { return @{} }
    try {
        $obj = Get-Content -Raw -Path $path | ConvertFrom-Json
        if (-not $obj) { return @{} }
        $map = @{}
        foreach ($p in $obj.PSObject.Properties) { $map[$p.Name] = [string]$p.Value }
        return $map
    } catch {
        Write-Warn "App map is invalid. Recreating on next save."
        return @{}
    }
}


function Save-M365CliAppMap {
    param([hashtable]$Map)
    if (-not $Map) { $Map = @{} }
    Ensure-Directories
    $path = Get-M365CliAppMapPath
    $Map | ConvertTo-Json -Depth 4 | Set-Content -Path $path -Encoding ASCII
}


function Handle-M365CliAppCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: m365cli app list|set|remove|show|run"
        return
    }
    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $map = Load-M365CliAppMap

    switch ($action) {
        "list" {
            if (-not $map -or $map.Keys.Count -eq 0) {
                Write-Info "No app mappings configured."
                Write-Host "Use: m365cli app set <app> --cmd <m365 command prefix>"
                return
            }
            $out = @()
            foreach ($k in ($map.Keys | Sort-Object)) {
                $out += [pscustomobject]@{ App = $k; Command = $map[$k] }
            }
            $out | Format-Table -AutoSize
        }
        "show" {
            $app = $rest | Select-Object -First 1
            if (-not $app) {
                Write-Warn "Usage: m365cli app show <app>"
                return
            }
            if ($map.ContainsKey($app)) {
                Write-Host $map[$app]
            } else {
                Write-Warn "App mapping not found."
            }
        }
        "set" {
            $app = $rest | Select-Object -First 1
            $parsed = Parse-NamedArgs ($rest | Select-Object -Skip 1)
            $cmd = Get-ArgValue $parsed.Map "cmd"
            if (-not $app -or -not $cmd) {
                Write-Warn "Usage: m365cli app set <app> --cmd <m365 command prefix>"
                return
            }
            $map[$app] = $cmd
            Save-M365CliAppMap $map
            Write-Info "App mapping saved."
        }
        "remove" {
            $app = $rest | Select-Object -First 1
            if (-not $app) {
                Write-Warn "Usage: m365cli app remove <app>"
                return
            }
            if ($map.ContainsKey($app)) {
                $map.Remove($app)
                Save-M365CliAppMap $map
                Write-Info "App mapping removed."
            } else {
                Write-Warn "App mapping not found."
            }
        }
        "run" {
            $app = $rest | Select-Object -First 1
            $args2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            if (-not $app -or -not $map.ContainsKey($app)) {
                Write-Warn "Usage: m365cli app run <app> <args...> (app must be mapped)"
                return
            }
            $path = Get-M365CliPath
            if (-not $path) {
                Write-Warn "CLI for Microsoft 365 not found. Use: m365cli install"
                return
            }
            $prefix = Split-Args $map[$app]
            & $path @prefix @args2
        }
        default {
            Write-Warn "Usage: m365cli app list|set|remove|show|run"
        }
    }
}
