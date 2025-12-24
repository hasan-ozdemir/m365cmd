# Handler: Apps
# Purpose: App catalog helpers and portal/CLI bridging.
function Get-AppCatalogPath {
    return (Join-Path $PSScriptRoot "apps.catalog.json")
}

function Load-AppCatalog {
    $path = Get-AppCatalogPath
    if (-not (Test-Path $path)) {
        Write-Warn "App catalog not found."
        return @()
    }
    try {
        $raw = Get-Content -Raw -Path $path
        return (ConvertFrom-Json $raw)
    } catch {
        Write-Warn ("App catalog invalid: " + $_.Exception.Message)
        return @()
    }
}

function Normalize-AppKey {
    param([string]$Text)
    if (-not $Text) { return "" }
    return ($Text.ToLowerInvariant() -replace "[^a-z0-9]", "")
}

function Resolve-AppEntry {
    param([string]$Name)
    if (-not $Name) { return $null }
    $apps = Load-AppCatalog
    if (-not $apps -or $apps.Count -eq 0) { return $null }
    $key = Normalize-AppKey $Name
    foreach ($a in $apps) {
        if ((Normalize-AppKey $a.id) -eq $key) { return $a }
        if ((Normalize-AppKey $a.name) -eq $key) { return $a }
        if ($a.aliases) {
            foreach ($al in @($a.aliases)) {
                if ((Normalize-AppKey $al) -eq $key) { return $a }
            }
        }
    }
    return $null
}

function Format-AppMethods {
    param([object]$Methods)
    if (-not $Methods) { return "" }
    $list = @($Methods)
    return ($list -join ",")
}

function Handle-AppsCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        $InputArgs = @("list")
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $asJson = $parsed.Map.ContainsKey("json")

    switch ($sub) {
        "list" {
            $apps = Load-AppCatalog
            if ($asJson) {
                $apps | ConvertTo-Json -Depth 6
            } else {
                $apps | Sort-Object id | ForEach-Object {
                    [pscustomobject]@{
                        Id      = $_.id
                        Name    = $_.name
                        Methods = (Format-AppMethods $_.methods)
                        Command = $_.command
                        Portal  = $_.portal
                    }
                } | Format-Table -AutoSize
            }
        }
        "get" {
            $name = $parsed.Positionals | Select-Object -First 1
            if (-not $name) { Write-Warn "Usage: apps get <appId|name>"; return }
            $entry = Resolve-AppEntry $name
            if (-not $entry) { Write-Warn "App not found."; return }
            $entry | ConvertTo-Json -Depth 6
        }
        "open" {
            $name = $parsed.Positionals | Select-Object -First 1
            if (-not $name) { Write-Warn "Usage: apps open <appId|name>"; return }
            $entry = Resolve-AppEntry $name
            if (-not $entry) { Write-Warn "App not found."; return }
            if ($entry.portal) { Write-Host $entry.portal } else { Write-Warn "No portal URL defined." }
        }
        "cli" {
            $name = $parsed.Positionals | Select-Object -First 1
            $args2 = @($parsed.Positionals | Select-Object -Skip 1)
            if (-not $name) { Write-Warn "Usage: apps cli <appId|name> [--cmd <prefix>] <args...>"; return }
            if (-not (Get-Command Handle-M365CliCommand -ErrorAction SilentlyContinue)) {
                Write-Warn "m365cli command handler not available."
                return
            }
            $entry = Resolve-AppEntry $name
            if (-not $entry) { Write-Warn "App not found."; return }
            $prefix = Get-ArgValue $parsed.Map "cmd"
            if (-not $prefix -and (Get-Command Load-M365CliAppMap -ErrorAction SilentlyContinue)) {
                $map = Load-M365CliAppMap
                if ($map -and $map.ContainsKey($entry.id)) { $prefix = $map[$entry.id] }
            }
            if (-not $prefix) {
                Write-Warn "No CLI mapping. Use: m365cli app set <app> --cmd <m365 prefix>"
                return
            }
            $parts = Split-Args $prefix
            Handle-M365CliCommand (@("run") + $parts + $args2)
        }
        default {
            Write-Warn "Usage: apps list|get|open|cli [--json]"
        }
    }
}

