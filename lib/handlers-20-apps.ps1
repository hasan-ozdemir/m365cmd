# Handler: Apps
# Purpose: App catalog helpers and portal/command mappings.
function Get-AppCatalogPath {
    return (Join-Path $PSScriptRoot "apps.catalog.json")
}

function Get-AppMapPath {
    return (Join-Path $Paths.Data "appmap.json")
}

function Load-AppMap {
    $path = Get-AppMapPath
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

function Save-AppMap {
    param([hashtable]$Map)
    if (-not $Map) { $Map = @{} }
    Ensure-Directories
    $path = Get-AppMapPath
    $Map | ConvertTo-Json -Depth 4 | Set-Content -Path $path -Encoding ASCII
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

function Join-CommandLine {
    param([string[]]$Parts)
    if (-not $Parts) { return "" }
    $escaped = foreach ($p in $Parts) {
        if ($null -eq $p) { continue }
        if ($p -match "\\s") {
            $clean = $p -replace '"', "'"
            '"' + $clean + '"'
        } else {
            $p
        }
    }
    return ($escaped -join " ")
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

function Resolve-AppMapKey {
    param([string]$Name)
    if (-not $Name) { return $null }
    $entry = Resolve-AppEntry $Name
    if ($entry -and $entry.id) { return $entry.id }
    return $Name
}

function Resolve-AppMappedPrefix {
    param([hashtable]$Map, [string]$Name)
    if (-not $Map -or -not $Name) { return $null }
    $key = Resolve-AppMapKey $Name
    if ($key -and $Map.ContainsKey($key)) { return $Map[$key] }
    $norm = Normalize-AppKey $Name
    foreach ($k in $Map.Keys) {
        if ((Normalize-AppKey $k) -eq $norm) { return $Map[$k] }
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
        "map" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: apps map list|set|remove|show|run"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest2
            $map = Load-AppMap
            switch ($action) {
                "list" {
                    if (-not $map -or $map.Keys.Count -eq 0) {
                        Write-Info "No app mappings configured."
                        return
                    }
                    $out = @()
                    foreach ($k in ($map.Keys | Sort-Object)) {
                        $out += [pscustomobject]@{ App = $k; Command = $map[$k] }
                    }
                    $out | Format-Table -AutoSize
                }
                "show" {
                    $name = $parsed2.Positionals | Select-Object -First 1
                    if (-not $name) { Write-Warn "Usage: apps map show <appId|name>"; return }
                    $cmd = Resolve-AppMappedPrefix $map $name
                    if ($cmd) { Write-Host $cmd } else { Write-Warn "App mapping not found." }
                }
                "set" {
                    $name = $parsed2.Positionals | Select-Object -First 1
                    $cmd = Get-ArgValue $parsed2.Map "cmd"
                    if (-not $name -or -not $cmd) {
                        Write-Warn "Usage: apps map set <appId|name> --cmd <m365cmd prefix>"
                        return
                    }
                    $key = Resolve-AppMapKey $name
                    $map[$key] = $cmd
                    Save-AppMap $map
                    Write-Info "App mapping saved."
                }
                "remove" {
                    $name = $parsed2.Positionals | Select-Object -First 1
                    if (-not $name) { Write-Warn "Usage: apps map remove <appId|name>"; return }
                    $key = Resolve-AppMapKey $name
                    if ($map.ContainsKey($key)) {
                        $map.Remove($key)
                        Save-AppMap $map
                        Write-Info "App mapping removed."
                    } else {
                        Write-Warn "App mapping not found."
                    }
                }
                "run" {
                    $name = $parsed2.Positionals | Select-Object -First 1
                    $args2 = @($parsed2.Positionals | Select-Object -Skip 1)
                    if (-not $name) { Write-Warn "Usage: apps map run <appId|name> <args...>"; return }
                    $prefix = Resolve-AppMappedPrefix $map $name
                    if (-not $prefix) {
                        Write-Warn "No mapping found. Use: apps map set <app> --cmd <m365cmd prefix>"
                        return
                    }
                    $parts = Split-Args $prefix
                    $line = Join-CommandLine ($parts + $args2)
                    Invoke-CommandLine -Line $line | Out-Null
                }
                default {
                    Write-Warn "Usage: apps map list|set|remove|show|run"
                }
            }
        }
        "cli" {
            $name = $parsed.Positionals | Select-Object -First 1
            $args2 = @($parsed.Positionals | Select-Object -Skip 1)
            if (-not $name) { Write-Warn "Usage: apps cli <appId|name> [--cmd <prefix>] <args...>"; return }
            $prefix = Get-ArgValue $parsed.Map "cmd"
            $map = Load-AppMap
            if (-not $prefix) { $prefix = Resolve-AppMappedPrefix $map $name }
            if (-not $prefix) {
                Write-Warn "No mapping found. Use: apps map set <app> --cmd <m365cmd prefix>"
                return
            }
            $parts = Split-Args $prefix
            $line = Join-CommandLine ($parts + $args2)
            Invoke-CommandLine -Line $line | Out-Null
        }
        "run" {
            Handle-AppsCommand (@("cli") + $rest)
        }
        default {
            Write-Warn "Usage: apps list|get|open|map|cli|run [--json]"
        }
    }
}
