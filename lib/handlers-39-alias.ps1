# Handler: Alias
# Purpose: Alias and preset management commands.
function Handle-AliasCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: alias list|get|set|remove [--global|--local]"
        return
    }
    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $isGlobal = $parsed.Map.ContainsKey("global")
    $scope = if ($isGlobal) { "global" } else { "local" }
    if (-not $global:Config.aliases) { $global:Config | Add-Member -NotePropertyName aliases -NotePropertyValue @{ global = @{}; local = @{} } -Force }
    if (-not $global:Config.aliases.$scope) { $global:Config.aliases.$scope = @{} }
    $map = $global:Config.aliases.$scope

    switch ($sub) {
        "list" {
            if (-not $map.Keys -or $map.Keys.Count -eq 0) {
                Write-Info "No aliases defined."
                return
            }
            foreach ($k in ($map.Keys | Sort-Object)) {
                Write-Host ("  " + $k + " = " + $map[$k])
            }
        }
        "get" {
            $name = $parsed.Positionals | Select-Object -First 1
            if (-not $name) { Write-Warn "Usage: alias get <name> [--global|--local]"; return }
            if ($map.ContainsKey($name)) { Write-Host $map[$name] } else { Write-Warn "Alias not found." }
        }
        "set" {
            $name = $parsed.Positionals | Select-Object -First 1
            $value = Get-ArgValue $parsed.Map "value"
            if (-not $value) { $value = Get-ArgValue $parsed.Map "command" }
            if (-not $value) { $value = Get-ArgValue $parsed.Map "set" }
            if (-not $name -or -not $value) {
                Write-Warn "Usage: alias set <name> --value <command> [--global|--local]"
                return
            }
            $map[$name] = $value
            Set-ConfigValue ("aliases." + $scope) $map
            Write-Info "Alias saved."
        }
        "remove" {
            $name = $parsed.Positionals | Select-Object -First 1
            if (-not $name) { Write-Warn "Usage: alias remove <name> [--global|--local]"; return }
            if ($map.ContainsKey($name)) {
                $map.Remove($name)
                Set-ConfigValue ("aliases." + $scope) $map
                Write-Info "Alias removed."
            } else {
                Write-Warn "Alias not found."
            }
        }
        default {
            Write-Warn "Usage: alias list|get|set|remove [--global|--local]"
        }
    }
}


function Handle-PresetCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: preset list|get|set|remove|run"
        return
    }
    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    if (-not $global:Config.presets) { $global:Config | Add-Member -NotePropertyName presets -NotePropertyValue @{} -Force }
    $map = $global:Config.presets

    switch ($sub) {
        "list" {
            if (-not $map.Keys -or $map.Keys.Count -eq 0) { Write-Info "No presets defined."; return }
            foreach ($k in ($map.Keys | Sort-Object)) { Write-Host ("  " + $k + " = " + $map[$k]) }
        }
        "get" {
            $name = $parsed.Positionals | Select-Object -First 1
            if (-not $name) { Write-Warn "Usage: preset get <name>"; return }
            if ($map.ContainsKey($name)) { Write-Host $map[$name] } else { Write-Warn "Preset not found." }
        }
        "set" {
            $name = $parsed.Positionals | Select-Object -First 1
            $value = Get-ArgValue $parsed.Map "value"
            if (-not $value) { $value = Get-ArgValue $parsed.Map "command" }
            if (-not $value) { $value = Get-ArgValue $parsed.Map "set" }
            if (-not $name -or -not $value) { Write-Warn "Usage: preset set <name> --value <cmd1; cmd2; ...>"; return }
            $map[$name] = $value
            Set-ConfigValue "presets" $map
            Write-Info "Preset saved."
        }
        "remove" {
            $name = $parsed.Positionals | Select-Object -First 1
            if (-not $name) { Write-Warn "Usage: preset remove <name>"; return }
            if ($map.ContainsKey($name)) {
                $map.Remove($name)
                Set-ConfigValue "presets" $map
                Write-Info "Preset removed."
            } else {
                Write-Warn "Preset not found."
            }
        }
        "run" {
            $name = $parsed.Positionals | Select-Object -First 1
            if (-not $name) { Write-Warn "Usage: preset run <name> [args...]"; return }
            if (-not $map.ContainsKey($name)) { Write-Warn "Preset not found."; return }
            $args = $parsed.Positionals | Select-Object -Skip 1
            $argText = if ($args) { $args -join " " } else { "" }
            $exp = $map[$name]
            $line = if ($exp -like "*{args}*") { $exp.Replace("{args}", $argText) } else { if ($argText) { $exp + " " + $argText } else { $exp } }
            $commands = Split-CommandSequence $line
            if (-not $commands -or $commands.Count -eq 0) { return }
            if (-not (Get-Command Invoke-CommandLine -ErrorAction SilentlyContinue)) {
                Write-Warn "Preset execution is only available in the REPL."
                return
            }
            foreach ($cmdLine in $commands) {
                $cont = Invoke-CommandLine -Line $cmdLine -Depth 0
                if (-not $cont) { break }
            }
        }
        default {
            Write-Warn "Usage: preset list|get|set|remove|run"
        }
    }
}
