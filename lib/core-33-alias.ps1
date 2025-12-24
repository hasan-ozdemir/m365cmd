# Core: Alias
# Purpose: Alias and preset helpers for the REPL.
function Split-CommandSequence {
    param([string]$Line)
    if (-not $Line) { return @() }
    $parts = @()
    $current = ""
    $inSingle = $false
    $inDouble = $false
    foreach ($ch in $Line.ToCharArray()) {
        if ($ch -eq "'" -and -not $inDouble) {
            $inSingle = -not $inSingle
            $current += $ch
            continue
        }
        if ($ch -eq '"' -and -not $inSingle) {
            $inDouble = -not $inDouble
            $current += $ch
            continue
        }
        if ($ch -eq ';' -and -not $inSingle -and -not $inDouble) {
            $seg = $current.Trim()
            if ($seg) { $parts += $seg }
            $current = ""
            continue
        }
        $current += $ch
    }
    $seg = $current.Trim()
    if ($seg) { $parts += $seg }
    return ,$parts
}


function Get-AliasMap {
    param([bool]$Global = $false)
    if (-not $global:Config -or -not $global:Config.aliases) { return @{} }
    if ($Global) {
        if (-not $global:Config.aliases.global) { return @{} }
        return $global:Config.aliases.global
    }
    if (-not $global:Config.aliases.local) { return @{} }
    return $global:Config.aliases.local
}


function Get-PresetMap {
    if (-not $global:Config -or -not $global:Config.presets) { return @{} }
    return $global:Config.presets
}


function Expand-AliasCommand {
    param(
        [string]$Cmd,
        [string[]]$InputArgs,
        [bool]$IsGlobal = $false
    )
    $map = Get-AliasMap -Global:$IsGlobal
    if (-not $map) { return $null }
    $key = $null
    foreach ($k in $map.Keys) {
        if ($k.ToLowerInvariant() -eq $Cmd.ToLowerInvariant()) { $key = $k; break }
    }
    if (-not $key) { return $null }
    $exp = $map[$key]
    if (-not $exp) { return $null }
    $argText = if ($InputArgs -and $InputArgs.Count -gt 0) { $InputArgs -join " " } else { "" }
    $line = if ($exp -like "*{args}*") { $exp.Replace("{args}", $argText) } else { if ($argText) { $exp + " " + $argText } else { $exp } }
    $lines = Split-CommandSequence $line
    return ,$lines
}
