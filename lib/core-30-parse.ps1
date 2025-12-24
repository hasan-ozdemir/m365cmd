# Core: Parse
# Purpose: Parse shared utilities.
function Parse-Value {
    param([string]$Raw)
    if ($null -eq $Raw) { return $null }
    try {
        return (ConvertFrom-Json -InputObject $Raw -ErrorAction Stop)
    } catch {
        return $Raw
    }
}

function Split-Args {
    param([string]$Line)
    $pattern = '("[^"]*"|''[^'']*''|\S+)'
    $matches = [regex]::Matches($Line, $pattern)
    $parts = @()
    foreach ($m in $matches) {
        $token = $m.Value
        if (($token.StartsWith('"') -and $token.EndsWith('"')) -or ($token.StartsWith("'") -and $token.EndsWith("'"))) {
            $token = $token.Substring(1, $token.Length - 2)
        }
        if ($token.Length -gt 0) {
            $parts += $token
        }
    }
    return ,$parts
}

function Parse-NamedArgs {
    param([string[]]$InputArgs)
    $map = @{}
    $positional = @()
    for ($i = 0; $i -lt $InputArgs.Count; $i++) {
        $arg = $InputArgs[$i]
        if ($arg.StartsWith("--")) {
            $key = $arg.Substring(2)
            $value = $true
            if ($key -like "*=*") {
                $kv = $key -split "=", 2
                $key = $kv[0]
                $value = $kv[1]
            } elseif (($i + 1) -lt $InputArgs.Count -and -not $InputArgs[$i + 1].StartsWith("--")) {
                $value = $InputArgs[$i + 1]
                $i++
            }
            $map[$key] = $value
        } else {
            $positional += $arg
        }
    }
    return [pscustomobject]@{
        Map        = $map
        Positionals = $positional
    }
}

function Parse-CommaList {
    param([string]$Value)
    if (-not $Value) { return @() }
    return @($Value -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
}

function Parse-GuidList {
    param([string]$Value)
    $items = Parse-CommaList $Value
    $list = @()
    foreach ($item in $items) {
        try {
            $list += [Guid]$item
        } catch {
            Write-Warn ("Invalid GUID: " + $item)
        }
    }
    return $list
}

function Parse-KvPairs {
    param([string]$Value)
    $pairs = @{}
    if (-not $Value) { return $pairs }
    $items = $Value -split ","
    foreach ($item in $items) {
        $kv = $item -split "=", 2
        if ($kv.Count -lt 2) { continue }
        $key = $kv[0].Trim()
        $raw = $kv[1].Trim()
        if ($key -eq "") { continue }
        $pairs[$key] = Parse-Value $raw
    }
    return $pairs
}

function Parse-Bool {
    param([object]$Value, [bool]$Default = $false)
    if ($Value -is [bool]) { return $Value }
    if ($null -eq $Value) { return $Default }
    $s = $Value.ToString().ToLowerInvariant()
    if ($s -in @("1", "true", "yes", "y")) { return $true }
    if ($s -in @("0", "false", "no", "n")) { return $false }
    return $Default
}
