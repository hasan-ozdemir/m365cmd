# Core: Context
# Purpose: CLI-like context file utilities.
function Get-ContextStorePath {
    return (Join-Path $Paths.Data "context.json")
}


function ConvertTo-ContextMap {
    param([object]$Obj)
    if ($null -eq $Obj) { return @{} }
    if ($Obj -is [hashtable]) { return $Obj }
    $map = @{}
    foreach ($p in $Obj.PSObject.Properties) {
        $map[$p.Name] = $p.Value
    }
    return $map
}


function Load-Context {
    $path = Get-ContextStorePath
    if (-not (Test-Path $path)) { return @{} }
    try {
        $raw = Get-Content -Raw -Path $path
        if ($null -eq $raw -or $raw.Trim() -eq "") { return @{} }
        try {
            $obj = ConvertFrom-Json -InputObject $raw -AsHashtable
            return (ConvertTo-ContextMap $obj)
        } catch {
            $obj = ConvertFrom-Json -InputObject $raw
            return (ConvertTo-ContextMap $obj)
        }
    } catch {
        Write-Warn "Context file is invalid. Recreate it with: context init"
        return @{}
    }
}


function Save-Context {
    param([hashtable]$Map)
    if ($null -eq $Map) { $Map = @{} }
    Ensure-Directories
    $path = Get-ContextStorePath
    $Map | ConvertTo-Json -Depth 6 | Set-Content -Path $path -Encoding ASCII
}


function Remove-Context {
    $path = Get-ContextStorePath
    if (Test-Path $path) {
        Remove-Item -Path $path -Force
        return $true
    }
    return $false
}


function Set-ContextOption {
    param(
        [string]$Name,
        [object]$Value
    )
    if (-not $Name) { return $false }
    $map = Load-Context
    $map[$Name] = $Value
    Save-Context $map
    return $true
}


function Remove-ContextOption {
    param([string]$Name)
    if (-not $Name) { return $false }
    $map = Load-Context
    if (-not $map.ContainsKey($Name)) { return $false }
    $map.Remove($Name) | Out-Null
    if ($map.Count -eq 0) {
        Remove-Context | Out-Null
    } else {
        Save-Context $map
    }
    return $true
}
