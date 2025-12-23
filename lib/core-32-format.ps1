# Core: Format
# Purpose: Format shared utilities.
function Convert-SettingValuesFromPairs {
    param([hashtable]$Pairs)
    $values = @()
    if (-not $Pairs) { return $values }
    foreach ($k in $Pairs.Keys) {
        $values += @{
            name  = $k
            value = [string]$Pairs[$k]
        }
    }
    return $values
}

function Normalize-ShareRoles {
    param([string]$Value)
    $roles = @()
    foreach ($r in (Parse-CommaList $Value)) {
        $v = $r.ToLowerInvariant()
        switch ($v) {
            "can-edit" { $roles += "write" }
            "edit" { $roles += "write" }
            "write" { $roles += "write" }
            "can-view" { $roles += "read" }
            "view" { $roles += "read" }
            "read" { $roles += "read" }
            default { $roles += $r }
        }
    }
    return @($roles | Select-Object -Unique)
}

function Write-GraphTable {
    param(
        [object]$Items,
        [string[]]$Properties
    )
    if (-not $Items) {
        Write-Info "No items found."
        return
    }
    $Items | Select-Object -Property $Properties | Format-Table -AutoSize
}

function Normalize-DrivePath {
    param([string]$Path)
    if (-not $Path) { return "" }
    $p = $Path -replace "\\", "/"
    $p = $p.TrimStart("/")
    return $p
}
