# Core: Resolve Appgroup
# Purpose: Resolve Appgroup shared utilities.
function Resolve-ApplicationObject {
    param([string]$Identity)
    if (-not $Identity) { return $null }
    $app = $null
    if ($Identity -match "^[0-9a-fA-F-]{36}$") {
        try {
            $app = Get-MgApplication -ApplicationId $Identity -ErrorAction Stop
        } catch {}
    }
    if (-not $app) {
        $esc = Escape-ODataString $Identity
        $app = Get-MgApplication -Filter "appId eq '$esc'" -All -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    return $app
}

function Resolve-GroupObject {
    param([string]$Identity)
    if (-not $Identity) { return $null }
    $group = $null
    if ($Identity -match "^[0-9a-fA-F-]{36}$") {
        try {
            $group = Get-MgGroup -GroupId $Identity -ErrorAction Stop
        } catch {}
    }
    if (-not $group) {
        $esc = Escape-ODataString $Identity
        $group = Get-MgGroup -Filter "displayName eq '$esc'" -All -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    return $group
}

function Resolve-ServicePrincipalObject {
    param([string]$Identity)
    if (-not $Identity) { return $null }
    $sp = $null
    if ($Identity -match "^[0-9a-fA-F-]{36}$") {
        try {
            $sp = Get-MgServicePrincipal -ServicePrincipalId $Identity -ErrorAction Stop
        } catch {}
    }
    if (-not $sp) {
        $esc = Escape-ODataString $Identity
        $sp = Get-MgServicePrincipal -Filter "appId eq '$esc'" -All -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    if (-not $sp) {
        $esc = Escape-ODataString $Identity
        $sp = Get-MgServicePrincipal -Filter "displayName eq '$esc'" -All -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    return $sp
}
