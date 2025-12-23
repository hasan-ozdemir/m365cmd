# Core: Resolve Userdrive
# Purpose: Resolve Userdrive shared utilities.
function Resolve-UserSegment {
    param([string]$User)
    if (-not $User -or $User -eq "me") { return "/me" }
    $u = Resolve-UserObject $User
    if (-not $u) {
        Write-Warn ("User not found: " + $User)
        return $null
    }
    return "/users/" + $u.Id
}

function Resolve-DriveBase {
    param([hashtable]$Map)
    $drive = Get-ArgValue $Map "drive"
    if ($drive) { return "/drives/" + $drive }
    $site = Get-ArgValue $Map "site"
    if ($site) { return "/sites/" + $site + "/drive" }
    $group = Get-ArgValue $Map "group"
    if ($group) { return "/groups/" + $group + "/drive" }
    $user = Get-ArgValue $Map "user"
    $seg = Resolve-UserSegment $user
    if (-not $seg) { return $null }
    return $seg + "/drive"
}

function Resolve-DriveItemId {
    param(
        [string]$Base,
        [string]$ItemId,
        [string]$Path,
        [switch]$Beta
    )
    if ($ItemId) { return $ItemId }
    if (-not $Base) { return $null }
    if (-not $Path) { return $null }
    $p = Normalize-DrivePath $Path
    if (-not $p) { return $null }
    $resp = Invoke-GraphRequest -Method "GET" -Uri ($Base + "/root:/" + $p) -Beta:$Beta
    if ($resp -and $resp.id) { return $resp.id }
    return $null
}

function Resolve-UserObject {
    param([string]$Identity)
    if (-not $Identity) { return $null }
    try {
        return Get-MgUser -UserId $Identity -Property Id,DisplayName,UserPrincipalName -ErrorAction Stop
    } catch {
        $esc = Escape-ODataString $Identity
        $u = Get-MgUser -Filter "userPrincipalName eq '$esc'" -Property Id,DisplayName,UserPrincipalName -All -ErrorAction SilentlyContinue | Select-Object -First 1
        return $u
    }
}
