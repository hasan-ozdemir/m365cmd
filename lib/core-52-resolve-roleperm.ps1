# Core: Resolve Roleperm
# Purpose: Resolve Roleperm shared utilities.
function Resolve-DirectoryRole {
    param([string]$Role)
    if (-not $Role) { return $null }
    if ($Role -match "^[0-9a-fA-F-]{36}$") {
        try {
            return Get-MgDirectoryRole -DirectoryRoleId $Role -ErrorAction Stop
        } catch {}
    }
    $esc = Escape-ODataString $Role
    return Get-MgDirectoryRole -Filter "displayName eq '$esc'" -All -ErrorAction SilentlyContinue | Select-Object -First 1
}

function Ensure-DirectoryRole {
    param([string]$RoleNameOrId)
    $role = Resolve-DirectoryRole $RoleNameOrId
    if ($role) { return $role }
    $esc = Escape-ODataString $RoleNameOrId
    $template = Get-MgDirectoryRoleTemplate -Filter "displayName eq '$esc'" -All -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $template) {
        return $null
    }
    try {
        New-MgDirectoryRole -RoleTemplateId $template.Id | Out-Null
        Start-Sleep -Seconds 2
        return Resolve-DirectoryRole $RoleNameOrId
    } catch {
        return $null
    }
}

function Convert-RequiredResourceAccessList {
    param([object[]]$RequiredResourceAccess)
    $list = @()
    foreach ($r in @($RequiredResourceAccess)) {
        if ($null -eq $r) { continue }
        $access = @()
        foreach ($a in @($r.ResourceAccess)) {
            $access += @{
                Id   = $a.Id
                Type = $a.Type
            }
        }
        $list += @{
            ResourceAppId = $r.ResourceAppId
            ResourceAccess = $access
        }
    }
    return $list
}

function Get-GraphServicePrincipal {
    if ($global:GraphServicePrincipal) { return $global:GraphServicePrincipal }
    $sp = Get-MgServicePrincipal -Filter "displayName eq 'Microsoft Graph'" -All -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $sp) {
        $sp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -All -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    $global:GraphServicePrincipal = $sp
    return $sp
}

function Resolve-GraphPermission {
    param(
        [object]$GraphSp,
        [string]$Name,
        [string]$Type
    )
    if (-not $GraphSp -or -not $Name) { return $null }
    $t = if ($Type) { $Type.ToLowerInvariant() } else { "" }
    $scope = $null
    $role = $null
    if ($t -eq "delegated" -or $t -eq "scope") {
        $scope = $GraphSp.Oauth2PermissionScopes | Where-Object { $_.Value -eq $Name } | Select-Object -First 1
    } elseif ($t -eq "application" -or $t -eq "role") {
        $role = $GraphSp.AppRoles | Where-Object { $_.Value -eq $Name } | Select-Object -First 1
    } else {
        $scope = $GraphSp.Oauth2PermissionScopes | Where-Object { $_.Value -eq $Name } | Select-Object -First 1
        $role = $GraphSp.AppRoles | Where-Object { $_.Value -eq $Name } | Select-Object -First 1
        if ($scope -and $role) {
            Write-Warn "Permission name matches both delegated and application types. Use --type delegated|application."
            return $null
        }
    }
    if ($scope) {
        return @{ Id = $scope.Id; Type = "Scope"; Value = $scope.Value }
    }
    if ($role) {
        return @{ Id = $role.Id; Type = "Role"; Value = $role.Value }
    }
    return $null
}
