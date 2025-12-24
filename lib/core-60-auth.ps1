# Core: Auth
# Purpose: Auth shared utilities.
$global:AppTokenCache = @{}

function Invoke-Login {
    param([string]$Mode)
    if (-not (Ensure-GraphModule)) { return }
    $cfg = $global:Config
    $scopes = $cfg.auth.scopes
    $contextScope = $cfg.auth.contextScope

    try {
        if ($Mode -and $Mode.ToLowerInvariant() -eq "device") {
            Connect-MgGraph -Scopes $scopes -ContextScope $contextScope -UseDeviceCode | Out-Null
        } else {
            Connect-MgGraph -Scopes $scopes -ContextScope $contextScope | Out-Null
        }
        Write-Info "Connected to Microsoft Graph."
        Register-GraphConnection
    } catch {
        Write-Err $_.Exception.Message
    }
}



function Invoke-Logout {
    if (Get-Command -Name Disconnect-MgGraph -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph | Out-Null
        Write-Info "Disconnected."
    } else {
        Write-Warn "Graph module is not loaded."
    }
}


function Get-ConnectionStorePath {
    return (Join-Path $Paths.Data "connections.json")
}


function Load-Connections {
    $path = Get-ConnectionStorePath
    if (-not (Test-Path $path)) { return @() }
    try {
        $data = Get-Content -Raw -Path $path | ConvertFrom-Json
        if ($null -eq $data) { return @() }
        return @($data)
    } catch {
        Write-Warn "Connections file is invalid. Recreating on next save."
        return @()
    }
}


function Save-Connections {
    param([object[]]$Items)
    if ($null -eq $Items) { $Items = @() }
    Ensure-Directories
    $path = Get-ConnectionStorePath
    $Items | ConvertTo-Json -Depth 6 | Set-Content -Path $path -Encoding ASCII
}


function Get-ConnectionByName {
    param([string]$Name)
    if (-not $Name) { return $null }
    $list = Load-Connections
    foreach ($c in $list) {
        if ($c.name -eq $Name) { return $c }
    }
    return $null
}


function Get-ActiveConnection {
    $list = Load-Connections
    foreach ($c in $list) {
        if ($c.active) { return $c }
    }
    return $null
}


function Set-ActiveConnectionByName {
    param([string]$Name)
    if (-not $Name) { return $false }
    $list = Load-Connections
    $found = $false
    foreach ($c in $list) {
        if ($c.name -eq $Name) {
            $c.active = $true
            $found = $true
        } else {
            $c.active = $false
        }
    }
    if ($found) { Save-Connections $list }
    return $found
}


function Rename-Connection {
    param(
        [string]$Name,
        [string]$NewName
    )
    if (-not $Name -or -not $NewName) { return $false }
    if ($Name -eq $NewName) { return $false }
    $list = Load-Connections
    foreach ($c in $list) {
        if ($c.name -eq $Name) {
            $c.name = $NewName
            Save-Connections $list
            return $true
        }
    }
    return $false
}


function Remove-Connection {
    param([string]$Name)
    if (-not $Name) { return $false }
    $list = Load-Connections
    $newList = @($list | Where-Object { $_.name -ne $Name })
    if ($newList.Count -eq $list.Count) { return $false }
    Save-Connections $newList
    return $true
}


function Register-GraphConnection {
    param([string]$Name)
    $ctx = Get-MgContextSafe
    if (-not $ctx) { return $null }
    $list = Load-Connections

    $name = $Name
    if (-not $name) {
        if ($ctx.Account) { $name = $ctx.Account } else { $name = ("tenant-" + $ctx.TenantId) }
    }
    if (-not $name) { $name = ("conn-" + [guid]::NewGuid().ToString("n")) }

    $entry = $null
    foreach ($c in $list) {
        if ($c.name -eq $name) { $entry = $c; break }
    }
    if (-not $entry) {
        $entry = [pscustomobject]@{
            name        = $name
            connectedAs = $ctx.Account
            authType    = "delegated"
            tenantId    = $ctx.TenantId
            scopes      = @($ctx.Scopes)
            active      = $true
            updatedAt   = (Get-Date).ToString("s")
        }
        $list += $entry
    } else {
        $entry.connectedAs = $ctx.Account
        $entry.authType = "delegated"
        $entry.tenantId = $ctx.TenantId
        $entry.scopes = @($ctx.Scopes)
        $entry.active = $true
        $entry.updatedAt = (Get-Date).ToString("s")
    }

    foreach ($c in $list) {
        if ($c.name -ne $entry.name) { $c.active = $false }
    }

    Save-Connections $list
    return $entry
}

function Get-AppTenantId {
    $tenantId = $global:Config.auth.app.tenantId
    if ($tenantId) { return $tenantId }
    if ($global:Config.tenant.tenantId) { return $global:Config.tenant.tenantId }
    if ($global:Config.tenant.defaultDomain) { return $global:Config.tenant.defaultDomain }
    return "organizations"
}


function Get-AppClientId {
    return $global:Config.auth.app.clientId
}


function Get-AppClientSecret {
    return $global:Config.auth.app.clientSecret
}


function Get-AppToken {
    param(
        [string]$Scope,
        [switch]$ForceRefresh
    )
    if (-not $Scope) { return $null }
    if (-not (Ensure-MsalModule)) { return $null }

    $cacheKey = $Scope.ToLowerInvariant()
    if (-not $ForceRefresh -and $global:AppTokenCache.ContainsKey($cacheKey)) {
        $cached = $global:AppTokenCache[$cacheKey]
        if ($cached -and $cached.ExpiresOn -gt (Get-Date).AddMinutes(5)) {
            return $cached.AccessToken
        }
    }

    $clientId = Get-AppClientId
    $clientSecret = Get-AppClientSecret
    if (-not $clientId -or -not $clientSecret) {
        Write-Warn "App credentials missing. Set auth.app.clientId and auth.app.clientSecret in config."
        return $null
    }
    $tenantId = Get-AppTenantId
    $params = @{
        TenantId     = $tenantId
        ClientId     = $clientId
        ClientSecret = $clientSecret
        Scopes       = @($Scope)
    }
    try {
        $tok = Get-MsalToken @params
        if ($tok -and $tok.AccessToken) {
            $global:AppTokenCache[$cacheKey] = @{
                AccessToken = $tok.AccessToken
                ExpiresOn   = $tok.ExpiresOn
            }
            return $tok.AccessToken
        }
    } catch {
        Write-Err $_.Exception.Message
    }
    return $null
}


function Invoke-ExternalApiRequest {
    param(
        [string]$Method,
        [string]$Url,
        [object]$Body,
        [hashtable]$Headers,
        [string]$Scope,
        [string]$BaseUrl,
        [switch]$AllowNullResponse
    )
    $token = Get-AppToken -Scope $Scope
    if (-not $token) { return $null }
    $fullUrl = if ($Url -match "^https?://") {
        $Url
    } else {
        $base = if ($BaseUrl) { $BaseUrl.TrimEnd("/") } else { "" }
        if ($Url.StartsWith("/")) { $base + $Url } else { $base + "/" + $Url }
    }
    $hdr = @{ Authorization = ("Bearer " + $token) }
    if ($Headers) {
        foreach ($k in $Headers.Keys) { $hdr[$k] = $Headers[$k] }
    }
    $params = @{ Method = $Method; Uri = $fullUrl; Headers = $hdr }
    if ($Body -ne $null) {
        if ($Body -is [string]) {
            $params.Body = $Body
            $params.ContentType = "application/json"
        } else {
            $params.Body = ($Body | ConvertTo-Json -Depth 10)
            $params.ContentType = "application/json"
        }
    }
    try {
        return Invoke-RestMethod @params
    } catch {
        if (-not $AllowNullResponse) {
            Write-Err $_.Exception.Message
        }
        return $null
    }
}


