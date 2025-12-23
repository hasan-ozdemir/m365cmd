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


