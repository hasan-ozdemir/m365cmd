# Core: Pp
# Purpose: Pp shared utilities.
$global:PpToken = $null
$global:PpTokenExpires = [datetime]::MinValue

function Get-PpTokenCachePath {
    return (Join-Path (Join-Path $Paths.Data "pp") "token.json")
}



function Load-PpTokenCache {
    $path = Get-PpTokenCachePath
    if (-not (Test-Path $path)) { return }
    try {
        $obj = Get-Content -Raw -Path $path | ConvertFrom-Json
        if ($obj -and $obj.accessToken) {
            $global:PpToken = $obj.accessToken
            if ($obj.expiresOn) {
                $global:PpTokenExpires = [datetime]::Parse($obj.expiresOn)
            }
        }
    } catch {
        Write-Warn "Power Platform token cache is invalid. Recreating on next login."
        Remove-Item -Path $path -Force -ErrorAction SilentlyContinue
    }
}



function Save-PpTokenCache {
    $path = Get-PpTokenCachePath
    $obj = [ordered]@{
        accessToken = $global:PpToken
        expiresOn   = $global:PpTokenExpires.ToString("o")
    }
    $obj | ConvertTo-Json -Depth 3 | Set-Content -Path $path -Encoding ASCII
}



function Clear-PpTokenCache {
    $global:PpToken = $null
    $global:PpTokenExpires = [datetime]::MinValue
    $path = Get-PpTokenCachePath
    if (Test-Path $path) {
        Remove-Item -Path $path -Force -ErrorAction SilentlyContinue
    }
}



function Get-PpTenantId {
    $tenantId = $global:Config.tenant.tenantId
    if ($tenantId) { return $tenantId }
    $domain = $global:Config.tenant.defaultDomain
    if ($domain) { return $domain }
    return "organizations"
}



function Get-PpToken {
    param(
        [switch]$ForceLogin,
        [switch]$DeviceCode
    )
    if (-not $global:PpToken) { Load-PpTokenCache }
    if (-not $ForceLogin -and $global:PpToken -and $global:PpTokenExpires -gt (Get-Date).AddMinutes(5)) {
        return $global:PpToken
    }
    if (-not (Ensure-MsalModule)) { return $null }

    $tenantId = Get-PpTenantId
    $clientId = $global:Config.pp.clientId
    $scopes = @("https://api.powerplatform.com/.default")
    $params = @{
        TenantId = $tenantId
        ClientId = $clientId
        Scopes   = $scopes
    }

    $cmd = Get-Command -Name Get-MsalToken -ErrorAction SilentlyContinue
    $supportsSilent = $cmd -and $cmd.Parameters.ContainsKey("Silent")
    $supportsDevice = $cmd -and $cmd.Parameters.ContainsKey("DeviceCode")

    if ($DeviceCode -and $supportsDevice) {
        $params.DeviceCode = $true
    }

    try {
        if (-not $ForceLogin -and $supportsSilent) {
            $params.Silent = $true
        }
        $tok = Get-MsalToken @params
    } catch {
        try {
            if ($params.ContainsKey("Silent")) { $params.Remove("Silent") }
            $tok = Get-MsalToken @params
        } catch {
            Write-Err $_.Exception.Message
            return $null
        }
    }

    if ($tok -and $tok.AccessToken) {
        $global:PpToken = $tok.AccessToken
        if ($tok.ExpiresOn) {
            $global:PpTokenExpires = $tok.ExpiresOn
        } else {
            $global:PpTokenExpires = (Get-Date).AddHours(1)
        }
        Save-PpTokenCache
        return $global:PpToken
    }
    return $null
}



function Invoke-PpRequest {
    param(
        [string]$Method,
        [string]$Path,
        [object]$Body,
        [hashtable]$Headers,
        [string]$ApiVersion,
        [switch]$AllowNullResponse
    )
    $token = Get-PpToken
    if (-not $token) { return $null }
    $base = $global:Config.pp.baseUrl.TrimEnd("/")
    $url = if ($Path -match "^https?://") { $Path } else { $base + $Path }
    $apiVer = if ($ApiVersion) { $ApiVersion } else { $global:Config.pp.apiVersion }
    if ($url -notmatch "api-version=") {
        $join = if ($url -match "\\?") { "&" } else { "?" }
        $url = $url + $join + "api-version=" + $apiVer
    }
    $hdr = @{ Authorization = "Bearer " + $token }
    if ($Headers) {
        foreach ($k in $Headers.Keys) { $hdr[$k] = $Headers[$k] }
    }
    $params = @{ Method = $Method; Uri = $url; Headers = $hdr }
    if ($Body -ne $null) {
        $params.ContentType = "application/json"
        if ($Body -is [string]) {
            $params.Body = $Body
        } else {
            $params.Body = ($Body | ConvertTo-Json -Depth 10)
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



