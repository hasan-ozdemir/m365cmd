# Core: Tokens
# Purpose: Shared token helpers for non-Graph resources.
$global:ResourceTokens = @{}

function Get-TokenCacheKey {
    param([string]$Scope)
    if (-not $Scope) { return "default" }
    $key = $Scope.ToLowerInvariant()
    $key = $key -replace "[^a-z0-9]+", "-"
    if ($key.Length -gt 80) { $key = $key.Substring(0, 80) }
    return $key
}


function Get-TokenCachePath {
    param([string]$Key)
    $dir = Join-Path $Paths.Data "tokens"
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    return (Join-Path $dir ($Key + ".json"))
}


function Load-TokenCache {
    param([string]$Key)
    $path = Get-TokenCachePath $Key
    if (-not (Test-Path $path)) { return $null }
    try {
        return (Get-Content -Raw -Path $path | ConvertFrom-Json)
    } catch {
        Write-Warn "Token cache is invalid. Reauth required."
        return $null
    }
}


function Save-TokenCache {
    param(
        [string]$Key,
        [string]$AccessToken,
        [datetime]$ExpiresOn
    )
    $obj = [ordered]@{
        accessToken = $AccessToken
        expiresOn   = $ExpiresOn.ToString("o")
    }
    $obj | ConvertTo-Json -Depth 3 | Set-Content -Path (Get-TokenCachePath $Key) -Encoding ASCII
}


function Get-AuthPublicClientId {
    if ($global:Config.auth.publicClientId) { return $global:Config.auth.publicClientId }
    if ($global:Config.pp.clientId) { return $global:Config.pp.clientId }
    return "04f0c124-f2bc-4f4b-8c20-74bf17c5a1f9"
}


function Get-DelegatedToken {
    param(
        [string]$Scope,
        [switch]$ForceLogin,
        [switch]$DeviceCode
    )
    if (-not $Scope) { return $null }
    if (-not (Ensure-MsalModule)) { return $null }

    $key = Get-TokenCacheKey $Scope
    $cache = Load-TokenCache $key
    if (-not $ForceLogin -and $cache -and $cache.accessToken -and $cache.expiresOn) {
        try {
            $exp = [datetime]::Parse($cache.expiresOn)
            if ($exp -gt (Get-Date).AddMinutes(5)) {
                return $cache.accessToken
            }
        } catch {}
    }

    $tenantId = Get-PpTenantId
    $clientId = Get-AuthPublicClientId
    $params = @{
        TenantId = $tenantId
        ClientId = $clientId
        Scopes   = @($Scope)
    }
    $cmd = Get-Command -Name Get-MsalToken -ErrorAction SilentlyContinue
    $supportsSilent = $cmd -and $cmd.Parameters.ContainsKey("Silent")
    $supportsDevice = $cmd -and $cmd.Parameters.ContainsKey("DeviceCode")

    if ($DeviceCode -and $supportsDevice) { $params.DeviceCode = $true }

    try {
        if (-not $ForceLogin -and $supportsSilent) { $params.Silent = $true }
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
        $exp = if ($tok.ExpiresOn) { $tok.ExpiresOn } else { (Get-Date).AddHours(1) }
        Save-TokenCache $key $tok.AccessToken $exp
        return $tok.AccessToken
    }
    return $null
}


function Decode-Jwt {
    param([string]$Token)
    if (-not $Token -or ($Token -notmatch "\.")) { return $null }
    $parts = $Token -split "\."
    if ($parts.Count -lt 2) { return $null }
    $pad = { param($s) while ($s.Length % 4 -ne 0) { $s += "=" }; return $s.Replace('-', '+').Replace('_', '/') }
    try {
        $headerJson = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String((&$pad $parts[0])))
        $payloadJson = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String((&$pad $parts[1])))
        return [pscustomobject]@{
            header  = (ConvertFrom-Json $headerJson)
            payload = (ConvertFrom-Json $payloadJson)
        }
    } catch {
        return $null
    }
}
