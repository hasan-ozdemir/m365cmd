# Core: Graph Requests
# Purpose: Graph Requests shared utilities.
function Invoke-GraphRequest {
    param(
        [string]$Method,
        [string]$Uri,
        [object]$Body,
        [hashtable]$Headers,
        [string]$ContentType,
        [switch]$Beta,
        [switch]$SuppressError,
        [ref]$ErrorRef
    )
    if (-not (Require-GraphConnection)) { return $null }
    $base = Get-GraphBaseUri -Beta:$Beta
    $url = if ($Uri -match "^https?://") {
        $Uri
    } else {
        if ($Uri.StartsWith("/")) { $base + $Uri } else { $base + "/" + $Uri }
    }
    $params = @{ Method = $Method; Uri = $url }
    if ($Headers -and $Headers.Count -gt 0) { $params.Headers = $Headers }
    if ($Body -ne $null) { $params.Body = $Body }
    if ($ContentType) { $params.ContentType = $ContentType }
    try {
        return Invoke-MgGraphRequest @params
    } catch {
        if ($ErrorRef) { $ErrorRef.Value = $_ }
        if (-not $SuppressError) {
            Write-Err $_.Exception.Message
        }
        return $null
    }
}

function Invoke-GraphRequestAuto {
    param(
        [string]$Method,
        [string]$Uri,
        [object]$Body,
        [hashtable]$Headers,
        [string]$ContentType,
        [string]$Api,
        [switch]$AllowFallback,
        [switch]$AllowNullResponse
    )
    $apiMode = if ($Api) { $Api.ToLowerInvariant() } else { $global:Config.graph.defaultApi }
    $useBeta = $apiMode -eq "beta"
    $useV1 = $apiMode -eq "v1"
    if ($useBeta) {
        return Invoke-GraphRequest -Method $Method -Uri $Uri -Body $Body -Headers $Headers -ContentType $ContentType -Beta
    }
    if ($useV1) {
        return Invoke-GraphRequest -Method $Method -Uri $Uri -Body $Body -Headers $Headers -ContentType $ContentType
    }

    $err = $null
    $resp = Invoke-GraphRequest -Method $Method -Uri $Uri -Body $Body -Headers $Headers -ContentType $ContentType -SuppressError -ErrorRef ([ref]$err)
    if ($resp -ne $null -or $AllowNullResponse) {
        return $resp
    }
    if ($AllowFallback) {
        $err2 = $null
        $resp2 = Invoke-GraphRequest -Method $Method -Uri $Uri -Body $Body -Headers $Headers -ContentType $ContentType -Beta -SuppressError -ErrorRef ([ref]$err2)
        if ($resp2 -ne $null -or $AllowNullResponse) {
            Write-Warn "v1 call failed; retried on beta."
            return $resp2
        }
        if ($err2) { Write-Err $err2.Exception.Message }
    }
    if ($err) { Write-Err $err.Exception.Message }
    return $null
}

function Invoke-GraphDownload {
    param(
        [string]$Uri,
        [string]$OutFile,
        [hashtable]$Headers,
        [switch]$Beta
    )
    if (-not $OutFile) {
        Write-Warn "Output file required."
        return
    }
    if (-not (Require-GraphConnection)) { return }
    $base = Get-GraphBaseUri -Beta:$Beta
    $url = if ($Uri -match "^https?://") {
        $Uri
    } else {
        if ($Uri.StartsWith("/")) { $base + $Uri } else { $base + "/" + $Uri }
    }

    $cmd = Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Parameters.ContainsKey("OutputFilePath")) {
        $params = @{ Method = "GET"; Uri = $url; OutputFilePath = $OutFile }
        if ($Headers -and $Headers.Count -gt 0) { $params.Headers = $Headers }
        try {
            Invoke-MgGraphRequest @params | Out-Null
            Write-Info ("Saved: " + $OutFile)
        } catch {
            Write-Err $_.Exception.Message
        }
        return
    }

    $resp = Invoke-GraphRequest -Method "GET" -Uri $url -Headers $Headers -Beta:$Beta
    if ($resp -is [byte[]]) {
        [System.IO.File]::WriteAllBytes($OutFile, $resp)
        Write-Info ("Saved: " + $OutFile)
        return
    }
    if ($resp -is [string]) {
        Set-Content -Path $OutFile -Value $resp -Encoding ASCII
        Write-Info ("Saved: " + $OutFile)
        return
    }
    if ($resp) {
        $resp | ConvertTo-Json -Depth 8 | Set-Content -Path $OutFile -Encoding ASCII
        Write-Info ("Saved: " + $OutFile)
    }
}
