# Core: Graph Context
# Purpose: Graph Context shared utilities.
function Get-GraphBaseUri {
    param([switch]$Beta)
    if ($Beta) { return "https://graph.microsoft.com/beta" }
    return "https://graph.microsoft.com/v1.0"
}

function Get-MgContextSafe {
    if (Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue) {
        return Get-MgContext
    }
    return $null
}

function Require-GraphConnection {
    if (-not (Ensure-GraphModule)) { return $false }
    $ctx = Get-MgContextSafe
    if (-not $ctx -or -not $ctx.Account) {
        Write-Warn "Not connected. Use /login."
        return $false
    }
    return $true
}
