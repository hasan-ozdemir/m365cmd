# Core: Graph Org
# Purpose: Graph Org shared utilities.
function Get-OrganizationDefault {
    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri "/organization" -AllowFallback
    if (-not $resp) { return $null }
    if ($resp.value) { return ($resp.value | Select-Object -First 1) }
    return $resp
}

function Resolve-TenantGuid {
    if ($global:Config.tenant.tenantId) { return $global:Config.tenant.tenantId }
    $ctx = Get-MgContextSafe
    if ($ctx -and $ctx.TenantId) { return $ctx.TenantId }
    if (Get-Command -Name Get-MgOrganization -ErrorAction SilentlyContinue) {
        try {
            $org = Get-OrganizationDefault
            if ($org -and $org.id) {
                Set-ConfigValue "tenant.tenantId" $org.id
                return $org.id
            }
        } catch {}
    }
    return $null
}
