# Handler: Adminportal
# Purpose: Adminportal command handlers.
function Handle-AdminPortalCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: admin open|user|license|role|group|domain|org|billing|health|message|security|purview|compliance"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    switch ($sub) {
        "open" { Write-Host "https://admin.microsoft.com" }
        "user" { Handle-UserCommand $rest }
        "license" { Handle-LicenseCommand $rest }
        "role" { Handle-RoleCommand $rest }
        "group" { Handle-GroupCommand $rest }
        "domain" { Handle-DomainCommand $rest }
        "org" { Handle-OrgCommand $rest }
        "billing" { Handle-BillingCommand $rest }
        "health" { Handle-HealthCommand $rest }
        "message" { Handle-MessageCommand $rest }
        "security" { Handle-SecurityCommand $rest }
        "purview" { Handle-PurviewCommand $rest }
        "compliance" { Handle-ComplianceCommand $rest }
        default {
            Write-Warn "Usage: admin open|user|license|role|group|domain|org|billing|health|message|security|purview|compliance"
        }
    }
}

