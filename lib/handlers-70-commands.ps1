# Handler: Commands
# Purpose: Top-level CLI-like commands.
function Handle-DocsCommand {
    $url = "https://pnp.github.io/cli-microsoft365/"
    Write-Host ("Docs: " + $url)
    try { Start-Process $url | Out-Null } catch {}
}


function Handle-RequestCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -lt 2) {
        Write-Warn "Usage: request <get|post|patch|put|delete> <url|path> [--body <json>] [--bodyFile <path>] [--headers <json>] [--beta|--v1|--auto] [--out <file>]"
        return
    }
    Handle-GraphCommand (@("req") + $InputArgs)
}


function Handle-VersionCommand {
    $ver = "dev"
    $git = Get-Command -Name "git" -ErrorAction SilentlyContinue
    if ($git) {
        try {
            $hash = & $git.Source "rev-parse" "--short" "HEAD" 2>$null
            if ($hash) { $ver = $hash.Trim() }
        } catch {}
    }
    Write-Host ("m365cmd version: " + $ver)
}


function Handle-SetupCommand {
    param([string[]]$InputArgs)
    $parsed = Parse-NamedArgs $InputArgs
    $interactive = Parse-Bool (Get-ArgValue $parsed.Map "interactive") $false
    $scripting = Parse-Bool (Get-ArgValue $parsed.Map "scripting") $false
    $skipApp = Parse-Bool (Get-ArgValue $parsed.Map "skipApp") $false

    if ($interactive -and $scripting) {
        Write-Warn "Specify either --interactive or --scripting, not both."
        return
    }

    if (-not $skipApp) {
        $clientId = Get-ArgValue $parsed.Map "clientId"
        $tenantId = Get-ArgValue $parsed.Map "tenantId"
        $clientSecret = Get-ArgValue $parsed.Map "clientSecret"
        if (-not $clientId) { $clientId = Read-Host "Client ID" }
        if (-not $tenantId) { $tenantId = Read-Host "Tenant ID (or domain)" }
        if (-not $clientSecret) { $clientSecret = Read-Host "Client Secret (leave blank if none)" }
        if ($clientId) { Set-ConfigValue "auth.app.clientId" $clientId }
        if ($tenantId) { Set-ConfigValue "auth.app.tenantId" $tenantId }
        if ($clientSecret) { Set-ConfigValue "auth.app.clientSecret" $clientSecret }
    }

    if ($interactive) {
        Set-ConfigValue "graph.defaultApi" "v1"
        Set-ConfigValue "graph.fallbackToBeta" $true
    }
    if ($scripting) {
        Set-ConfigValue "graph.defaultApi" "v1"
        Set-ConfigValue "graph.fallbackToBeta" $false
    }
    Write-Info "Setup completed."
}
