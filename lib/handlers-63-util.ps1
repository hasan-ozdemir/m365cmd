# Handler: Util
# Purpose: Utility helpers (access tokens).
function Resolve-ResourceScope {
    param([string]$Resource)
    if (-not $Resource) { return $null }
    $r = $Resource.ToLowerInvariant()
    if ($r -eq "graph") {
        return "https://graph.microsoft.com/.default"
    }
    if ($r -eq "sharepoint") {
        $prefix = $global:Config.tenant.defaultPrefix
        if (-not $prefix) { return $null }
        return ("https://" + $prefix + ".sharepoint.com/.default")
    }
    if ($Resource -match "^https?://") {
        if ($Resource.EndsWith("/.default")) { return $Resource }
        return ($Resource.TrimEnd("/") + "/.default")
    }
    return $Resource
}


function Handle-UtilCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: util accesstoken get --resource <graph|sharepoint|url> [--new] [--decoded]"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    if ($sub -eq "accesstoken") {
        $action = if ($rest.Count -gt 0) { $rest[0].ToLowerInvariant() } else { "" }
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed2 = Parse-NamedArgs $rest2
        if ($action -ne "get") {
            Write-Warn "Usage: util accesstoken get --resource <graph|sharepoint|url> [--new] [--decoded]"
            return
        }
        $resource = Get-ArgValue $parsed2.Map "resource"
        if (-not $resource) { $resource = $parsed2.Positionals | Select-Object -First 1 }
        if (-not $resource) {
            Write-Warn "Usage: util accesstoken get --resource <graph|sharepoint|url> [--new] [--decoded]"
            return
        }
        $scope = Resolve-ResourceScope $resource
        if (-not $scope) {
            Write-Warn "Unable to resolve resource scope."
            return
        }
        $force = Parse-Bool (Get-ArgValue $parsed2.Map "new") $false
        $decoded = Parse-Bool (Get-ArgValue $parsed2.Map "decoded") $false

        $token = Get-DelegatedToken -Scope $scope -ForceLogin:$force
        if (-not $token) {
            $token = Get-AppToken -Scope $scope -ForceRefresh:$force
        }
        if (-not $token) {
            Write-Warn "Token acquisition failed. Configure auth.app.* or login with delegated access."
            return
        }
        if ($decoded) {
            $parts = Decode-Jwt $token
            if ($parts) {
                Write-Host ($parts.header | ConvertTo-Json -Depth 5)
                Write-Host ($parts.payload | ConvertTo-Json -Depth 8)
            } else {
                Write-Warn "Unable to decode token."
            }
        } else {
            Write-Host $token
        }
        return
    }

    Write-Warn "Usage: util accesstoken get --resource <graph|sharepoint|url> [--new] [--decoded]"
}
