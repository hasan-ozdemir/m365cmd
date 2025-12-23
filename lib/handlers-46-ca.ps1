# Handler: Ca
# Purpose: Ca command handlers.
function Handle-CACommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: ca policy|location ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    if (-not $rest -or $rest.Count -eq 0) {
        Write-Warn "Usage: ca policy|location ..."
        return
    }
    $action = $rest[0].ToLowerInvariant()
    $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest2
    $apiInfo = Resolve-CAApiSettings $parsed

    $base = switch ($sub) {
        "policy" { "/identity/conditionalAccess/policies" }
        "policies" { "/identity/conditionalAccess/policies" }
        "location" { "/identity/conditionalAccess/namedLocations" }
        "locations" { "/identity/conditionalAccess/namedLocations" }
        "namedlocation" { "/identity/conditionalAccess/namedLocations" }
        "namedlocations" { "/identity/conditionalAccess/namedLocations" }
        default { $null }
    }
    if (-not $base) {
        Write-Warn "Usage: ca policy|location ..."
        return
    }

    switch ($action) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @()
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
            if ($resp -and $resp.value) {
                $resp.value | ConvertTo-Json -Depth 10
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: ca $sub get <id> [--beta|--auto]"
                return
            }
            $qh = Build-QueryAndHeaders $parsed.Map @()
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "create" {
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $body) {
                Write-Warn "Usage: ca $sub create --json <payload> [--bodyFile <file>] [--beta|--auto]"
                return
            }
            $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
            if ($resp) {
                Write-Info "Created."
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $id -or -not $body) {
                Write-Warn "Usage: ca $sub update <id> --json <payload> [--bodyFile <file>] [--beta|--auto]"
                return
            }
            $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
            if ($resp -ne $null) { Write-Info "Updated." }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: ca $sub delete <id> [--force]"
                return
            }
            $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
            if (-not $force) {
                $confirm = Read-Host "Type DELETE to confirm"
                if ($confirm -ne "DELETE") {
                    Write-Info "Canceled."
                    return
                }
            }
            $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri ($base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback -AllowNullResponse
            if ($resp -ne $null) { Write-Info "Deleted." }
        }
        default {
            Write-Warn "Usage: ca $sub list|get|create|update|delete"
        }
    }
}
