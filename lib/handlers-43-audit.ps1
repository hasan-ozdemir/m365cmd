# Handler: Audit
# Purpose: Audit command handlers.
function Resolve-AuditPath {
    param([string]$Type)
    $t = if ($Type) { $Type.ToLowerInvariant() } else { "directory" }
    switch ($t) {
        "directory" { return "/auditLogs/directoryAudits" }
        "signin" { return "/auditLogs/signIns" }
        "signins" { return "/auditLogs/signIns" }
        "provisioning" { return "/auditLogs/provisioning" }
        default { return "/auditLogs/directoryAudits" }
    }
}

function Handle-AuditCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: audit list|get [--type directory|signin|provisioning]"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    if ($sub -in @("alert", "alerts")) {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: security alert list|get|update [--legacy]"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed2 = Parse-NamedArgs $rest2
        $legacy = $parsed2.Map.ContainsKey("legacy")
        $apiInfo = Resolve-CAApiSettings $parsed2
        $base = if ($legacy) { "/security/alerts" } else { "/security/alerts_v2" }

        switch ($action) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed2.Map @()
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                if ($resp -and $resp.value) {
                    $resp.value | ConvertTo-Json -Depth 10
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 10
                }
            }
            "get" {
                $id = $parsed2.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: security alert get <id> [--legacy]"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "update" {
                $id = $parsed2.Positionals | Select-Object -First 1
                $body = Read-JsonPayload (Get-ArgValue $parsed2.Map "json") (Get-ArgValue $parsed2.Map "bodyFile") (Get-ArgValue $parsed2.Map "set")
                if (-not $id -or -not $body) {
                    Write-Warn "Usage: security alert update <id> --json <payload> OR --set key=value [--legacy]"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                if ($resp -ne $null) { Write-Info "Alert updated." }
            }
            default {
                Write-Warn "Usage: security alert list|get|update [--legacy]"
            }
        }
        return
    }

    if ($sub -in @("incident", "incidents")) {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: security incident list|get|update"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed2 = Parse-NamedArgs $rest2
        $apiInfo = Resolve-CAApiSettings $parsed2
        $base = "/security/incidents"

        switch ($action) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed2.Map @()
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                if ($resp -and $resp.value) {
                    $resp.value | ConvertTo-Json -Depth 10
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 10
                }
            }
            "get" {
                $id = $parsed2.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: security incident get <id>"
                    return
                }
                $qh = Build-QueryAndHeaders $parsed2.Map @()
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "update" {
                $id = $parsed2.Positionals | Select-Object -First 1
                $body = Read-JsonPayload (Get-ArgValue $parsed2.Map "json") (Get-ArgValue $parsed2.Map "bodyFile") (Get-ArgValue $parsed2.Map "set")
                if (-not $id -or -not $body) {
                    Write-Warn "Usage: security incident update <id> --json <payload> OR --set key=value"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                if ($resp -ne $null) { Write-Info "Incident updated." }
            }
            default {
                Write-Warn "Usage: security incident list|get|update"
            }
        }
        return
    }

    if ($sub -in @("hunt", "hunting")) {
        $parsed2 = Parse-NamedArgs $rest
        $query = Get-ArgValue $parsed2.Map "query"
        $queryFile = Get-ArgValue $parsed2.Map "queryFile"
        $apiInfo = Resolve-CAApiSettings $parsed2
        if (-not $query -and -not $queryFile) {
            Write-Warn "Usage: security hunt --query <kql> [--beta|--auto]"
            return
        }
        if ($queryFile) {
            if (-not (Test-Path $queryFile)) {
                Write-Warn "Query file not found."
                return
            }
            $query = Get-Content -Raw -Path $queryFile
        }
        $body = @{ query = $query }
        $resp = Invoke-GraphRequestAuto -Method "POST" -Uri "/security/runHuntingQuery" -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
        if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        return
    }

    if ($sub -in @("ti", "indicator", "indicators", "tiindicator", "tiindicators")) {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: security ti list|get|create|update|delete"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed2 = Parse-NamedArgs $rest2
        $apiInfo = Resolve-CAApiSettings $parsed2
        $base = "/security/tiIndicators"

        switch ($action) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed2.Map @()
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                if ($resp -and $resp.value) {
                    $resp.value | ConvertTo-Json -Depth 10
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 10
                }
            }
            "get" {
                $id = $parsed2.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: security ti get <id>"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "create" {
                $body = Read-JsonPayload (Get-ArgValue $parsed2.Map "json") (Get-ArgValue $parsed2.Map "bodyFile") $null
                if (-not $body) {
                    Write-Warn "Usage: security ti create --json <payload>"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "update" {
                $id = $parsed2.Positionals | Select-Object -First 1
                $body = Read-JsonPayload (Get-ArgValue $parsed2.Map "json") (Get-ArgValue $parsed2.Map "bodyFile") (Get-ArgValue $parsed2.Map "set")
                if (-not $id -or -not $body) {
                    Write-Warn "Usage: security ti update <id> --json <payload> OR --set key=value"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                if ($resp -ne $null) { Write-Info "TI indicator updated." }
            }
            "delete" {
                $id = $parsed2.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: security ti delete <id> [--force]"
                    return
                }
                $force = Parse-Bool (Get-ArgValue $parsed2.Map "force") $false
                if (-not $force) {
                    $confirm = Read-Host "Type DELETE to confirm"
                    if ($confirm -ne "DELETE") {
                        Write-Info "Canceled."
                        return
                    }
                }
                $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri ($base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback -AllowNullResponse
                if ($resp -ne $null) { Write-Info "TI indicator deleted." }
            }
            default {
                Write-Warn "Usage: security ti list|get|create|update|delete"
            }
        }
        return
    }

    $type = Get-ArgValue $parsed.Map "type"
    $path = Resolve-AuditPath $type

    switch ($sub) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @()
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($path + $qh.Query) -Headers $qh.Headers -AllowFallback:$true
            if ($resp -and $resp.value) {
                $resp.value | ConvertTo-Json -Depth 10
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: audit get <id> [--type directory|signin|provisioning]"
                return
            }
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($path + "/" + $id) -AllowFallback:$true
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        default {
            Write-Warn "Usage: audit list|get [--type directory|signin|provisioning]"
        }
    }
}

