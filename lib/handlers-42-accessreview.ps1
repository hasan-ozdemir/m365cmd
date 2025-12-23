# Handler: Accessreview
# Purpose: Accessreview command handlers.
function Handle-AccessReviewCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: accessreview list|get|create|update|delete|instance|decision|history"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $base = "/identityGovernance/accessReviews/definitions"

    switch ($sub) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @()
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query)
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: accessreview get <id>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "create" {
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") $null
            if (-not $body) {
                Write-Warn "Usage: accessreview create --json <payload>"
                return
            }
            $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $id -or -not $body) {
                Write-Warn "Usage: accessreview update <id> --json <payload> OR --set key=value"
                return
            }
            $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
            if ($resp -ne $null) { Write-Info "Access review updated." }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: accessreview delete <id>"
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
            $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($base + "/" + $id)
            if ($resp -ne $null) { Write-Info "Access review deleted." }
        }
        "instance" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: accessreview instance list --def <definitionId>"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $parsed2 = Parse-NamedArgs ($rest | Select-Object -Skip 1)
            $defId = Get-ArgValue $parsed2.Map "def"
            if (-not $defId) {
                Write-Warn "Usage: accessreview instance list --def <definitionId>"
                return
            }
            if ($action -ne "list") {
                Write-Warn "Usage: accessreview instance list --def <definitionId>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $defId + "/instances")
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "decision" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: accessreview decision list|submit|apply --def <definitionId> --instance <instanceId>"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $parsed2 = Parse-NamedArgs ($rest | Select-Object -Skip 1)
            $defId = Get-ArgValue $parsed2.Map "def"
            $instId = Get-ArgValue $parsed2.Map "instance"
            if (-not $defId -or -not $instId) {
                Write-Warn "Usage: accessreview decision list|submit|apply --def <definitionId> --instance <instanceId>"
                return
            }
            $decBase = $base + "/" + $defId + "/instances/" + $instId + "/decisions"
            switch ($action) {
                "list" {
                    $resp = Invoke-GraphRequest -Method "GET" -Uri $decBase
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "submit" {
                    $decisionId = $parsed2.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed2.Map "json") (Get-ArgValue $parsed2.Map "bodyFile") (Get-ArgValue $parsed2.Map "set")
                    if (-not $decisionId -or -not $body) {
                        Write-Warn "Usage: accessreview decision submit <decisionId> --def <definitionId> --instance <instanceId> --json <payload> OR --set key=value"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "POST" -Uri ($decBase + "/" + $decisionId) -Body $body
                    if ($resp -ne $null) { Write-Info "Decision submitted." }
                }
                "apply" {
                    $resp = Invoke-GraphRequest -Method "POST" -Uri ($base + "/" + $defId + "/instances/" + $instId + "/applyDecisions")
                    if ($resp -ne $null) { Write-Info "Decisions applied." }
                }
                default {
                    Write-Warn "Usage: accessreview decision list|submit|apply --def <definitionId> --instance <instanceId>"
                }
            }
        }
        "history" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: accessreview history list|get|create|delete|instance"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $parsed2 = Parse-NamedArgs ($rest | Select-Object -Skip 1)
            $base2 = "/identityGovernance/accessReviews/historyDefinitions"
            switch ($action) {
                "list" {
                    $resp = Invoke-GraphRequest -Method "GET" -Uri $base2
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "get" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: accessreview history get <id>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ($base2 + "/" + $id)
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "create" {
                    $body = Read-JsonPayload (Get-ArgValue $parsed2.Map "json") (Get-ArgValue $parsed2.Map "bodyFile") $null
                    if (-not $body) {
                        Write-Warn "Usage: accessreview history create --json <payload>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "POST" -Uri $base2 -Body $body
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "delete" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: accessreview history delete <id>"
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
                    $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($base2 + "/" + $id)
                    if ($resp -ne $null) { Write-Info "History definition deleted." }
                }
                "instance" {
                    $id = Get-ArgValue $parsed2.Map "id"
                    if (-not $id) {
                        Write-Warn "Usage: accessreview history instance list --id <historyDefinitionId>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ($base2 + "/" + $id + "/instances")
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                default {
                    Write-Warn "Usage: accessreview history list|get|create|delete|instance"
                }
            }
        }
        default {
            Write-Warn "Usage: accessreview list|get|create|update|delete|instance|decision|history"
        }
    }
}
