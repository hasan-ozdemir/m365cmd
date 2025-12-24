# Handler: Purview
# Purpose: Purview command handlers.
function Handle-PurviewCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: purview ediscovery|srr ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    if (-not $rest -or $rest.Count -eq 0) {
        Write-Warn "Usage: purview ediscovery|srr ..."
        return
    }
    $action = $rest[0].ToLowerInvariant()
    $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest2
    $apiInfo = Resolve-CAApiSettings $parsed

    switch ($sub) {
        "ediscovery" {
            function Invoke-EdiscoveryCaseCrud {
                param([string]$Action, [object]$Parsed)
                $apiInfo = Resolve-CAApiSettings $Parsed
                $base = "/security/cases/ediscoveryCases"
                switch ($Action) {
                    "list" {
                        $qh = Build-QueryAndHeaders $Parsed.Map @("id", "displayName", "status")
                        $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                        if ($resp -and $resp.value) {
                            Write-GraphTable $resp.value @("Id", "DisplayName", "Status")
                        } elseif ($resp) {
                            $resp | ConvertTo-Json -Depth 10
                        }
                    }
                    "get" {
                        $id = $Parsed.Positionals | Select-Object -First 1
                        if (-not $id) {
                            Write-Warn "Usage: purview ediscovery case get <caseId>"
                            return
                        }
                        $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                        if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                    }
                    "create" {
                        $body = Read-JsonPayload (Get-ArgValue $Parsed.Map "json") (Get-ArgValue $Parsed.Map "bodyFile") $null
                        if (-not $body) {
                            Write-Warn "Usage: purview ediscovery case create --json <payload>"
                            return
                        }
                        $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                        if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                    }
                    "update" {
                        $id = $Parsed.Positionals | Select-Object -First 1
                        $body = Read-JsonPayload (Get-ArgValue $Parsed.Map "json") (Get-ArgValue $Parsed.Map "bodyFile") (Get-ArgValue $Parsed.Map "set")
                        if (-not $id -or -not $body) {
                            Write-Warn "Usage: purview ediscovery case update <caseId> --json <payload> OR --set key=value"
                            return
                        }
                        $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                        if ($resp -ne $null) { Write-Info "Case updated." }
                    }
                    "delete" {
                        $id = $Parsed.Positionals | Select-Object -First 1
                        if (-not $id) {
                            Write-Warn "Usage: purview ediscovery case delete <caseId> [--force]"
                            return
                        }
                        $force = Parse-Bool (Get-ArgValue $Parsed.Map "force") $false
                        if (-not $force) {
                            $confirm = Read-Host "Type DELETE to confirm"
                            if ($confirm -ne "DELETE") {
                                Write-Info "Canceled."
                                return
                            }
                        }
                        $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri ($base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback -AllowNullResponse
                        if ($resp -ne $null) { Write-Info "Case deleted." }
                    }
                    default {
                        Write-Warn "Usage: purview ediscovery case list|get|create|update|delete"
                    }
                }
            }

            function Invoke-EdiscoveryCaseSubCrud {
                param([string]$Kind, [string]$Action, [object]$Parsed, [string]$Segment)
                $caseId = Get-ArgValue $Parsed.Map "case"
                if (-not $caseId) {
                    Write-Warn ("Usage: purview ediscovery " + $Kind + " " + $Action + " --case <caseId>")
                    return
                }
                $apiInfo = Resolve-CAApiSettings $Parsed
                $base = "/security/cases/ediscoveryCases/" + $caseId + "/" + $Segment
                switch ($Action) {
                    "list" {
                        $qh = Build-QueryAndHeaders $Parsed.Map @()
                        $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                        if ($resp -and $resp.value) {
                            $resp.value | ConvertTo-Json -Depth 10
                        } elseif ($resp) {
                            $resp | ConvertTo-Json -Depth 10
                        }
                    }
                    "get" {
                        $id = $Parsed.Positionals | Select-Object -First 1
                        if (-not $id) {
                            Write-Warn ("Usage: purview ediscovery " + $Kind + " get <id> --case <caseId>")
                            return
                        }
                        $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                        if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                    }
                    "create" {
                        $body = Read-JsonPayload (Get-ArgValue $Parsed.Map "json") (Get-ArgValue $Parsed.Map "bodyFile") $null
                        if (-not $body) {
                            Write-Warn ("Usage: purview ediscovery " + $Kind + " create --case <caseId> --json <payload>")
                            return
                        }
                        $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                        if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                    }
                    "update" {
                        $id = $Parsed.Positionals | Select-Object -First 1
                        $body = Read-JsonPayload (Get-ArgValue $Parsed.Map "json") (Get-ArgValue $Parsed.Map "bodyFile") (Get-ArgValue $Parsed.Map "set")
                        if (-not $id -or -not $body) {
                            Write-Warn ("Usage: purview ediscovery " + $Kind + " update <id> --case <caseId> --json <payload> OR --set key=value")
                            return
                        }
                        $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                        if ($resp -ne $null) { Write-Info ($Kind + " updated.") }
                    }
                    "delete" {
                        $id = $Parsed.Positionals | Select-Object -First 1
                        if (-not $id) {
                            Write-Warn ("Usage: purview ediscovery " + $Kind + " delete <id> --case <caseId> [--force]")
                            return
                        }
                        $force = Parse-Bool (Get-ArgValue $Parsed.Map "force") $false
                        if (-not $force) {
                            $confirm = Read-Host "Type DELETE to confirm"
                            if ($confirm -ne "DELETE") {
                                Write-Info "Canceled."
                                return
                            }
                        }
                        $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri ($base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback -AllowNullResponse
                        if ($resp -ne $null) { Write-Info ($Kind + " deleted.") }
                    }
                    default {
                        Write-Warn ("Usage: purview ediscovery " + $Kind + " list|get|create|update|delete --case <caseId>")
                    }
                }
            }

            $kind = $action
            $kindArgs = $rest2
            if ($action -in @("case","custodian","datasource","hold","search","reviewset")) {
                if (-not $rest2 -or $rest2.Count -eq 0) {
                    Write-Warn ("Usage: purview ediscovery " + $action + " list|get|create|update|delete")
                    return
                }
                $action = $rest2[0].ToLowerInvariant()
                $kindArgs = if ($rest2.Count -gt 1) { $rest2[1..($rest2.Count - 1)] } else { @() }
            } else {
                $kind = "case"
            }
            $parsedKind = Parse-NamedArgs $kindArgs

            switch ($kind) {
                "case" { Invoke-EdiscoveryCaseCrud -Action $action -Parsed $parsedKind }
                "custodian" { Invoke-EdiscoveryCaseSubCrud -Kind "custodian" -Action $action -Parsed $parsedKind -Segment "custodians" }
                "datasource" { Invoke-EdiscoveryCaseSubCrud -Kind "datasource" -Action $action -Parsed $parsedKind -Segment "noncustodialDataSources" }
                "hold" { Invoke-EdiscoveryCaseSubCrud -Kind "hold" -Action $action -Parsed $parsedKind -Segment "legalHolds" }
                "search" { Invoke-EdiscoveryCaseSubCrud -Kind "search" -Action $action -Parsed $parsedKind -Segment "searches" }
                "reviewset" { Invoke-EdiscoveryCaseSubCrud -Kind "reviewset" -Action $action -Parsed $parsedKind -Segment "reviewSets" }
                default { Write-Warn "Usage: purview ediscovery case|custodian|datasource|hold|search|reviewset ..." }
            }
        }
        "srr" { 
            $base = "/security/subjectRightsRequests"
            switch ($action) {
                "list" {
                    $qh = Build-QueryAndHeaders $parsed.Map @("id", "subjectType", "status", "createdDateTime")
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp -and $resp.value) {
                        Write-GraphTable $resp.value @("Id", "SubjectType", "Status", "CreatedDateTime")
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 10
                    }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: purview srr get <requestId>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "create" {
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") $null
                    if (-not $body) {
                        Write-Warn "Usage: purview srr create --json <payload>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: purview srr update <requestId> --json <payload> OR --set key=value"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp -ne $null) { Write-Info "SRR updated." }
                }
                default {
                    Write-Warn "Usage: purview srr list|get|create|update"
                }
            }
        }
        "retention" {
            $labelAction = $action
            $labelArgs = $rest2
            if ($action -eq "label" -or $action -eq "labels") {
                if (-not $rest2 -or $rest2.Count -eq 0) {
                    Write-Warn "Usage: purview retention label list|get"
                    return
                }
                $labelAction = $rest2[0].ToLowerInvariant()
                $labelArgs = if ($rest2.Count -gt 1) { $rest2[1..($rest2.Count - 1)] } else { @() }
            }
            $parsedLabel = Parse-NamedArgs $labelArgs
            $apiInfo = Resolve-CAApiSettings $parsedLabel
            $base = "/security/labels/retentionLabels"

            switch ($labelAction) {
                "list" {
                    $qh = Build-QueryAndHeaders $parsedLabel.Map @("id", "displayName", "behaviorDuringRetentionPeriod")
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp -and $resp.value) {
                        Write-GraphTable $resp.value @("Id", "DisplayName", "BehaviorDuringRetentionPeriod")
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 10
                    }
                }
                "get" {
                    $id = $parsedLabel.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: purview retention label get <labelId>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                default {
                    Write-Warn "Usage: purview retention label list|get"
                }
            }
        }
        "dlp" {
            $policyAction = $action
            $policyArgs = $rest2
            if ($action -eq "policy" -or $action -eq "policies") {
                if (-not $rest2 -or $rest2.Count -eq 0) {
                    Write-Warn "Usage: purview dlp policy list|get|create|update|delete"
                    return
                }
                $policyAction = $rest2[0].ToLowerInvariant()
                $policyArgs = if ($rest2.Count -gt 1) { $rest2[1..($rest2.Count - 1)] } else { @() }
            }
            $parsedPolicy = Parse-NamedArgs $policyArgs
            $apiInfo = Resolve-CAApiSettings $parsedPolicy
            $base = "/security/dataLossPreventionPolicies"

            switch ($policyAction) {
                "list" {
                    $qh = Build-QueryAndHeaders $parsedPolicy.Map @("id", "displayName", "mode", "isEnabled")
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp -and $resp.value) {
                        Write-GraphTable $resp.value @("Id", "DisplayName", "Mode", "IsEnabled")
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 10
                    }
                }
                "get" {
                    $id = $parsedPolicy.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: purview dlp policy get <policyId>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "create" {
                    $body = Read-JsonPayload (Get-ArgValue $parsedPolicy.Map "json") (Get-ArgValue $parsedPolicy.Map "bodyFile") $null
                    if (-not $body) {
                        Write-Warn "Usage: purview dlp policy create --json <payload>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "update" {
                    $id = $parsedPolicy.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsedPolicy.Map "json") (Get-ArgValue $parsedPolicy.Map "bodyFile") (Get-ArgValue $parsedPolicy.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: purview dlp policy update <policyId> --json <payload> OR --set key=value"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp -ne $null) { Write-Info "DLP policy updated." }
                }
                "delete" {
                    $id = $parsedPolicy.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: purview dlp policy delete <policyId> [--force]"
                        return
                    }
                    $force = Parse-Bool (Get-ArgValue $parsedPolicy.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host "Type DELETE to confirm"
                        if ($confirm -ne "DELETE") {
                            Write-Info "Canceled."
                            return
                        }
                    }
                    $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri ($base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback -AllowNullResponse
                    if ($resp -ne $null) { Write-Info "DLP policy deleted." }
                }
                default {
                    Write-Warn "Usage: purview dlp policy list|get|create|update|delete"
                }
            }
        }
        default {
            Write-Warn "Usage: purview ediscovery|srr|retention|dlp ..."
        }
    }
}

