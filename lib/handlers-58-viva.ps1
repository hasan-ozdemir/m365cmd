# Handler: Viva
# Purpose: Viva command handlers.
function Handle-VivaCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: viva provider|content ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    if (-not $rest -or $rest.Count -eq 0) {
        Write-Warn "Usage: viva provider|content ..."
        return
    }
    $action = $rest[0].ToLowerInvariant()
    $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest2
    $apiInfo = Resolve-CAApiSettings $parsed

    switch ($sub) {
        "provider" {
            $base = "/employeeExperience/learningProviders"
            switch ($action) {
                "list" {
                    $qh = Build-QueryAndHeaders $parsed.Map @("id", "displayName", "isEnabled")
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp -and $resp.value) {
                        Write-GraphTable $resp.value @("Id", "DisplayName", "IsEnabled")
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 10
                    }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: viva provider get <providerId>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "create" {
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") $null
                    if (-not $body) {
                        Write-Warn "Usage: viva provider create --json <payload>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: viva provider update <providerId> --json <payload> OR --set key=value"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp -ne $null) { Write-Info "Provider updated." }
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: viva provider delete <providerId> [--force]"
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
                    if ($resp -ne $null) { Write-Info "Provider deleted." }
                }
                default {
                    Write-Warn "Usage: viva provider list|get|create|update|delete"
                }
            }
        }
        "content" {
            $provider = Get-ArgValue $parsed.Map "provider"
            if (-not $provider) {
                Write-Warn "Usage: viva content <action> --provider <id>"
                return
            }
            $base = "/employeeExperience/learningProviders/" + $provider + "/learningContents"
            switch ($action) {
                "list" {
                    $qh = Build-QueryAndHeaders $parsed.Map @("id", "title", "isActive")
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp -and $resp.value) {
                        Write-GraphTable $resp.value @("Id", "Title", "IsActive")
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 10
                    }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: viva content get <contentId> --provider <id>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "create" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") $null
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: viva content create <externalId> --provider <id> --json <payload>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp -ne $null) { Write-Info "Content created." }
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: viva content update <contentId> --provider <id> --json <payload> OR --set key=value"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp -ne $null) { Write-Info "Content updated." }
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: viva content delete <contentId> --provider <id> [--force]"
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
                    if ($resp -ne $null) { Write-Info "Content deleted." }
                }
                default {
                    Write-Warn "Usage: viva content list|get|create|update|delete --provider <id>"
                }
            }
        }
        default {
            Write-Warn "Usage: viva provider|content ..."
        }
    }
}
