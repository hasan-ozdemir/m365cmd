# Handler: Intune
# Purpose: Intune command handlers.
function Invoke-IntuneCrud {
    param(
        [string]$Action,
        [string]$Base,
        [object]$Parsed,
        [string]$UsagePrefix,
        [string[]]$DefaultProps,
        [string[]]$TableProps
    )
    $apiInfo = Resolve-CAApiSettings $Parsed
    switch ($Action) {
        "list" {
            $qh = Build-QueryAndHeaders $Parsed.Map $DefaultProps
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($Base + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
            if ($resp -and $resp.value) {
                if ($TableProps -and $TableProps.Count -gt 0) {
                    Write-GraphTable $resp.value $TableProps
                } else {
                    $resp.value | ConvertTo-Json -Depth 10
                }
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "get" {
            $id = $Parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn ("Usage: " + $UsagePrefix + " get <id>")
                return
            }
            $qh = Build-QueryAndHeaders $Parsed.Map @()
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($Base + "/" + $id + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "create" {
            $body = Read-JsonPayload (Get-ArgValue $Parsed.Map "json") (Get-ArgValue $Parsed.Map "bodyFile") $null
            if (-not $body) {
                Write-Warn ("Usage: " + $UsagePrefix + " create --json <payload> [--bodyFile <file>]")
                return
            }
            $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $Base -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
            if ($resp) {
                Write-Info "Created."
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "update" {
            $id = $Parsed.Positionals | Select-Object -First 1
            $body = Read-JsonPayload (Get-ArgValue $Parsed.Map "json") (Get-ArgValue $Parsed.Map "bodyFile") (Get-ArgValue $Parsed.Map "set")
            if (-not $id -or -not $body) {
                Write-Warn ("Usage: " + $UsagePrefix + " update <id> --json <payload> [--bodyFile <file>] OR --set key=value")
                return
            }
            $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($Base + "/" + $id) -Body $body -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
            if ($resp -ne $null) { Write-Info "Updated." }
        }
        "delete" {
            $id = $Parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn ("Usage: " + $UsagePrefix + " delete <id> [--force]")
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
            $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri ($Base + "/" + $id) -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback -AllowNullResponse
            if ($resp -ne $null) { Write-Info "Deleted." }
        }
        default {
            Write-Warn ("Usage: " + $UsagePrefix + " list|get|create|update|delete")
        }
    }
}

function Handle-IntuneCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: intune device|config|compliance|app|script|shellscript|healthscript|report ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    if (-not $rest -or $rest.Count -eq 0) {
        Write-Warn "Usage: intune device|config|compliance|app|script|shellscript|healthscript list|get|create|update|delete OR intune report ..."
        return
    }
    $action = $rest[0].ToLowerInvariant()
    $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest2

    switch ($sub) {
        "device" { 
            Invoke-IntuneCrud -Action $action -Base "/deviceManagement/managedDevices" -Parsed $parsed -UsagePrefix "intune device" `
                -DefaultProps @("id", "deviceName", "userPrincipalName", "operatingSystem", "complianceState") `
                -TableProps @("Id", "DeviceName", "UserPrincipalName", "OperatingSystem", "ComplianceState")
        }
        "config" {
            Invoke-IntuneCrud -Action $action -Base "/deviceManagement/deviceConfigurations" -Parsed $parsed -UsagePrefix "intune config" `
                -DefaultProps @("id", "displayName", "description", "platforms", "createdDateTime") `
                -TableProps @("Id", "DisplayName", "Description", "Platforms", "CreatedDateTime")
        }
        "compliance" {
            Invoke-IntuneCrud -Action $action -Base "/deviceManagement/deviceCompliancePolicies" -Parsed $parsed -UsagePrefix "intune compliance" `
                -DefaultProps @("id", "displayName", "description", "platforms", "createdDateTime") `
                -TableProps @("Id", "DisplayName", "Description", "Platforms", "CreatedDateTime")
        }
        "app" {
            Invoke-IntuneCrud -Action $action -Base "/deviceAppManagement/mobileApps" -Parsed $parsed -UsagePrefix "intune app" `
                -DefaultProps @("id", "displayName", "publisher", "isFeatured", "createdDateTime") `
                -TableProps @("Id", "DisplayName", "Publisher", "IsFeatured", "CreatedDateTime")
        }
        "script" {
            Invoke-IntuneCrud -Action $action -Base "/deviceManagement/deviceManagementScripts" -Parsed $parsed -UsagePrefix "intune script" `
                -DefaultProps @("id", "displayName", "createdDateTime", "lastModifiedDateTime") `
                -TableProps @("Id", "DisplayName", "CreatedDateTime", "LastModifiedDateTime")
        }
        "shellscript" {
            Invoke-IntuneCrud -Action $action -Base "/deviceManagement/deviceShellScripts" -Parsed $parsed -UsagePrefix "intune shellscript" `
                -DefaultProps @("id", "displayName", "createdDateTime", "lastModifiedDateTime") `
                -TableProps @("Id", "DisplayName", "CreatedDateTime", "LastModifiedDateTime")
        }
        "healthscript" {
            Invoke-IntuneCrud -Action $action -Base "/deviceManagement/deviceHealthScripts" -Parsed $parsed -UsagePrefix "intune healthscript" `
                -DefaultProps @("id", "displayName", "createdDateTime", "lastModifiedDateTime") `
                -TableProps @("Id", "DisplayName", "CreatedDateTime", "LastModifiedDateTime")
        }
        "report" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: intune report list|get|export|download|status"
                return
            }
            $action2 = $rest[0].ToLowerInvariant()
            $rest3 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest3
            $useBeta = $parsed2.Map.ContainsKey("beta")
            $useV1 = $parsed2.Map.ContainsKey("v1")
            $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
            if ($useBeta -or $useV1) { $allowFallback = $false }
            $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
            $base = "/deviceManagement/reports/exportJobs"

            switch ($action2) {
                "list" {
                    $qh = Build-QueryAndHeaders $parsed2.Map @()
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Api $api -AllowFallback:$allowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "get" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: intune report get <jobId>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "status" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: intune report status <jobId>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "export" {
                    $name = Get-ArgValue $parsed2.Map "name"
                    if (-not $name) {
                        Write-Warn "Usage: intune report export --name <reportName> [--format csv|json] [--select col1,col2] [--filter <odata>]"
                        return
                    }
                    $format = Get-ArgValue $parsed2.Map "format"
                    if (-not $format) { $format = "csv" }
                    $select = Get-ArgValue $parsed2.Map "select"
                    $filter = Get-ArgValue $parsed2.Map "filter"
                    $body = @{ reportName = $name; format = $format }
                    if ($select) { $body.select = Parse-CommaList $select }
                    if ($filter) { $body.filter = $filter }
                    $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $api -AllowFallback:$allowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "download" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    $out = Get-ArgValue $parsed2.Map "out"
                    if (-not $id -or -not $out) {
                        Write-Warn "Usage: intune report download <jobId> --out <file>"
                        return
                    }
                    $job = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                    if (-not $job -or -not $job.url) {
                        Write-Warn "Export job URL not ready. Check status first."
                        return
                    }
                    try {
                        Invoke-WebRequest -Uri $job.url -OutFile $out | Out-Null
                        Write-Info ("Saved: " + $out)
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                default {
                    Write-Warn "Usage: intune report list|get|export|download|status"
                }
            }
        }
        default {
            Write-Warn "Usage: intune device|config|compliance|app|script|shellscript|healthscript list|get|create|update|delete OR intune report ..."
        }
    }
}

