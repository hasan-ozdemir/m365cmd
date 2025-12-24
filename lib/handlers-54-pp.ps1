# Handler: Pp
# Purpose: Pp command handlers.
function Handle-PPCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: pp login|logout|status|env|req"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }

    switch ($sub) {
        "login" {
            $parsed = Parse-NamedArgs $rest
            $device = $parsed.Map.ContainsKey("device")
            $force = $parsed.Map.ContainsKey("force")
            $token = Get-PpToken -ForceLogin:$force -DeviceCode:$device
            if ($token) { Write-Info "Connected to Power Platform API." }
        }
        "logout" {
            Clear-PpTokenCache
            Write-Info "Power Platform token cleared."
        }
        "status" {
            Load-PpTokenCache
            if ($global:PpToken -and $global:PpTokenExpires -gt (Get-Date)) {
                Write-Host ("PP login: connected (expires " + $global:PpTokenExpires.ToString("s") + ")")
            } else {
                Write-Host "PP login: not connected"
            }
        }
        "env" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: pp env list|get|enable|disable"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed = Parse-NamedArgs $rest2
            $base = "/environmentmanagement/environments"
            switch ($action) {
                "list" {
                    $query = @()
                    $top = Get-ArgValue $parsed.Map "top"
                    $filter = Get-ArgValue $parsed.Map "filter"
                    $skiptoken = Get-ArgValue $parsed.Map "skiptoken"
                    if ($top) { $query += "`$top=" + $top }
                    if ($filter) { $query += "`$filter=" + (Encode-QueryValue $filter) }
                    if ($skiptoken) { $query += "`$skiptoken=" + (Encode-QueryValue $skiptoken) }
                    $path = $base + (if ($query.Count -gt 0) { "?" + ($query -join "&") } else { "" })
                    $resp = Invoke-PpRequest -Method "GET" -Path $path
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: pp env get <environmentId>"
                        return
                    }
                    $resp = Invoke-PpRequest -Method "GET" -Path ($base + "/" + $id)
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "enable" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: pp env enable <environmentId>"
                        return
                    }
                    $resp = Invoke-PpRequest -Method "POST" -Path ($base + "/" + $id + "/Enable") -AllowNullResponse
                    if ($resp -ne $null) { Write-Info "Environment enabled." }
                }
                "disable" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: pp env disable <environmentId>"
                        return
                    }
                    $resp = Invoke-PpRequest -Method "POST" -Path ($base + "/" + $id + "/Disable") -AllowNullResponse
                    if ($resp -ne $null) { Write-Info "Environment disabled." }
                }
                default {
                    Write-Warn "Usage: pp env list|get|enable|disable"
                }
            }
        }
        "app" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: pp app list|get --env <environmentId>"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed = Parse-NamedArgs $rest2
            $env = Get-ArgValue $parsed.Map "env"
            if (-not $env) { $env = Get-ArgValue $parsed.Map "environment" }
            if (-not $env) {
                Write-Warn "Usage: pp app list|get --env <environmentId>"
                return
            }
            $base = "/powerapps/environments/" + $env + "/apps"
            switch ($action) {
                "list" {
                    $query = @()
                    $top = Get-ArgValue $parsed.Map "top"
                    $skiptoken = Get-ArgValue $parsed.Map "skiptoken"
                    if ($top) { $query += "`$top=" + $top }
                    if ($skiptoken) { $query += "`$skiptoken=" + (Encode-QueryValue $skiptoken) }
                    $path = $base + (if ($query.Count -gt 0) { "?" + ($query -join "&") } else { "" })
                    $resp = Invoke-PpRequest -Method "GET" -Path $path
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: pp app get <appId|displayName> --env <environmentId>"
                        return
                    }
                    $path = $base
                    $items = @()
                    $next = $path
                    while ($next) {
                        $resp = Invoke-PpRequest -Method "GET" -Path $next -AllowNullResponse
                        if (-not $resp) { break }
                        if ($resp.value) { $items += $resp.value }
                        $next = $resp.nextLink
                        if (-not $next) { break }
                    }
                    $match = $items | Where-Object { $_.name -eq $id -or $_.properties.displayName -eq $id } | Select-Object -First 1
                    if ($match) {
                        $match | ConvertTo-Json -Depth 10
                    } else {
                        Write-Warn "App not found."
                    }
                }
                default {
                    Write-Warn "Usage: pp app list|get --env <environmentId>"
                }
            }
        }
        "flow" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: pp flow list|get|runs|actions --env <environmentId>"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed = Parse-NamedArgs $rest2
            $env = Get-ArgValue $parsed.Map "env"
            if (-not $env) { $env = Get-ArgValue $parsed.Map "environment" }
            if (-not $env) {
                Write-Warn "Usage: pp flow <action> --env <environmentId>"
                return
            }
            $workflowId = Get-ArgValue $parsed.Map "workflowId"
            if (-not $workflowId) { $workflowId = Get-ArgValue $parsed.Map "flowId" }
            switch ($action) {
                "list" {
                    $query = @()
                    $top = Get-ArgValue $parsed.Map "top"
                    $skiptoken = Get-ArgValue $parsed.Map "skiptoken"
                    if ($workflowId) { $query += "workflowId=" + (Encode-QueryValue $workflowId) }
                    if ($top) { $query += "`$top=" + $top }
                    if ($skiptoken) { $query += "`$skiptoken=" + (Encode-QueryValue $skiptoken) }
                    $path = "/powerautomate/environments/" + $env + "/cloudFlows" + (if ($query.Count -gt 0) { "?" + ($query -join "&") } else { "" })
                    $resp = Invoke-PpRequest -Method "GET" -Path $path
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: pp flow get <workflowId> --env <environmentId>"
                        return
                    }
                    $path = "/powerautomate/environments/" + $env + "/cloudFlows?workflowId=" + (Encode-QueryValue $id)
                    $resp = Invoke-PpRequest -Method "GET" -Path $path
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "runs" {
                    if (-not $workflowId) {
                        $workflowId = $parsed.Positionals | Select-Object -First 1
                    }
                    if (-not $workflowId) {
                        Write-Warn "Usage: pp flow runs --env <environmentId> --workflowId <id>"
                        return
                    }
                    $query = @("workflowId=" + (Encode-QueryValue $workflowId))
                    $top = Get-ArgValue $parsed.Map "top"
                    $skiptoken = Get-ArgValue $parsed.Map "skiptoken"
                    if ($top) { $query += "`$top=" + $top }
                    if ($skiptoken) { $query += "`$skiptoken=" + (Encode-QueryValue $skiptoken) }
                    $path = "/powerautomate/environments/" + $env + "/flowRuns?" + ($query -join "&")
                    $resp = Invoke-PpRequest -Method "GET" -Path $path
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "actions" {
                    if (-not $workflowId) {
                        $workflowId = $parsed.Positionals | Select-Object -First 1
                    }
                    if (-not $workflowId) {
                        Write-Warn "Usage: pp flow actions --env <environmentId> --workflowId <id>"
                        return
                    }
                    $query = @("workflowId=" + (Encode-QueryValue $workflowId))
                    $path = "/powerautomate/environments/" + $env + "/flowActions?" + ($query -join "&")
                    $resp = Invoke-PpRequest -Method "GET" -Path $path
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                default {
                    Write-Warn "Usage: pp flow list|get|runs|actions --env <environmentId>"
                }
            }
        }
        "req" {
            $parsed = Parse-NamedArgs $rest
            $method = $parsed.Positionals | Select-Object -First 1
            $path = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $method -or -not $path) {
                Write-Warn "Usage: pp req <method> <path> [--json <payload>] [--bodyFile <file>] [--apiVersion <ver>]"
                return
            }
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") $null
            $apiVersion = Get-ArgValue $parsed.Map "apiVersion"
            $resp = Invoke-PpRequest -Method ($method.ToUpperInvariant()) -Path $path -Body $body -ApiVersion $apiVersion -AllowNullResponse
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        default {
            Write-Warn "Usage: pp login|logout|status|env|req"
        }
    }
}

