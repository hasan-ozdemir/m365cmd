# Handler: Forms
# Purpose: Forms command handlers.
function Get-FormsApiBase {
    $base = $global:Config.forms.baseUrl
    $api = $global:Config.forms.apiBase
    if (-not $base) { $base = "https://forms.office.com" }
    if (-not $api) { return $base }
    return ($base.TrimEnd("/") + $api)
}

function Resolve-FormsExcelItemPath {
    param([hashtable]$Map)
    $itemPath = Resolve-ExcelItemPath $Map
    if (-not $itemPath) {
        Write-Warn "Workbook item required: use --item <id> or --path <path> (and optional --user/--drive/--site)"
        return $null
    }
    return $itemPath
}

function Get-ExcelTableList {
    param([string]$ItemPath)
    if (-not $ItemPath) { return $null }
    return Invoke-GraphRequest -Method "GET" -Uri ($ItemPath + "/workbook/tables")
}

function Resolve-ExcelTableNameOrFirst {
    param(
        [string]$ItemPath,
        [string]$Table
    )
    if ($Table) { return $Table }
    $resp = Get-ExcelTableList $ItemPath
    if ($resp -and $resp.value -and $resp.value.Count -gt 0) {
        return $resp.value[0].id
    }
    Write-Warn "No tables found in workbook."
    return $null
}

function Get-ExcelTableRows {
    param(
        [string]$ItemPath,
        [string]$Table,
        [hashtable]$Map
    )
    if (-not $ItemPath -or -not $Table) { return $null }
    $qh = Build-QueryAndHeaders $Map @()
    $uri = $ItemPath + "/workbook/tables/" + $Table + "/rows" + $qh.Query
    return Invoke-GraphRequest -Method "GET" -Uri $uri -Headers $qh.Headers
}

function Format-ExcelRowValues {
    param([object]$Row)
    if (-not $Row -or -not $Row.values -or $Row.values.Count -eq 0) { return "" }
    $vals = $Row.values[0]
    return ($vals | ForEach-Object { "$_" }) -join " | "
}

function Write-ExcelRowsTable {
    param([object[]]$Rows)
    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-Info "No rows found."
        return
    }
    $rowsOut = $Rows | ForEach-Object {
        [pscustomobject]@{
            Index  = $_.index
            Values = (Format-ExcelRowValues $_)
        }
    }
    $rowsOut | Format-Table -AutoSize
}

function Handle-FormsExcelCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: forms excel tables|rows|watch ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $itemPath = Resolve-FormsExcelItemPath $parsed.Map
    if (-not $itemPath) { return }

    switch ($action) {
        "tables" {
            $resp = Get-ExcelTableList $itemPath
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Id","Name","ShowHeaders","ShowTotals")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 8
            }
        }
        "rows" {
            $table = Get-ArgValue $parsed.Map "table"
            $table = Resolve-ExcelTableNameOrFirst $itemPath $table
            if (-not $table) { return }
            $resp = Get-ExcelTableRows $itemPath $table $parsed.Map
            if (-not $resp) { return }
            $asJson = $parsed.Map.ContainsKey("json")
            if ($asJson) {
                if ($resp.value) { $resp.value | ConvertTo-Json -Depth 8 } else { $resp | ConvertTo-Json -Depth 8 }
            } else {
                if ($resp.value) { Write-ExcelRowsTable $resp.value } else { $resp | ConvertTo-Json -Depth 8 }
            }
        }
        "watch" {
            $table = Get-ArgValue $parsed.Map "table"
            $table = Resolve-ExcelTableNameOrFirst $itemPath $table
            if (-not $table) { return }
            $interval = 10
            $intervalRaw = Get-ArgValue $parsed.Map "interval"
            if ($intervalRaw) { try { $interval = [int]$intervalRaw } catch {} }
            if ($interval -le 0) { $interval = 10 }
            $fromNow = Parse-Bool (Get-ArgValue $parsed.Map "fromNow") $true
            $asJson = $parsed.Map.ContainsKey("json")
            $maxRaw = Get-ArgValue $parsed.Map "max"
            $max = $null
            if ($maxRaw) { try { $max = [int]$maxRaw } catch {} }

            $lastIndex = -1
            if ($fromNow) {
                $respInit = Get-ExcelTableRows $itemPath $table $parsed.Map
                if ($respInit -and $respInit.value) {
                    $lastIndex = ($respInit.value | Measure-Object -Property index -Maximum).Maximum
                }
                Write-Info ("Watching for new rows (interval " + $interval + "s).")
            } else {
                $respInit = Get-ExcelTableRows $itemPath $table $parsed.Map
                if ($respInit -and $respInit.value) {
                    if ($asJson) {
                        $respInit.value | ConvertTo-Json -Depth 8
                    } else {
                        Write-ExcelRowsTable $respInit.value
                    }
                    $lastIndex = ($respInit.value | Measure-Object -Property index -Maximum).Maximum
                }
                Write-Info ("Watching for new rows (interval " + $interval + "s).")
            }

            $printed = 0
            while ($true) {
                Start-Sleep -Seconds $interval
                $resp = Get-ExcelTableRows $itemPath $table $parsed.Map
                if (-not $resp -or -not $resp.value) { continue }
                $newRows = @($resp.value | Where-Object { $_.index -gt $lastIndex })
                if ($newRows.Count -gt 0) {
                    if ($asJson) {
                        $newRows | ConvertTo-Json -Depth 8
                    } else {
                        Write-ExcelRowsTable $newRows
                    }
                    $lastIndex = ($newRows | Measure-Object -Property index -Maximum).Maximum
                    $printed += $newRows.Count
                    if ($max -and $printed -ge $max) { break }
                }
            }
        }
        default {
            Write-Warn "Usage: forms excel tables|rows|watch ..."
        }
    }
}

function Handle-FormsAdminCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: forms admin get|update"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $path = "/admin/forms"

    switch ($action) {
        "get" {
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri $path -Api "beta"
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "update" {
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            $bodyFile = Get-ArgValue $parsed.Map "bodyFile"
            $setRaw = Get-ArgValue $parsed.Map "set"
            $body = Read-JsonPayload $jsonRaw $bodyFile $setRaw
            if (-not $body) {
                Write-Warn "Usage: forms admin update --json <payload> OR --set key=value[,key=value]"
                return
            }
            if ($setRaw -and -not $body.settings) {
                $body = @{ settings = $body }
            }
            $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri $path -Body $body -Api "beta"
            if ($resp -ne $null) { Write-Info "Forms settings updated." }
        }
        default {
            Write-Warn "Usage: forms admin get|update"
        }
    }
}

function Handle-FormsReportCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: forms report list|run"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $reports = @(
        "getFormsUserActivityUserCounts",
        "getFormsUserActivityCounts",
        "getFormsUserActivityUserDetail"
    )

    switch ($sub) {
        "list" {
            $reports | Sort-Object | Format-Wide -Column 2
        }
        "run" {
            $name = $parsed.Positionals | Select-Object -First 1
            if (-not $name) {
                Write-Warn "Usage: forms report run <name> [--period D7|D30|D90|D180] [--date YYYY-MM-DD] [--format csv|json] [--out <file>] [--beta|--v1|--auto]"
                return
            }
            $period = Get-ArgValue $parsed.Map "period"
            $date = Get-ArgValue $parsed.Map "date"
            $format = Get-ArgValue $parsed.Map "format"
            $out = Get-ArgValue $parsed.Map "out"
            $useBeta = $parsed.Map.ContainsKey("beta")
            $useV1 = $parsed.Map.ContainsKey("v1")
            $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
            if ($parsed.Map.ContainsKey("auto")) { $allowFallback = $true }
            if ($useBeta -or $useV1) { $allowFallback = $false }
            $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }

            $paramSeg = ""
            if ($period) {
                $paramSeg = "(period='" + $period + "')"
            } elseif ($date) {
                $paramSeg = "(date=" + $date + ")"
            }
            $path = "/reports/" + $name + $paramSeg
            if ($format -and $format.ToLowerInvariant() -eq "json") {
                $path = $path + "?`$format=application/json"
            }

            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri $path -Api $api -AllowFallback:$allowFallback -AllowNullResponse
            if ($out) {
                if ($resp -is [string]) {
                    Set-Content -Path $out -Value $resp -Encoding ASCII
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 10 | Set-Content -Path $out -Encoding ASCII
                }
                Write-Info ("Saved: " + $out)
            } elseif ($resp -is [string]) {
                Write-Host $resp
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        default {
            Write-Warn "Usage: forms report list|run"
        }
    }
}

function Handle-FormsFlowCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: forms flow list|get|runs|actions --env <environmentId> [--name <text>] [--contains <text>] [--workflowId <id>]"
        return
    }
    $action = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $env = Get-ArgValue $parsed.Map "env"
    if (-not $env) { $env = Get-ArgValue $parsed.Map "environment" }
    if (-not $env) {
        Write-Warn "Usage: forms flow <action> --env <environmentId>"
        return
    }

    $name = Get-ArgValue $parsed.Map "name"
    $contains = Get-ArgValue $parsed.Map "contains"
    $form = Get-ArgValue $parsed.Map "form"
    $workflowId = Get-ArgValue $parsed.Map "workflowId"
    if (-not $workflowId) { $workflowId = Get-ArgValue $parsed.Map "flowId" }
    $top = Get-ArgValue $parsed.Map "top"
    $skiptoken = Get-ArgValue $parsed.Map "skiptoken"
    $all = $parsed.Map.ContainsKey("all")

    switch ($action) {
        "list" {
            $query = @()
            if ($top) { $query += "`$top=" + $top }
            if ($skiptoken) { $query += "`$skiptoken=" + (Encode-QueryValue $skiptoken) }
            $path = "/powerautomate/environments/" + $env + "/cloudFlows" + (if ($query.Count -gt 0) { "?" + ($query -join "&") } else { "" })
            $items = @()
            $next = $path
            while ($next) {
                $resp = Invoke-PpRequest -Method "GET" -Path $next -AllowNullResponse
                if (-not $resp) { break }
                if ($resp.value) { $items += $resp.value }
                if ($all) {
                    $next = $resp.nextLink
                } else {
                    $next = $null
                }
            }
            $filterText = if ($name) { $name } elseif ($contains) { $contains } elseif ($form) { $form } else { $null }
            if ($filterText) {
                $items = @($items | Where-Object {
                    ($_.properties.displayName -and $_.properties.displayName -like ("*" + $filterText + "*")) -or
                    ($_.name -and $_.name -like ("*" + $filterText + "*")) -or
                    ($_.properties.description -and $_.properties.description -like ("*" + $filterText + "*"))
                })
            }
            if ($parsed.Map.ContainsKey("json")) {
                $items | ConvertTo-Json -Depth 8
            } else {
                if ($items.Count -eq 0) { Write-Info "No flows found." }
                else { $items | Select-Object name, @{n="displayName";e={$_.properties.displayName}}, @{n="state";e={$_.properties.state}} | Format-Table -AutoSize }
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id -and $workflowId) { $id = $workflowId }
            if (-not $id -and $name) {
                $tmp = Handle-FormsFlowCommand @("list","--env", $env, "--name", $name, "--json")
                return
            }
            if (-not $id) {
                Write-Warn "Usage: forms flow get <workflowId> --env <environmentId>"
                return
            }
            $path = "/powerautomate/environments/" + $env + "/cloudFlows?workflowId=" + (Encode-QueryValue $id)
            $resp = Invoke-PpRequest -Method "GET" -Path $path
            if ($resp) { $resp | ConvertTo-Json -Depth 8 }
        }
        "runs" {
            if (-not $workflowId) { $workflowId = $parsed.Positionals | Select-Object -First 1 }
            if (-not $workflowId) {
                Write-Warn "Usage: forms flow runs --env <environmentId> --workflowId <id>"
                return
            }
            $query = @("workflowId=" + (Encode-QueryValue $workflowId))
            if ($top) { $query += "`$top=" + $top }
            if ($skiptoken) { $query += "`$skiptoken=" + (Encode-QueryValue $skiptoken) }
            $path = "/powerautomate/environments/" + $env + "/flowRuns?" + ($query -join "&")
            $resp = Invoke-PpRequest -Method "GET" -Path $path
            if ($resp) { $resp | ConvertTo-Json -Depth 8 }
        }
        "actions" {
            if (-not $workflowId) { $workflowId = $parsed.Positionals | Select-Object -First 1 }
            if (-not $workflowId) {
                Write-Warn "Usage: forms flow actions --env <environmentId> --workflowId <id>"
                return
            }
            $query = @("workflowId=" + (Encode-QueryValue $workflowId))
            $path = "/powerautomate/environments/" + $env + "/flowActions?" + ($query -join "&")
            $resp = Invoke-PpRequest -Method "GET" -Path $path
            if ($resp) { $resp | ConvertTo-Json -Depth 8 }
        }
        default {
            Write-Warn "Usage: forms flow list|get|runs|actions --env <environmentId> [--name <text>] [--contains <text>] [--workflowId <id>]"
        }
    }
}

function Handle-FormsRawCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -lt 2) {
        Write-Warn "Usage: forms raw <get|post|patch|put|delete> <path> [--body <json>] [--bodyFile <path>] [--out <file>]"
        return
    }
    Write-Warn "Forms raw uses Forms service endpoints that may be unsupported or change without notice."
    $method = $InputArgs[0].ToUpperInvariant()
    $path = $InputArgs[1]
    $rest = if ($InputArgs.Count -gt 2) { $InputArgs[2..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $bodyRaw = Get-ArgValue $parsed.Map "body"
    $bodyFile = Get-ArgValue $parsed.Map "bodyFile"
    $out = Get-ArgValue $parsed.Map "out"
    $body = $null
    if ($bodyFile) {
        if (-not (Test-Path $bodyFile)) {
            Write-Warn "Body file not found."
            return
        }
        $body = Get-Content -Raw -Path $bodyFile
    } elseif ($bodyRaw) {
        $body = Parse-Value $bodyRaw
    }
    $base = Get-FormsApiBase
    $scope = "https://forms.office.com/.default"
    $resp = Invoke-ExternalApiRequest -Method $method -Url $path -Body $body -Scope $scope -BaseUrl $base -AllowNullResponse
    if ($out) {
        if ($resp -is [string]) {
            Set-Content -Path $out -Value $resp -Encoding ASCII
        } elseif ($resp) {
            $resp | ConvertTo-Json -Depth 8 | Set-Content -Path $out -Encoding ASCII
        }
        Write-Info ("Saved: " + $out)
    } elseif ($resp) {
        if ($resp.value) { $resp.value | ConvertTo-Json -Depth 8 } else { $resp | ConvertTo-Json -Depth 8 }
    }
}

function Handle-FormsCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: forms open|info|admin|report|raw|excel|flow ..."
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    switch ($sub) {
        "open" {
            Write-Host "https://forms.office.com/"
        }
        "info" {
            Write-Warn "Microsoft Forms does not have a supported Microsoft Graph API. Use forms report/admin or forms raw (undocumented) and Excel responses."
        }
        "admin"  { Handle-FormsAdminCommand $rest }
        "report" { Handle-FormsReportCommand $rest }
        "raw"    { Handle-FormsRawCommand $rest }
        "excel"  { Handle-FormsExcelCommand $rest }
        "flow"   { Handle-FormsFlowCommand $rest }
        default  { Write-Warn "Usage: forms open|info|admin|report|raw|excel|flow ..." }
    }
}

