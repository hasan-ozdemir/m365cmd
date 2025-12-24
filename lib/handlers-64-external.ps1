# Handler: External
# Purpose: External connections/items helpers.
function Handle-ExternalCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: external connection|item <args...>"
        return
    }
    if (-not (Require-GraphConnection)) { return }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }

    switch ($sub) {
        "connection" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: external connection list|get|add|remove|schema|urltoitemresolver|doctor"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed = Parse-NamedArgs $rest2
            switch ($action) {
                "list" {
                    $resp = Invoke-GraphRequest -Method "GET" -Uri "/external/connections"
                    if ($resp) { $resp.value | ConvertTo-Json -Depth 8 }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) { $id = Get-ArgValue $parsed.Map "id" }
                    if (-not $id) { Write-Warn "Usage: external connection get <id>"; return }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/external/connections/" + $id)
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                }
                "add" {
                    $id = Get-ArgValue $parsed.Map "id"
                    if (-not $id) { $id = $parsed.Positionals | Select-Object -First 1 }
                    $name = Get-ArgValue $parsed.Map "name"
                    $desc = Get-ArgValue $parsed.Map "description"
                    $jsonRaw = Get-ArgValue $parsed.Map "json"
                    $body = if ($jsonRaw) { Parse-Value $jsonRaw } else { @{ id = $id; name = $name; description = $desc } }
                    if (-not $body -or -not $body.id) {
                        Write-Warn "Usage: external connection add --id <id> --name <name> [--description <text>] OR --json <payload>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "POST" -Uri "/external/connections" -Body $body
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                }
                "remove" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) { $id = Get-ArgValue $parsed.Map "id" }
                    if (-not $id) { Write-Warn "Usage: external connection remove <id> [--force]"; return }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host "Type DELETE to confirm"
                        if ($confirm -ne "DELETE") { Write-Info "Canceled."; return }
                    }
                    $resp = Invoke-GraphRequest -Method "DELETE" -Uri ("/external/connections/" + $id)
                    if ($resp -ne $null) { Write-Info "Connection removed." }
                }
                "schema" {
                    $action2 = $parsed.Positionals | Select-Object -First 1
                    if (-not $action2 -or $action2 -ne "add") {
                        Write-Warn "Usage: external connection schema add --id <id> --schema <json> [--wait]"
                        return
                    }
                    $id = Get-ArgValue $parsed.Map "id"
                    if (-not $id) { $id = $parsed.Positionals | Select-Object -Skip 1 -First 1 }
                    $schemaRaw = Get-ArgValue $parsed.Map "schema"
                    $wait = Parse-Bool (Get-ArgValue $parsed.Map "wait") $false
                    if (-not $id -or -not $schemaRaw) {
                        Write-Warn "Usage: external connection schema add --id <id> --schema <json> [--wait]"
                        return
                    }
                    $schema = Parse-Value $schemaRaw
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri ("/external/connections/" + $id + "/schema") -Body $schema
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                    if ($wait) {
                        Write-Info "Waiting for schema provisioning (polling latest operation)..."
                        Start-Sleep -Seconds 10
                        $op = Invoke-GraphRequest -Method "GET" -Uri ("/external/connections/" + $id + "/operations")
                        if ($op -and $op.value) { $op.value | Select-Object -First 1 | ConvertTo-Json -Depth 6 }
                    }
                }
                "urltoitemresolver" {
                    $action2 = $parsed.Positionals | Select-Object -First 1
                    if (-not $action2 -or $action2 -ne "add") {
                        Write-Warn "Usage: external connection urltoitemresolver add --id <id> --baseUrls <u1,u2> --urlPattern <pattern> --itemId <id> --priority <n>"
                        return
                    }
                    $id = Get-ArgValue $parsed.Map "id"
                    if (-not $id) { $id = $parsed.Positionals | Select-Object -Skip 1 -First 1 }
                    $baseUrls = Parse-CommaList (Get-ArgValue $parsed.Map "baseUrls")
                    $urlPattern = Get-ArgValue $parsed.Map "urlPattern"
                    $itemId = Get-ArgValue $parsed.Map "itemId"
                    $priority = Get-ArgValue $parsed.Map "priority"
                    if (-not $id -or -not $baseUrls -or -not $urlPattern -or -not $itemId -or -not $priority) {
                        Write-Warn "Usage: external connection urltoitemresolver add --id <id> --baseUrls <u1,u2> --urlPattern <pattern> --itemId <id> --priority <n>"
                        return
                    }
                    $body = @{
                        activitySettings = @{
                            urlToItemResolvers = @(
                                @{
                                    '@odata.type' = "#microsoft.graph.externalConnectors.itemIdResolver"
                                    itemId        = $itemId
                                    priority      = [int]$priority
                                    urlMatchInfo  = @{
                                        baseUrls  = $baseUrls
                                        urlPattern = $urlPattern
                                    }
                                }
                            )
                        }
                    }
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri ("/external/connections/" + $id) -Body $body
                    if ($resp -ne $null) { Write-Info "URL resolver added." }
                }
                "doctor" {
                    $id = Get-ArgValue $parsed.Map "id"
                    if (-not $id) { $id = $parsed.Positionals | Select-Object -First 1 }
                    if (-not $id) { Write-Warn "Usage: external connection doctor <id>"; return }
                    $conn = Invoke-GraphRequest -Method "GET" -Uri ("/external/connections/" + $id)
                    if (-not $conn) { Write-Warn "Connection not found."; return }
                    $schema = Invoke-GraphRequest -Method "GET" -Uri ("/external/connections/" + $id + "/schema")
                    $ok = if ($schema) { "ok" } else { "missing" }
                    Write-Host ("Connection : " + $conn.id)
                    Write-Host ("Schema     : " + $ok)
                }
                default {
                    Write-Warn "Usage: external connection list|get|add|remove|schema|urltoitemresolver|doctor"
                }
            }
        }
        "item" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: external item add --connection <id> --id <itemId> --json <payload>"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed = Parse-NamedArgs $rest2
            if ($action -ne "add") {
                Write-Warn "Usage: external item add --connection <id> --id <itemId> --json <payload>"
                return
            }
            $conn = Get-ArgValue $parsed.Map "connection"
            if (-not $conn) { $conn = Get-ArgValue $parsed.Map "conn" }
            $id = Get-ArgValue $parsed.Map "id"
            if (-not $id) { $id = $parsed.Positionals | Select-Object -First 1 }
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            if (-not $conn -or -not $id -or -not $jsonRaw) {
                Write-Warn "Usage: external item add --connection <id> --id <itemId> --json <payload>"
                return
            }
            $body = Parse-Value $jsonRaw
            $resp = Invoke-GraphRequest -Method "PUT" -Uri ("/external/connections/" + $conn + "/items/" + $id) -Body $body
            if ($resp) { $resp | ConvertTo-Json -Depth 8 }
        }
        default {
            Write-Warn "Usage: external connection|item <args...>"
        }
    }
}
