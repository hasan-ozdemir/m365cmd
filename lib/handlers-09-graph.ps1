# Handler: Graph
# Purpose: Graph command handlers.
function Handle-SubscriptionCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: subscription list|get|create|update|delete"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $useBeta = Parse-Bool (Get-ArgValue $parsed.Map "beta") $false
    $base = "/subscriptions"

    switch ($action) {
        "list" {
            $resp = Invoke-GraphRequest -Method "GET" -Uri $base -Beta:$useBeta
            if ($resp -and $resp.value) {
                $resp.value | ConvertTo-Json -Depth 10
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: subscription get <id>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id) -Beta:$useBeta
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "create" {
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") $null
            if (-not $body) {
                Write-Warn "Usage: subscription create --json <payload>"
                return
            }
            $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body -Beta:$useBeta
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $id -or -not $body) {
                Write-Warn "Usage: subscription update <id> --json <payload> OR --set key=value"
                return
            }
            $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Beta:$useBeta
            if ($resp -ne $null) { Write-Info "Subscription updated." }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: subscription delete <id>"
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
            $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($base + "/" + $id) -Beta:$useBeta
            if ($resp -ne $null) { Write-Info "Subscription deleted." }
        }
        default {
            Write-Warn "Usage: subscription list|get|create|update|delete"
        }
    }
}


function Handle-TeamsTabCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: teamstab list|get|create|update|delete --team <teamId> --channel <channelId> OR --chat <chatId>"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $team = Get-ArgValue $parsed.Map "team"
    $channel = Get-ArgValue $parsed.Map "channel"
    $chat = Get-ArgValue $parsed.Map "chat"
    $base = $null
    if ($chat) {
        $base = "/chats/" + $chat + "/tabs"
    } elseif ($team -and $channel) {
        $base = "/teams/" + $team + "/channels/" + $channel + "/tabs"
    } else {
        Write-Warn "Usage: teamstab <action> --team <teamId> --channel <channelId> OR --chat <chatId>"
        return
    }

    switch ($action) {
        "list" {
            $resp = Invoke-GraphRequest -Method "GET" -Uri $base
            if ($resp -and $resp.value) {
                $resp.value | ConvertTo-Json -Depth 10
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: teamstab get <tabId> --team <teamId> --channel <channelId> OR --chat <chatId>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "create" {
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $body) {
                Write-Warn "Usage: teamstab create --team <teamId> --channel <channelId> --json <payload> OR --chat <chatId> --json <payload>"
                return
            }
            $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $id -or -not $body) {
                Write-Warn "Usage: teamstab update <tabId> --team <teamId> --channel <channelId> --json <payload> OR --chat <chatId> --json <payload>"
                return
            }
            $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
            if ($resp -ne $null) { Write-Info "Tab updated." }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: teamstab delete <tabId> --team <teamId> --channel <channelId> OR --chat <chatId>"
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
            if ($resp -ne $null) { Write-Info "Tab deleted." }
        }
        default {
            Write-Warn "Usage: teamstab list|get|create|update|delete --team <teamId> --channel <channelId> OR --chat <chatId>"
        }
    }
}


function Handle-TeamsAppCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: teamsapp list|get|update|delete"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $useBeta = Parse-Bool (Get-ArgValue $parsed.Map "beta") $false
    $base = "/appCatalogs/teamsApps"

    switch ($action) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @("id","displayName","distributionMethod")
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Beta:$useBeta
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Id","DisplayName","DistributionMethod")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: teamsapp get <appId>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id) -Beta:$useBeta
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "add" {
            $pkg = Get-ArgValue $parsed.Map "package"
            if (-not $pkg) {
                Write-Warn "Usage: teamsapp add --package <zipPath>"
                return
            }
            if (-not (Test-Path $pkg)) {
                Write-Warn "Package file not found."
                return
            }
            $bytes = [System.IO.File]::ReadAllBytes($pkg)
            $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $bytes -ContentType "application/zip" -Beta:$useBeta
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $pkg = Get-ArgValue $parsed.Map "package"
            if (-not $id -or -not $pkg) {
                Write-Warn "Usage: teamsapp update <appId> --package <zipPath>"
                return
            }
            if (-not (Test-Path $pkg)) {
                Write-Warn "Package file not found."
                return
            }
            $bytes = [System.IO.File]::ReadAllBytes($pkg)
            $resp = Invoke-GraphRequest -Method "POST" -Uri ($base + "/" + $id + "/appDefinitions") -Body $bytes -ContentType "application/zip" -Beta:$useBeta
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: teamsapp delete <appId>"
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
            $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($base + "/" + $id) -Beta:$useBeta
            if ($resp -ne $null) { Write-Info "Teams app deleted." }
        }
        default {
            Write-Warn "Usage: teamsapp list|get|add|update|delete"
        }
    }
}

function Handle-TeamsAppInstallCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: teamsappinst list|get|add|remove --team <teamId>|--chat <chatId>|--user <upn|id>"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $team = Get-ArgValue $parsed.Map "team"
    $chat = Get-ArgValue $parsed.Map "chat"
    $user = Get-ArgValue $parsed.Map "user"

    $base = $null
    if ($team) {
        $base = "/teams/" + $team + "/installedApps"
    } elseif ($chat) {
        $base = "/chats/" + $chat + "/installedApps"
    } elseif ($user) {
        $seg = Resolve-UserSegment $user
        if (-not $seg) { return }
        $base = $seg + "/teamwork/installedApps"
    } else {
        Write-Warn "Usage: teamsappinst <action> --team <teamId>|--chat <chatId>|--user <upn|id>"
        return
    }

    $useBeta = $parsed.Map.ContainsKey("beta")
    $useV1 = $parsed.Map.ContainsKey("v1")
    $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
    if ($useBeta -or $useV1) { $allowFallback = $false }
    $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }

    switch ($action) {
        "list" {
            $expand = Get-ArgValue $parsed.Map "expand"
            $qh = Build-QueryAndHeaders $parsed.Map @()
            $query = $qh.Query
            if ($expand) {
                $join = if ($query -match "\\?") { "&" } else { "?" }
                $query = $query + $join + "`$expand=" + (Encode-QueryValue $expand)
            }
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $query) -Headers $qh.Headers -Api $api -AllowFallback:$allowFallback
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: teamsappinst get <installationId> --team <teamId>|--chat <chatId>|--user <upn|id>"
                return
            }
            $qh = Build-QueryAndHeaders $parsed.Map @()
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id + $qh.Query) -Headers $qh.Headers -Api $api -AllowFallback:$allowFallback
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "add" {
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") $null
            if (-not $body) {
                $appId = Get-ArgValue $parsed.Map "app"
                if (-not $appId) {
                    Write-Warn "Usage: teamsappinst add --team <teamId>|--chat <chatId>|--user <upn|id> --app <appId> OR --json <payload>"
                    return
                }
                $baseApi = if ($useBeta) { "https://graph.microsoft.com/beta" } else { "https://graph.microsoft.com/v1.0" }
                $body = @{
                    "teamsApp@odata.bind" = ($baseApi + "/appCatalogs/teamsApps('" + $appId + "')")
                }
            }
            $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $api -AllowFallback:$allowFallback
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "remove" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: teamsappinst remove <installationId> --team <teamId>|--chat <chatId>|--user <upn|id>"
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
            $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback -AllowNullResponse
            if ($resp -ne $null) { Write-Info "Teams app installation removed." } else { Write-Info "Remove requested." }
        }
        default {
            Write-Warn "Usage: teamsappinst list|get|add|remove --team <teamId>|--chat <chatId>|--user <upn|id>"
        }
    }
}


function Handle-MeetingCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: meeting transcript|recording ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    if ($sub -in @("list", "get", "create", "update", "delete")) {
        $parsed = Parse-NamedArgs $rest
        $userSeg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
        if (-not $userSeg) { return }
        $base = $userSeg + "/onlineMeetings"
        $useBeta = $parsed.Map.ContainsKey("beta")
        $useV1 = $parsed.Map.ContainsKey("v1")
        $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
        if ($useBeta -or $useV1) { $allowFallback = $false }
        $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }

        switch ($sub) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed.Map @()
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: meeting get <meetingId> [--user <upn|id>]"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "create" {
                $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") $null
                if (-not $body) {
                    Write-Warn "Usage: meeting create --json <payload> [--user <upn|id>]"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "update" {
                $id = $parsed.Positionals | Select-Object -First 1
                $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                if (-not $id -or -not $body) {
                    Write-Warn "Usage: meeting update <meetingId> --json <payload> OR --set key=value [--user <upn|id>]"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $api -AllowFallback:$allowFallback
                if ($resp -ne $null) { Write-Info "Meeting updated." }
            }
            "delete" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: meeting delete <meetingId> [--user <upn|id>]"
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
                $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback -AllowNullResponse
                if ($resp -ne $null) { Write-Info "Meeting deleted." } else { Write-Info "Delete requested." }
            }
        }
        return
    }
    if (-not $rest -or $rest.Count -eq 0) {
        Write-Warn "Usage: meeting transcript|recording list|get|content --meeting <meetingId> [--user <upn|id>]"
        return
    }
    $action = $rest[0].ToLowerInvariant()
    $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest2
    $meetingId = Get-ArgValue $parsed.Map "meeting"
    if (-not $meetingId) {
        Write-Warn "Usage: meeting <type> <action> --meeting <meetingId> [--user <upn|id>]"
        return
    }
    $userSeg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
    if (-not $userSeg) { return }

    $useBeta = $parsed.Map.ContainsKey("beta")
    $useV1 = $parsed.Map.ContainsKey("v1")
    $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
    if ($useBeta -or $useV1) { $allowFallback = $false }
    $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }

    if ($sub -eq "transcript" -or $sub -eq "transcripts") {
        $base = $userSeg + "/onlineMeetings/" + $meetingId + "/transcripts"
        switch ($action) {
            "list" {
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri $base -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: meeting transcript get <transcriptId> --meeting <meetingId> [--user <upn|id>]"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "content" {
                $id = $parsed.Positionals | Select-Object -First 1
                $out = Get-ArgValue $parsed.Map "out"
                if (-not $id -or -not $out) {
                    Write-Warn "Usage: meeting transcript content <transcriptId> --meeting <meetingId> --out <file> [--user <upn|id>]"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                if (-not $resp -or -not $resp.transcriptContentUrl) {
                    Write-Warn "Transcript content URL not found."
                    return
                }
                try {
                    Invoke-WebRequest -Uri $resp.transcriptContentUrl -OutFile $out | Out-Null
                    Write-Info ("Saved: " + $out)
                } catch {
                    Write-Err $_.Exception.Message
                }
            }
            default {
                Write-Warn "Usage: meeting transcript list|get|content --meeting <meetingId> [--user <upn|id>]"
            }
        }
        return
    }

    if ($sub -eq "recording" -or $sub -eq "recordings") {
        $base = $userSeg + "/onlineMeetings/" + $meetingId + "/recordings"
        switch ($action) {
            "list" {
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri $base -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: meeting recording get <recordingId> --meeting <meetingId> [--user <upn|id>]"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "content" {
                $id = $parsed.Positionals | Select-Object -First 1
                $out = Get-ArgValue $parsed.Map "out"
                if (-not $id -or -not $out) {
                    Write-Warn "Usage: meeting recording content <recordingId> --meeting <meetingId> --out <file> [--user <upn|id>]"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                if (-not $resp -or -not $resp.recordingContentUrl) {
                    Write-Warn "Recording content URL not found."
                    return
                }
                try {
                    Invoke-WebRequest -Uri $resp.recordingContentUrl -OutFile $out | Out-Null
                    Write-Info ("Saved: " + $out)
                } catch {
                    Write-Err $_.Exception.Message
                }
            }
            default {
                Write-Warn "Usage: meeting recording list|get|content --meeting <meetingId> [--user <upn|id>]"
            }
        }
        return
    }

    Write-Warn "Usage: meeting transcript|recording list|get|content --meeting <meetingId> [--user <upn|id>]"
}


function Handle-SearchCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: search query --entity <type> --text <query> OR --requestsJson <payload>"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    if ($sub -ne "query") {
        Write-Warn "Usage: search query ..."
        return
    }
    $parsed = Parse-NamedArgs $rest
    $requestsJson = Get-ArgValue $parsed.Map "requestsJson"
    $body = $null
    if ($requestsJson) {
        $body = Parse-Value $requestsJson
    } else {
        $entity = Get-ArgValue $parsed.Map "entity"
        $text = Get-ArgValue $parsed.Map "text"
        if (-not $entity -or -not $text) {
            Write-Warn "Usage: search query --entity <type> --text <query>"
            return
        }
        $from = Get-ArgValue $parsed.Map "from"
        $size = Get-ArgValue $parsed.Map "size"
        $request = @{
            entityTypes = @($entity)
            query       = @{ queryString = $text }
        }
        if ($from) { $request.from = [int]$from }
        if ($size) { $request.size = [int]$size }
        $body = @{ requests = @($request) }
    }
    $resp = Invoke-GraphRequest -Method "POST" -Uri "/search/query" -Body $body
    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
}



function Handle-ExternalConnectionCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: extconn list|get|create|update|delete|item"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }

    if ($sub -eq "item" -or $sub -eq "items") {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: extconn item list|get|create|update|delete --conn <id>"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed = Parse-NamedArgs $rest2
        $conn = Get-ArgValue $parsed.Map "conn"
        if (-not $conn) {
            Write-Warn "Usage: extconn item <action> --conn <id>"
            return
        }
        $base = "/external/connections/" + $conn + "/items"
        switch ($action) {
            "list" {
                $resp = Invoke-GraphRequest -Method "GET" -Uri $base
                if ($resp -and $resp.value) {
                    $resp.value | ConvertTo-Json -Depth 10
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 10
                }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: extconn item get <id> --conn <id>"
                    return
                }
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "create" {
                $id = $parsed.Positionals | Select-Object -First 1
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                if (-not $id -or -not $jsonRaw) {
                    Write-Warn "Usage: extconn item create <id> --conn <id> --json <payload>"
                    return
                }
                $body = Parse-Value $jsonRaw
                $resp = Invoke-GraphRequest -Method "PUT" -Uri ($base + "/" + $id) -Body $body
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "update" {
                $id = $parsed.Positionals | Select-Object -First 1
                $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                if (-not $id -or -not $body) {
                    Write-Warn "Usage: extconn item update <id> --conn <id> --json <payload> OR --set key=value"
                    return
                }
                $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
                if ($resp -ne $null) { Write-Info "External item updated." }
            }
            "delete" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: extconn item delete <id> --conn <id>"
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
                if ($resp -ne $null) { Write-Info "External item deleted." }
            }
            default {
                Write-Warn "Usage: extconn item list|get|create|update|delete --conn <id>"
            }
        }
        return
    }

    $action = $sub
    $parsed = Parse-NamedArgs $rest
    $base = "/external/connections"
    switch ($action) {
        "list" {
            $resp = Invoke-GraphRequest -Method "GET" -Uri $base
            if ($resp -and $resp.value) {
                $resp.value | ConvertTo-Json -Depth 10
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: extconn get <id>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "create" {
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            if (-not $jsonRaw) {
                Write-Warn "Usage: extconn create --json <payload>"
                return
            }
            $body = Parse-Value $jsonRaw
            $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $id -or -not $body) {
                Write-Warn "Usage: extconn update <id> --json <payload> OR --set key=value"
                return
            }
            $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
            if ($resp -ne $null) { Write-Info "External connection updated." }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: extconn delete <id>"
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
            if ($resp -ne $null) { Write-Info "External connection deleted." }
        }
        default {
            Write-Warn "Usage: extconn list|get|create|update|delete"
        }
    }
}



function Handle-GraphCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: graph cmdlets|perms|req|meta|get|list|create|update|delete|action|batch"
        return
    }
    $sub = $Args[0].ToLowerInvariant()
    if ($sub -ne "meta" -and -not (Ensure-GraphModule)) { return }
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    switch ($sub) {
        "cmdlets" {
            $filter = Get-ArgValue $parsed.Map "filter"
            $cmds = Get-Command -Module Microsoft.Graph* -ErrorAction SilentlyContinue
            if ($filter) {
                $cmds = $cmds | Where-Object { $_.Name -like ("*" + $filter + "*") }
            }
            $cmds | Sort-Object Name | Select-Object -ExpandProperty Name | Format-Wide -Column 3
        }
        "perms" {
            if (-not (Require-GraphConnection)) { return }
            $graphSp = Get-GraphServicePrincipal
            if (-not $graphSp) {
                Write-Warn "Microsoft Graph service principal not found."
                return
            }
            $type = (Get-ArgValue $parsed.Map "type")
            $filter = (Get-ArgValue $parsed.Map "filter")
            $items = @()
            if (-not $type -or $type -eq "delegated" -or $type -eq "scope") {
                foreach ($s in @($graphSp.Oauth2PermissionScopes)) {
                    if ($filter -and ($s.Value -notlike ("*" + $filter + "*"))) { continue }
                    $items += [pscustomobject]@{
                        Type  = "Delegated"
                        Value = $s.Value
                        Id    = $s.Id
                    }
                }
            }
            if (-not $type -or $type -eq "application" -or $type -eq "role") {
                foreach ($r in @($graphSp.AppRoles)) {
                    if ($filter -and ($r.Value -notlike ("*" + $filter + "*"))) { continue }
                    $items += [pscustomobject]@{
                        Type  = "Application"
                        Value = $r.Value
                        Id    = $r.Id
                    }
                }
            }
            if (-not $items) {
                Write-Info "No permissions found."
                return
            }
            $items | Sort-Object Type, Value | Format-Table -AutoSize
        }
        "req" {
            if (-not (Require-GraphConnection)) { return }
            $method = $parsed.Positionals | Select-Object -First 1
            $path = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $method -or -not $path) {
                Write-Warn "Usage: graph req <get|post|patch|put|delete> <path|url> [--body <json>] [--bodyFile <path>] [--headers <json>] [--beta|--v1|--auto] [--contentType <type>] [--out <file>]"
                return
            }
            $method = $method.ToUpperInvariant()
            $bodyRaw = Get-ArgValue $parsed.Map "body"
            $bodyFile = Get-ArgValue $parsed.Map "bodyFile"
            $headersRaw = Get-ArgValue $parsed.Map "headers"
            $contentType = Get-ArgValue $parsed.Map "contentType"
            $out = Get-ArgValue $parsed.Map "out"
            $useBeta = $parsed.Map.ContainsKey("beta")
            $useV1 = $parsed.Map.ContainsKey("v1")
            $useAuto = $parsed.Map.ContainsKey("auto")

            $headers = @{}
            if ($headersRaw) {
                $hObj = Parse-Value $headersRaw
                if ($hObj -is [hashtable]) {
                    $headers = $hObj
                } elseif ($hObj) {
                    foreach ($p in $hObj.PSObject.Properties) {
                        $headers[$p.Name] = $p.Value
                    }
                }
            }

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

            if ($out -and $method -eq "GET") {
                Invoke-GraphDownload -Uri $path -OutFile $out -Headers $headers -Beta:$useBeta
                return
            }

            if ($useAuto -and -not $useBeta -and -not $useV1) {
                $resp = Invoke-GraphRequestAuto -Method $method -Uri $path -Body $body -Headers $headers -ContentType $contentType -AllowFallback
            } elseif ($useBeta) {
                $resp = Invoke-GraphRequest -Method $method -Uri $path -Body $body -Headers $headers -ContentType $contentType -Beta
            } else {
                $resp = Invoke-GraphRequest -Method $method -Uri $path -Body $body -Headers $headers -ContentType $contentType
            }
            if ($out) {
                if ($resp -is [string]) {
                    Set-Content -Path $out -Value $resp -Encoding ASCII
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 10 | Set-Content -Path $out -Encoding ASCII
                }
                Write-Info ("Saved: " + $out)
            } elseif ($resp) {
                if ($resp.value) {
                    $resp.value | ConvertTo-Json -Depth 10
                } else {
                    $resp | ConvertTo-Json -Depth 10
                }
            }
        }
        "meta" {
            $action = if ($rest.Count -gt 0) { $rest[0].ToLowerInvariant() } else { "" }
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest2
            $useBeta = $parsed2.Map.ContainsKey("beta")
            $useV1 = $parsed2.Map.ContainsKey("v1")
            $filter = Get-ArgValue $parsed2.Map "filter"
            $type = Get-ArgValue $parsed2.Map "type"

            if (-not $action) {
                Write-Warn "Usage: graph meta sync|list|show|paths|diff"
                return
            }

            switch ($action) {
                "sync" {
                    $force = $parsed2.Map.ContainsKey("force")
                    if ($useBeta -and -not $useV1) {
                        Sync-GraphMetadata -Beta -Force:$force
                    } elseif ($useV1 -and -not $useBeta) {
                        Sync-GraphMetadata -Force:$force
                    } else {
                        Sync-GraphMetadata -Force:$force
                        Sync-GraphMetadata -Beta -Force:$force
                    }
                }
                "list" {
                    $idx = Get-GraphMetadataIndex -Beta:$useBeta
                    if (-not $idx) { return }
                    $kind = if ($type) { $type.ToLowerInvariant() } else { "entityset" }
                    $items = @()
                    switch ($kind) {
                        "entity" { $items = $idx.EntityTypes | Select-Object FullName, BaseType }
                        "action" { $items = $idx.Actions | Select-Object FullName, IsBound, ReturnType }
                        "function" { $items = $idx.Functions | Select-Object FullName, IsBound, ReturnType }
                        "enum" { $items = $idx.Enums | Select-Object FullName }
                        "complex" { $items = $idx.ComplexTypes | Select-Object FullName }
                        default { $items = $idx.EntitySets | Select-Object Name, EntityType }
                    }
                    if ($filter) {
                        $items = $items | Where-Object { $_.Name -like ("*" + $filter + "*") -or $_.FullName -like ("*" + $filter + "*") }
                    }
                    if (-not $items) {
                        Write-Info "No metadata items found."
                        return
                    }
                    $items | Format-Table -AutoSize
                }
                "show" {
                    $name = $parsed2.Positionals | Select-Object -First 1
                    if (-not $name) {
                        Write-Warn "Usage: graph meta show <name>"
                        return
                    }
                    $idx = Get-GraphMetadataIndex -Beta:$useBeta
                    if (-not $idx) { return }
                    $match = $idx.EntitySets | Where-Object { $_.Name -eq $name } | Select-Object -First 1
                    if ($match) {
                        Write-Host ("EntitySet: " + $match.Name)
                        Write-Host ("EntityType: " + $match.EntityType)
                        return
                    }
                    $etype = $idx.EntityTypes | Where-Object { $_.FullName -eq $name -or $_.Name -eq $name } | Select-Object -First 1
                    if ($etype) {
                        Write-Host ("EntityType: " + $etype.FullName)
                        if ($etype.BaseType) { Write-Host ("BaseType: " + $etype.BaseType) }
                        if ($etype.Properties) {
                            Write-Host "Properties:"
                            $etype.Properties | Select-Object Name, Type, Nullable | Format-Table -AutoSize
                        }
                        if ($etype.Navigation) {
                            Write-Host "Navigation:"
                            $etype.Navigation | Select-Object Name, Type, ContainsTarget | Format-Table -AutoSize
                        }
                        return
                    }
                    $act = $idx.Actions | Where-Object { $_.FullName -eq $name -or $_.Name -eq $name } | Select-Object -First 1
                    if ($act) {
                        Write-Host ("Action: " + $act.FullName)
                        Write-Host ("IsBound: " + $act.IsBound)
                        if ($act.ReturnType) { Write-Host ("ReturnType: " + $act.ReturnType) }
                        if ($act.Parameters) {
                            Write-Host "Parameters:"
                            $act.Parameters | Select-Object Name, Type, Nullable | Format-Table -AutoSize
                        }
                        return
                    }
                    $fn = $idx.Functions | Where-Object { $_.FullName -eq $name -or $_.Name -eq $name } | Select-Object -First 1
                    if ($fn) {
                        Write-Host ("Function: " + $fn.FullName)
                        Write-Host ("IsBound: " + $fn.IsBound)
                        if ($fn.ReturnType) { Write-Host ("ReturnType: " + $fn.ReturnType) }
                        if ($fn.Parameters) {
                            Write-Host "Parameters:"
                            $fn.Parameters | Select-Object Name, Type, Nullable | Format-Table -AutoSize
                        }
                        return
                    }
                    Write-Warn "No metadata item found for name."
                }
                "paths" {
                    $name = $parsed2.Positionals | Select-Object -First 1
                    if (-not $name) {
                        Write-Warn "Usage: graph meta paths <name>"
                        return
                    }
                    $idx = Get-GraphMetadataIndex -Beta:$useBeta
                    if (-not $idx) { return }
                    $set = $idx.EntitySets | Where-Object { $_.Name -eq $name } | Select-Object -First 1
                    if ($set) {
                        Write-Host ("/" + $set.Name)
                        return
                    }
                    $etype = $idx.EntityTypes | Where-Object { $_.FullName -eq $name -or $_.Name -eq $name } | Select-Object -First 1
                    if ($etype) {
                        $sets = $idx.EntitySets | Where-Object { $_.EntityType -eq $etype.FullName }
                        if ($sets) {
                            foreach ($s in $sets) {
                                Write-Host ("/" + $s.Name)
                            }
                            return
                        }
                    }
                    $act = $idx.Actions | Where-Object { $_.FullName -eq $name -or $_.Name -eq $name } | Select-Object -First 1
                    if ($act) {
                        Write-Host ("/" + $act.FullName)
                        if ($act.IsBound -eq "true") { Write-Host "Bound action: call on entity or collection path." }
                        return
                    }
                    $fn = $idx.Functions | Where-Object { $_.FullName -eq $name -or $_.Name -eq $name } | Select-Object -First 1
                    if ($fn) {
                        Write-Host ("/" + $fn.FullName)
                        if ($fn.IsBound -eq "true") { Write-Host "Bound function: call on entity or collection path." }
                        return
                    }
                    Write-Warn "No metadata item found for name."
                }
                "diff" {
                    Sync-GraphMetadataIfNeeded
                    $kind = Get-ArgValue $parsed2.Map "type"
                    $filter = Get-ArgValue $parsed2.Map "filter"
                    $v1only = $parsed2.Map.ContainsKey("v1only")
                    $format = Get-ArgValue $parsed2.Map "format"
                    $topRaw = Get-ArgValue $parsed2.Map "top"
                    $items = Compare-GraphMetadata -Kind $kind -V1Only:$v1only
                    if (-not $items) {
                        Write-Info "No differences found."
                        return
                    }
                    if ($filter) {
                        $items = $items | Where-Object {
                            $_.Name -like ("*" + $filter + "*") -or
                            $_.EntityType -like ("*" + $filter + "*")
                        }
                    }
                    if ($topRaw) {
                        try {
                            $top = [int]$topRaw
                            if ($top -gt 0) { $items = $items | Select-Object -First $top }
                        } catch {}
                    }
                    $asJson = $parsed2.Map.ContainsKey("json") -or ($format -and $format.ToLowerInvariant() -eq "json")
                    if ($asJson) {
                        $items | ConvertTo-Json -Depth 8
                    } else {
                        $items | Format-Table -AutoSize
                    }
                }
                default {
                    Write-Warn "Usage: graph meta sync|list|show|paths|diff"
                }
            }
        }
        "list" {
            if (-not (Require-GraphConnection)) { return }
            $path = $parsed.Positionals | Select-Object -First 1
            if (-not $path) {
                Write-Warn "Usage: graph list <path> [--filter <odata>] [--top <n>] [--select ...] [--orderby ...] [--search text] [--beta]"
                return
            }
            $useBeta = $parsed.Map.ContainsKey("beta")
            $useV1 = $parsed.Map.ContainsKey("v1")
            $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
            if ($parsed.Map.ContainsKey("auto")) { $allowFallback = $true }
            if ($useBeta -or $useV1) { $allowFallback = $false }
            $qh = Build-QueryAndHeaders $parsed.Map @()
            $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($path + $qh.Query) -Headers $qh.Headers -Api $api -AllowFallback:$allowFallback
            if ($resp) {
                if ($resp.value) {
                    $resp.value | ConvertTo-Json -Depth 10
                } else {
                    $resp | ConvertTo-Json -Depth 10
                }
            }
        }
        "get" {
            if (-not (Require-GraphConnection)) { return }
            $path = $parsed.Positionals | Select-Object -First 1
            if (-not $path) {
                Write-Warn "Usage: graph get <path> [--select ...] [--expand ...] [--beta]"
                return
            }
            $useBeta = $parsed.Map.ContainsKey("beta")
            $useV1 = $parsed.Map.ContainsKey("v1")
            $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
            if ($parsed.Map.ContainsKey("auto")) { $allowFallback = $true }
            if ($useBeta -or $useV1) { $allowFallback = $false }
            $qh = Build-QueryAndHeaders $parsed.Map @()
            $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($path + $qh.Query) -Headers $qh.Headers -Api $api -AllowFallback:$allowFallback
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "create" {
            if (-not (Require-GraphConnection)) { return }
            $path = $parsed.Positionals | Select-Object -First 1
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            $bodyFile = Get-ArgValue $parsed.Map "bodyFile"
            if (-not $path -or (-not $jsonRaw -and -not $bodyFile)) {
                Write-Warn "Usage: graph create <path> --json <payload> [--beta|--auto] [--bodyFile <file>]"
                return
            }
            $body = $null
            if ($bodyFile) {
                if (-not (Test-Path $bodyFile)) {
                    Write-Warn "Body file not found."
                    return
                }
                $body = Get-Content -Raw -Path $bodyFile
            } else {
                $body = Parse-Value $jsonRaw
            }
            $useBeta = $parsed.Map.ContainsKey("beta")
            $useV1 = $parsed.Map.ContainsKey("v1")
            $allowFallback = $parsed.Map.ContainsKey("auto")
            $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
            $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $path -Body $body -Api $api -AllowFallback:$allowFallback
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "update" {
            if (-not (Require-GraphConnection)) { return }
            $path = $parsed.Positionals | Select-Object -First 1
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            $bodyFile = Get-ArgValue $parsed.Map "bodyFile"
            if (-not $path -or (-not $jsonRaw -and -not $bodyFile)) {
                Write-Warn "Usage: graph update <path> --json <payload> [--beta|--auto] [--bodyFile <file>]"
                return
            }
            $body = $null
            if ($bodyFile) {
                if (-not (Test-Path $bodyFile)) {
                    Write-Warn "Body file not found."
                    return
                }
                $body = Get-Content -Raw -Path $bodyFile
            } else {
                $body = Parse-Value $jsonRaw
            }
            $useBeta = $parsed.Map.ContainsKey("beta")
            $useV1 = $parsed.Map.ContainsKey("v1")
            $allowFallback = $parsed.Map.ContainsKey("auto")
            $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
            $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri $path -Body $body -Api $api -AllowFallback:$allowFallback
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "delete" {
            if (-not (Require-GraphConnection)) { return }
            $path = $parsed.Positionals | Select-Object -First 1
            if (-not $path) {
                Write-Warn "Usage: graph delete <path> [--force] [--beta|--auto]"
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
            $useBeta = $parsed.Map.ContainsKey("beta")
            $useV1 = $parsed.Map.ContainsKey("v1")
            $allowFallback = $parsed.Map.ContainsKey("auto")
            $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
            $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri $path -Api $api -AllowFallback:$allowFallback -AllowNullResponse
            if ($resp -ne $null) { Write-Info "Deleted." } else { Write-Info "Delete requested." }
        }
        "action" {
            if (-not (Require-GraphConnection)) { return }
            $path = $parsed.Positionals | Select-Object -First 1
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            $bodyFile = Get-ArgValue $parsed.Map "bodyFile"
            if (-not $path -or (-not $jsonRaw -and -not $bodyFile)) {
                Write-Warn "Usage: graph action <path> --json <payload> [--beta|--auto] [--bodyFile <file>]"
                return
            }
            $body = $null
            if ($bodyFile) {
                if (-not (Test-Path $bodyFile)) {
                    Write-Warn "Body file not found."
                    return
                }
                $body = Get-Content -Raw -Path $bodyFile
            } else {
                $body = Parse-Value $jsonRaw
            }
            $useBeta = $parsed.Map.ContainsKey("beta")
            $useV1 = $parsed.Map.ContainsKey("v1")
            $allowFallback = $parsed.Map.ContainsKey("auto")
            $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
            $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $path -Body $body -Api $api -AllowFallback:$allowFallback
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "batch" {
            if (-not (Require-GraphConnection)) { return }
            $file = Get-ArgValue $parsed.Map "file"
            if (-not $file) {
                Write-Warn "Usage: graph batch --file <json> [--beta|--auto]"
                return
            }
            if (-not (Test-Path $file)) {
                Write-Warn "Batch file not found."
                return
            }
            $raw = Get-Content -Raw -Path $file
            $body = Parse-Value $raw
            if (-not $body) { $body = $raw }
            $useBeta = $parsed.Map.ContainsKey("beta")
            $useV1 = $parsed.Map.ContainsKey("v1")
            $allowFallback = $parsed.Map.ContainsKey("auto")
            $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
            $resp = Invoke-GraphRequestAuto -Method "POST" -Uri "/`$batch" -Body $body -Api $api -AllowFallback:$allowFallback
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        default {
            Write-Warn "Usage: graph cmdlets|perms|req|meta|get|list|create|update|delete|action|batch"
        }
    }
}


