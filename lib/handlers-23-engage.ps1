# Handler: Engage
# Purpose: Engage command handlers.
function Get-EngageTokenPath {
    return (Join-Path $Paths.Data "engage.token")
}

function Get-EngageToken {
    $path = Get-EngageTokenPath
    if (Test-Path $path) {
        try {
            return (Get-Content -Raw -Path $path).Trim()
        } catch {}
    }
    return $null
}

function Save-EngageToken {
    param([string]$Token)
    Ensure-Directories
    $path = Get-EngageTokenPath
    Set-Content -Path $path -Value $Token -Encoding ASCII
}

function Clear-EngageToken {
    $path = Get-EngageTokenPath
    if (Test-Path $path) {
        Remove-Item -Path $path -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-EngageRequest {
    param(
        [string]$Method,
        [string]$Path,
        [object]$Body,
        [string]$Token,
        [string]$BaseUrl
    )
    if (-not $Token) { $Token = Get-EngageToken }
    if (-not $Token) {
        Write-Warn "Engage token missing. Use: engage token set --value <token>"
        return $null
    }
    $base = if ($BaseUrl) { $BaseUrl.TrimEnd("/") } else { "https://www.yammer.com" }
    $url = if ($Path -match "^https?://") { $Path } else { $base + $Path }
    $headers = @{ Authorization = ("Bearer " + $Token) }
    $params = @{ Method = $Method; Uri = $url; Headers = $headers }
    if ($Body -ne $null) {
        $params.ContentType = "application/json"
        if ($Body -is [string]) {
            $params.Body = $Body
        } else {
            $params.Body = ($Body | ConvertTo-Json -Depth 10)
        }
    }
    try {
        return Invoke-RestMethod @params
    } catch {
        Write-Err $_.Exception.Message
        return $null
    }
}

function Get-CommunityGroupId {
    param(
        [string]$CommunityId,
        [string]$Api,
        [switch]$AllowFallback
    )
    if (-not $CommunityId) { return $null }
    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ("/employeeExperience/communities/" + $CommunityId + "/group") -Api $Api -AllowFallback:$AllowFallback
    if ($resp -and $resp.id) { return $resp.id }
    return $null
}

function Invoke-EngageFormRequest {
    param(
        [string]$Method,
        [string]$Path,
        [hashtable]$Form,
        [string]$Token,
        [string]$BaseUrl
    )
    if (-not $Token) { $Token = Get-EngageToken }
    if (-not $Token) {
        Write-Warn "Engage token missing. Use: engage token set --value <token>"
        return $null
    }
    $base = if ($BaseUrl) { $BaseUrl.TrimEnd("/") } else { "https://www.yammer.com" }
    $url = if ($Path -match "^https?://") { $Path } else { $base + $Path }
    $headers = @{ Authorization = ("Bearer " + $Token) }
    $params = @{ Method = $Method; Uri = $url; Headers = $headers; ContentType = "application/x-www-form-urlencoded" }
    if ($Form) { $params.Body = $Form }
    try {
        return Invoke-RestMethod @params
    } catch {
        Write-Err $_.Exception.Message
        return $null
    }
}

function Handle-EngageMessageCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: engage message list|post|delete ..."
        return
    }
    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($action) {
        "list" {
            $group = Get-ArgValue $parsed.Map "group"
            $thread = Get-ArgValue $parsed.Map "thread"
            $topic = Get-ArgValue $parsed.Map "topic"
            $og = Get-ArgValue $parsed.Map "opengraph"
            $feed = Get-ArgValue $parsed.Map "feed"
            $limit = Get-ArgValue $parsed.Map "limit"
            if (-not $limit) { $limit = Get-ArgValue $parsed.Map "top" }
            $older = Get-ArgValue $parsed.Map "older"
            $newer = Get-ArgValue $parsed.Map "newer"
            $since = Get-ArgValue $parsed.Map "since"
            $threaded = Get-ArgValue $parsed.Map "threaded"

            $path = "/api/v1/messages.json"
            if ($group) {
                $path = "/api/v1/messages/in_group/" + $group + ".json"
            } elseif ($thread) {
                $path = "/api/v1/messages/in_thread/" + $thread + ".json"
            } elseif ($topic) {
                $path = "/api/v1/messages/about_topic/" + $topic + ".json"
            } elseif ($og) {
                $path = "/api/v1/messages/open_graph_objects/" + $og + ".json"
            } elseif ($feed) {
                switch ($feed.ToLowerInvariant()) {
                    "my" { $path = "/api/v1/messages/my_feed.json" }
                    "following" { $path = "/api/v1/messages/my_feed.json" }
                    "sent" { $path = "/api/v1/messages/sent.json" }
                    "received" { $path = "/api/v1/messages/received.json" }
                    "private" { $path = "/api/v1/messages/private.json" }
                    "algo" { $path = "/api/v1/messages/algo.json" }
                    default { $path = "/api/v1/messages.json" }
                }
            }

            $query = @()
            if ($limit) { $query += "limit=" + (Encode-QueryValue $limit) }
            if ($older) { $query += "older_than=" + (Encode-QueryValue $older) }
            if ($newer) { $query += "newer_than=" + (Encode-QueryValue $newer) }
            if ($since) { $query += "since_id=" + (Encode-QueryValue $since) }
            if ($threaded) { $query += "threaded=" + (Encode-QueryValue $threaded) }
            if ($query.Count -gt 0) { $path = $path + "?" + ($query -join "&") }

            $resp = Invoke-EngageRequest -Method "GET" -Path $path
            if ($resp) { $resp | ConvertTo-Json -Depth 8 }
        }
        "post" {
            $bodyText = Get-ArgValue $parsed.Map "body"
            $bodyFile = Get-ArgValue $parsed.Map "bodyFile"
            if (-not $bodyText -and $bodyFile) {
                if (Test-Path $bodyFile) { $bodyText = (Get-Content -Raw -Path $bodyFile) }
            }
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            if ($jsonRaw) {
                $payload = Parse-Value $jsonRaw
                $resp = Invoke-EngageRequest -Method "POST" -Path "/api/v1/messages.json" -Body $payload
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                return
            }
            if (-not $bodyText) {
                Write-Warn "Usage: engage message post --body <text> [--group <id>] [--reply <id>] [--directTo <id1,id2>] [--network <id>]"
                return
            }
            $group = Get-ArgValue $parsed.Map "group"
            $reply = Get-ArgValue $parsed.Map "reply"
            if (-not $reply) { $reply = Get-ArgValue $parsed.Map "repliedTo" }
            $direct = Get-ArgValue $parsed.Map "directTo"
            $network = Get-ArgValue $parsed.Map "network"
            $form = @{ body = $bodyText }
            if ($group) { $form.group_id = $group }
            if ($reply) { $form.replied_to_id = $reply }
            if ($direct) {
                $ids = Parse-CommaList $direct
                if ($ids.Count -gt 0) { $form.direct_to_user_ids = ($ids -join ",") }
            }
            if ($network) { $form.network_id = $network }
            $resp = Invoke-EngageFormRequest -Method "POST" -Path "/api/v1/messages.json" -Form $form
            if ($resp) { $resp | ConvertTo-Json -Depth 8 }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: engage message delete <messageId> [--force]"
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
            $resp = Invoke-EngageRequest -Method "DELETE" -Path ("/api/v1/messages/" + $id + ".json")
            if ($resp -ne $null) { Write-Info "Message deleted." }
        }
        default {
            Write-Warn "Usage: engage message list|post|delete ..."
        }
    }
}

function Handle-EngageCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: engage open|info|token|community|message|raw"
        return
    }
    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }

    switch ($sub) {
        "open" { Write-Host "https://engage.cloud.microsoft/" }
        "info" {
            Write-Warn "Viva Engage Graph APIs are preview/limited. Legacy REST APIs require delegated tokens."
        }
        "token" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: engage token set --value <token> | engage token show | engage token clear"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed = Parse-NamedArgs $rest2
            switch ($action) {
                "set" {
                    $val = Get-ArgValue $parsed.Map "value"
                    $file = Get-ArgValue $parsed.Map "file"
                    if (-not $val -and $file) {
                        if (Test-Path $file) { $val = (Get-Content -Raw -Path $file) }
                    }
                    if (-not $val) {
                        Write-Warn "Usage: engage token set --value <token> OR --file <path>"
                        return
                    }
                    Save-EngageToken $val
                    Write-Info "Engage token saved."
                }
                "show" {
                    $tok = Get-EngageToken
                    if ($tok) { Write-Host $tok } else { Write-Warn "No token saved." }
                }
                "clear" {
                    Clear-EngageToken
                    Write-Info "Engage token cleared."
                }
                default {
                    Write-Warn "Usage: engage token set|show|clear"
                }
            }
        }
        "community" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: engage community list|get|create|update|delete|owners|group|members [--beta|--auto]"
                return
            }
            if (-not (Require-GraphConnection)) { return }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed = Parse-NamedArgs $rest2
            $useBeta = $parsed.Map.ContainsKey("beta")
            $useV1 = $parsed.Map.ContainsKey("v1")
            $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
            if ($parsed.Map.ContainsKey("auto")) { $allowFallback = $true }
            if ($useBeta -or $useV1) { $allowFallback = $false }
            $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
            $base = "/employeeExperience/communities"

            switch ($action) {
                "list" {
                    $qh = Build-QueryAndHeaders $parsed.Map @("id","displayName","description","visibility")
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Api $api -AllowFallback:$allowFallback
                    if ($resp) {
                        if ($resp.value) { $resp.value | ConvertTo-Json -Depth 8 } else { $resp | ConvertTo-Json -Depth 8 }
                    }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: engage community get <id>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                }
                "create" {
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $body) {
                        Write-Warn "Usage: engage community create --json <payload>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $api -AllowFallback:$allowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: engage community update <id> --json <payload> OR --set key=value"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $api -AllowFallback:$allowFallback
                    if ($resp -ne $null) { Write-Info "Community updated." }
                }
                "owners" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: engage community owners <id>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id + "/owners") -Api $api -AllowFallback:$allowFallback
                    if ($resp) {
                        if ($resp.value) { $resp.value | ConvertTo-Json -Depth 8 } else { $resp | ConvertTo-Json -Depth 8 }
                    }
                }
                "group" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: engage community group <id>"
                        return
                    }
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id + "/group") -Api $api -AllowFallback:$allowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                }
                "members" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: engage community members <id>"
                        return
                    }
                    $groupId = Get-CommunityGroupId -CommunityId $id -Api $api -AllowFallback:$allowFallback
                    if (-not $groupId) {
                        Write-Warn "Community group not found."
                        return
                    }
                    $qh = Build-QueryAndHeaders $parsed.Map @("id","displayName","userPrincipalName")
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/groups/" + $groupId + "/members" + $qh.Query) -Headers $qh.Headers
                    if ($resp -and $resp.value) {
                        $resp.value | ConvertTo-Json -Depth 8
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 8
                    }
                }
                "member" {
                    if (-not $rest2 -or $rest2.Count -lt 2) {
                        Write-Warn "Usage: engage community member add|remove <communityId> --user <upn|id>"
                        return
                    }
                    $memberAction = $rest2[0].ToLowerInvariant()
                    $communityId = $rest2 | Select-Object -Skip 1 -First 1
                    $user = Get-ArgValue $parsed.Map "user"
                    if (-not $user) { $user = Get-ArgValue $parsed.Map "upn" }
                    if (-not $user) { $user = $parsed.Positionals | Select-Object -Skip 1 -First 1 }
                    if (-not $communityId -or -not $user) {
                        Write-Warn "Usage: engage community member add|remove <communityId> --user <upn|id>"
                        return
                    }
                    $groupId = Get-CommunityGroupId -CommunityId $communityId -Api $api -AllowFallback:$allowFallback
                    if (-not $groupId) {
                        Write-Warn "Community group not found."
                        return
                    }
                    $u = Resolve-UserObject $user
                    if (-not $u -or -not $u.Id) {
                        Write-Warn "User not found."
                        return
                    }
                    if ($memberAction -eq "add") {
                        $body = @{ "@odata.id" = ("https://graph.microsoft.com/v1.0/directoryObjects/" + $u.Id) }
                        $resp = Invoke-GraphRequest -Method "POST" -Uri ("/groups/" + $groupId + "/members/`$ref") -Body $body
                        if ($resp -ne $null) { Write-Info "Member added." }
                    } elseif ($memberAction -eq "remove") {
                        $resp = Invoke-GraphRequest -Method "DELETE" -Uri ("/groups/" + $groupId + "/members/" + $u.Id + "/`$ref")
                        if ($resp -ne $null) { Write-Info "Member removed." }
                    } else {
                        Write-Warn "Usage: engage community member add|remove <communityId> --user <upn|id>"
                    }
                }
                "owner" {
                    if (-not $rest2 -or $rest2.Count -lt 2) {
                        Write-Warn "Usage: engage community owner add|remove <communityId> --user <upn|id>"
                        return
                    }
                    $ownerAction = $rest2[0].ToLowerInvariant()
                    $communityId = $rest2 | Select-Object -Skip 1 -First 1
                    $user = Get-ArgValue $parsed.Map "user"
                    if (-not $user) { $user = Get-ArgValue $parsed.Map "upn" }
                    if (-not $user) { $user = $parsed.Positionals | Select-Object -Skip 1 -First 1 }
                    if (-not $communityId -or -not $user) {
                        Write-Warn "Usage: engage community owner add|remove <communityId> --user <upn|id>"
                        return
                    }
                    $groupId = Get-CommunityGroupId -CommunityId $communityId -Api $api -AllowFallback:$allowFallback
                    if (-not $groupId) {
                        Write-Warn "Community group not found."
                        return
                    }
                    $u = Resolve-UserObject $user
                    if (-not $u -or -not $u.Id) {
                        Write-Warn "User not found."
                        return
                    }
                    if ($ownerAction -eq "add") {
                        $body = @{ "@odata.id" = ("https://graph.microsoft.com/v1.0/directoryObjects/" + $u.Id) }
                        $resp = Invoke-GraphRequest -Method "POST" -Uri ("/groups/" + $groupId + "/owners/`$ref") -Body $body
                        if ($resp -ne $null) { Write-Info "Owner added." }
                    } elseif ($ownerAction -eq "remove") {
                        $resp = Invoke-GraphRequest -Method "DELETE" -Uri ("/groups/" + $groupId + "/owners/" + $u.Id + "/`$ref")
                        if ($resp -ne $null) { Write-Info "Owner removed." }
                    } else {
                        Write-Warn "Usage: engage community owner add|remove <communityId> --user <upn|id>"
                    }
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: engage community delete <id> [--force]"
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
                    $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                    if ($resp -ne $null) { Write-Info "Community deleted." }
                }
                default {
                    Write-Warn "Usage: engage community list|get|create|update|delete|owners|group|members [--beta|--auto]"
                }
            }
        }
        "message" { Handle-EngageMessageCommand $rest }
        "raw" {
            if (-not $rest -or $rest.Count -lt 2) {
                Write-Warn "Usage: engage raw <get|post|patch|put|delete> <path> [--json <payload>] [--bodyFile <file>] [--token <token>] [--base <url>] [--out <file>]"
                return
            }
            $method = $rest[0].ToUpperInvariant()
            $path = $rest[1]
            $rest2 = if ($rest.Count -gt 2) { $rest[2..($rest.Count - 1)] } else { @() }
            $parsed = Parse-NamedArgs $rest2
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") $null
            $token = Get-ArgValue $parsed.Map "token"
            $base = Get-ArgValue $parsed.Map "base"
            $out = Get-ArgValue $parsed.Map "out"
            $resp = Invoke-EngageRequest -Method $method -Path $path -Body $body -Token $token -BaseUrl $base
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
        default {
            Write-Warn "Usage: engage open|info|token|community|message|raw"
        }
    }
}
