# Handler: Copilot
# Purpose: Copilot command handlers.
function Get-CopilotTimeZone {
    $tz = [System.TimeZoneInfo]::Local
    if ($tz -and $tz.Id) { return $tz.Id }
    return "UTC"
}

function New-CopilotContextMessage {
    param(
        [string]$Text,
        [string]$Description
    )
    $msg = @{ text = $Text }
    if ($Description) { $msg.description = $Description }
    return $msg
}

function Truncate-Text {
    param(
        [string]$Text,
        [int]$Max = 2000
    )
    if (-not $Text) { return $Text }
    if ($Max -le 0) { return $Text }
    if ($Text.Length -le $Max) { return $Text }
    return ($Text.Substring(0, $Max) + "...")
}

function Get-CopilotMailContext {
    param(
        [string]$MessageId,
        [string]$UserSegment
    )
    if (-not $MessageId) { return $null }
    if (-not $UserSegment) { $UserSegment = "/me" }
    $select = "?`$select=subject,bodyPreview,from,toRecipients,ccRecipients,receivedDateTime"
    $resp = Invoke-GraphRequest -Method "GET" -Uri ($UserSegment + "/messages/" + $MessageId + $select)
    if (-not $resp) { return $null }
    $from = if ($resp.from -and $resp.from.emailAddress) { $resp.from.emailAddress.address } else { "" }
    $to = @()
    foreach ($r in @($resp.toRecipients)) {
        if ($r.emailAddress) { $to += $r.emailAddress.address }
    }
    $cc = @()
    foreach ($r in @($resp.ccRecipients)) {
        if ($r.emailAddress) { $cc += $r.emailAddress.address }
    }
    $lines = @()
    if ($resp.subject) { $lines += ("Subject: " + $resp.subject) }
    if ($from) { $lines += ("From: " + $from) }
    if ($to.Count -gt 0) { $lines += ("To: " + ($to -join ", ")) }
    if ($cc.Count -gt 0) { $lines += ("Cc: " + ($cc -join ", ")) }
    if ($resp.receivedDateTime) { $lines += ("Received: " + $resp.receivedDateTime) }
    if ($resp.bodyPreview) { $lines += ("Preview: " + $resp.bodyPreview) }
    return ($lines -join "`n")
}

function Get-CopilotEventContext {
    param(
        [string]$EventId,
        [string]$UserSegment
    )
    if (-not $EventId) { return $null }
    if (-not $UserSegment) { $UserSegment = "/me" }
    $select = "?`$select=subject,bodyPreview,start,end,location,organizer,attendees"
    $resp = Invoke-GraphRequest -Method "GET" -Uri ($UserSegment + "/events/" + $EventId + $select)
    if (-not $resp) { return $null }
    $lines = @()
    if ($resp.subject) { $lines += ("Subject: " + $resp.subject) }
    if ($resp.organizer -and $resp.organizer.emailAddress) { $lines += ("Organizer: " + $resp.organizer.emailAddress.address) }
    if ($resp.start -and $resp.start.dateTime) { $lines += ("Start: " + $resp.start.dateTime + " " + $resp.start.timeZone) }
    if ($resp.end -and $resp.end.dateTime) { $lines += ("End: " + $resp.end.dateTime + " " + $resp.end.timeZone) }
    if ($resp.location -and $resp.location.displayName) { $lines += ("Location: " + $resp.location.displayName) }
    if ($resp.attendees) {
        $att = @()
        foreach ($a in @($resp.attendees)) {
            if ($a.emailAddress) { $att += $a.emailAddress.address }
        }
        if ($att.Count -gt 0) { $lines += ("Attendees: " + ($att -join ", ")) }
    }
    if ($resp.bodyPreview) { $lines += ("Preview: " + $resp.bodyPreview) }
    return ($lines -join "`n")
}

function Get-CopilotPersonContext {
    param([string]$Identity)
    if (-not $Identity) { return $null }
    $user = $null
    if ($Identity -match "^[0-9a-fA-F-]{36}$" -or $Identity -match "@") {
        $user = Invoke-GraphRequest -Method "GET" -Uri ("/users/" + $Identity + "?`$select=displayName,jobTitle,mail,department,officeLocation,mobilePhone")
    } else {
        $esc = Escape-ODataString $Identity
        $resp = Invoke-GraphRequest -Method "GET" -Uri ("/users?`$select=displayName,jobTitle,mail,department,officeLocation,mobilePhone&`$filter=displayName eq '" + $esc + "'")
        if ($resp -and $resp.value) { $user = $resp.value | Select-Object -First 1 }
    }
    if (-not $user) { return $null }
    $lines = @()
    if ($user.displayName) { $lines += ("Name: " + $user.displayName) }
    if ($user.jobTitle) { $lines += ("Title: " + $user.jobTitle) }
    if ($user.mail) { $lines += ("Mail: " + $user.mail) }
    if ($user.department) { $lines += ("Department: " + $user.department) }
    if ($user.officeLocation) { $lines += ("Office: " + $user.officeLocation) }
    if ($user.mobilePhone) { $lines += ("Mobile: " + $user.mobilePhone) }
    return ($lines -join "`n")
}

function Add-CopilotAdditionalContext {
    param(
        [hashtable]$Body,
        [object[]]$Items
    )
    if (-not $Items -or $Items.Count -eq 0) { return }
    if (-not $Body.additionalContext) {
        $Body.additionalContext = @()
    }
    $current = @($Body.additionalContext)
    $Body.additionalContext = @($current + $Items)
}


function New-CopilotConversation {
    $resp = Invoke-GraphRequest -Method "POST" -Uri "/copilot/conversations" -Body @{} -Beta
    if ($resp -and $resp.id) {
        $global:CopilotConversationId = $resp.id
    }
    return $resp
}


function Get-CopilotConversationId {
    param([string]$Explicit)
    if ($Explicit) { return $Explicit }
    if ($global:CopilotConversationId) { return $global:CopilotConversationId }
    return $null
}


function Extract-CopilotText {
    param([object]$Response)
    if (-not $Response) { return $null }
    $msgs = $Response.messages
    if (-not $msgs) { return $null }
    $last = $msgs | Select-Object -Last 1
    if ($last -and $last.text) { return $last.text }
    return $null
}

function Get-CopilotHitResource {
    param([object]$Hit)
    if (-not $Hit) { return $null }
    if ($Hit.resource) { return $Hit.resource }
    if ($Hit.resourceUrl -or $Hit.url) { return $Hit }
    return $Hit
}

function Get-DriveItemInfoFromHit {
    param([object]$Hit)
    $res = Get-CopilotHitResource $Hit
    if (-not $res) { return $null }
    $id = $res.id
    $driveId = $null
    if ($res.parentReference -and $res.parentReference.driveId) { $driveId = $res.parentReference.driveId }
    if (-not $driveId -and $res.driveId) { $driveId = $res.driveId }
    if (-not $driveId -and $res.drive -and $res.drive.id) { $driveId = $res.drive.id }
    $webUrl = $res.webUrl
    if (-not $webUrl -and $res.resourceUrl) { $webUrl = $res.resourceUrl }
    if (-not $webUrl -and $res.url) { $webUrl = $res.url }
    $name = $res.name
    return [pscustomobject]@{
        Id      = $id
        DriveId = $driveId
        WebUrl  = $webUrl
        Name    = $name
    }
}

function Get-FileUriFromHit {
    param([object]$Hit)
    $res = Get-CopilotHitResource $Hit
    if (-not $res) { return $null }
    if ($res.webUrl) { return $res.webUrl }
    if ($res.resourceUrl) { return $res.resourceUrl }
    if ($res.url) { return $res.url }
    return $null
}

function Resolve-IndexList {
    param(
        [string]$Raw,
        [int]$Max
    )
    $list = @()
    if (-not $Raw) { return $list }
    foreach ($r in (Parse-CommaList $Raw)) {
        try {
            $i = [int]$r
            if ($i -gt 0 -and $i -le $Max) { $list += ($i - 1) }
        } catch {}
    }
    return @($list | Select-Object -Unique)
}

function Add-CopilotFilesFromHits {
    param(
        [hashtable]$Body,
        [object[]]$Hits,
        [string]$IndexRaw
    )
    if (-not $Hits -or $Hits.Count -eq 0) { return }
    if (-not $IndexRaw) { return }
    $idxList = Resolve-IndexList -Raw $IndexRaw -Max $Hits.Count
    if ($idxList.Count -eq 0) { return }
    $uris = @()
    foreach ($i in $idxList) {
        $uri = Get-FileUriFromHit $Hits[$i]
        if ($uri) { $uris += $uri }
    }
    if ($uris.Count -eq 0) { return }
    if (-not $Body.contextualResources) { $Body.contextualResources = @{} }
    if (-not $Body.contextualResources.files) { $Body.contextualResources.files = @() }
    $existing = @($Body.contextualResources.files | ForEach-Object { $_.uri })
    foreach ($u in $uris) {
        if ($existing -notcontains $u) {
            $Body.contextualResources.files += @{ uri = $u }
        }
    }
}

function Add-CopilotFilesFromHitsTop {
    param(
        [hashtable]$Body,
        [object[]]$Hits,
        [string]$TopRaw
    )
    if (-not $Hits -or $Hits.Count -eq 0) { return }
    if (-not $TopRaw) { return }
    $top = 0
    try { $top = [int]$TopRaw } catch { $top = 0 }
    if ($top -le 0) { return }
    if ($top -gt $Hits.Count) { $top = $Hits.Count }
    $idxRaw = (1..$top) -join ","
    Add-CopilotFilesFromHits -Body $Body -Hits $Hits -IndexRaw $idxRaw
}

function Show-CopilotHits {
    param([object[]]$Hits)
    if (-not $Hits -or $Hits.Count -eq 0) {
        Write-Info "No hits to show."
        return
    }
    $rows = @()
    $i = 1
    foreach ($h in $Hits) {
        $info = Get-DriveItemInfoFromHit $h
        if (-not $info) { continue }
        $rows += [pscustomobject]@{
            Index = $i
            Name  = $info.Name
            WebUrl = $info.WebUrl
            Id    = $info.Id
        }
        $i++
    }
    $rows | Format-Table -AutoSize
}

function Resolve-CopilotHitIndex {
    param(
        [object[]]$Hits,
        [object]$Parsed
    )
    if (-not $Hits -or $Hits.Count -eq 0) { return -1 }
    $idxRaw = Get-ArgValue $Parsed.Map "index"
    if (-not $idxRaw) { $idxRaw = $Parsed.Positionals | Select-Object -First 1 }
    if (-not $idxRaw) { return -1 }
    try {
        $idx = [int]$idxRaw
        if ($idx -le 0) { return -1 }
        if ($idx -gt $Hits.Count) { return -1 }
        return ($idx - 1)
    } catch {
        return -1
    }
}

function Download-CopilotHit {
    param(
        [object]$Hit,
        [string]$OutFile
    )
    $info = Get-DriveItemInfoFromHit $Hit
    if (-not $info -or -not $info.Id -or -not $info.DriveId) {
        Write-Warn "DriveId/ItemId missing on hit; cannot download."
        return
    }
    if (-not $OutFile) {
        $OutFile = $info.Name
    }
    if (-not $OutFile) {
        Write-Warn "Output file name required."
        return
    }
    $uri = "/drives/" + $info.DriveId + "/items/" + $info.Id + "/content"
    Invoke-GraphDownload -Uri $uri -OutFile $OutFile
}

function Get-GraphAccessToken {
    $ctx = Get-MgContextSafe
    if ($ctx -and $ctx.AccessToken) { return $ctx.AccessToken }
    $cmd = Get-Command -Name Get-MgGraphAccessToken -ErrorAction SilentlyContinue
    if ($cmd) {
        try {
            $tok = Get-MgGraphAccessToken -ErrorAction Stop
            if ($tok -is [string]) { return $tok }
            if ($tok -and $tok.AccessToken) { return $tok.AccessToken }
        } catch {}
    }
    return $null
}

function Invoke-CopilotChatStreamLive {
    param(
        [string]$ConversationId,
        [object]$Body
    )
    $token = Get-GraphAccessToken
    if (-not $token) {
        Write-Warn "Could not obtain Graph access token for streaming. Try /login again."
        return $false
    }
    $url = "https://graph.microsoft.com/beta/copilot/conversations/" + $ConversationId + "/chatOverStream"
    $json = ($Body | ConvertTo-Json -Depth 8)
    $client = New-Object System.Net.Http.HttpClient
    $req = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Post, $url)
    $req.Headers.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $token)
    $req.Headers.Accept.Add([System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new("text/event-stream"))
    $req.Content = New-Object System.Net.Http.StringContent($json, [System.Text.Encoding]::UTF8, "application/json")
    $resp = $null
    $stream = $null
    $reader = $null
    try {
        $resp = $client.SendAsync($req, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
        if (-not $resp.IsSuccessStatusCode) {
            $err = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            Write-Err ("Copilot stream failed: HTTP " + [int]$resp.StatusCode + " " + $resp.ReasonPhrase + " " + $err)
            return $false
        }
        $stream = $resp.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
        $reader = New-Object System.IO.StreamReader($stream)
        $lastText = ""
        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            if (-not $line) { continue }
            if ($line.StartsWith("data:")) {
                $payload = $line.Substring(5).Trim()
                if (-not $payload -or $payload -eq "[DONE]") { continue }
                $obj = $null
                try { $obj = $payload | ConvertFrom-Json -ErrorAction Stop } catch { $obj = $null }
                if (-not $obj) { continue }
                $text = Extract-CopilotText $obj
                if (-not $text) { continue }
                if ($lastText -and $text.StartsWith($lastText)) {
                    $delta = $text.Substring($lastText.Length)
                    if ($delta) { Write-Host -NoNewline $delta }
                } else {
                    if ($lastText) { Write-Host "" }
                    Write-Host $text
                }
                $lastText = $text
            }
        }
        if ($lastText) { Write-Host "" }
        return $true
    } catch {
        Write-Err $_.Exception.Message
        return $false
    } finally {
        if ($reader) { try { $reader.Close() } catch {} }
        if ($stream) { try { $stream.Dispose() } catch {} }
        if ($resp) { try { $resp.Dispose() } catch {} }
        if ($client) { try { $client.Dispose() } catch {} }
    }
}

function Write-CopilotStreamText {
    param([string]$SseText)
    if (-not $SseText) { return }
    $lastText = ""
    $lines = $SseText -split "`n"
    foreach ($line in $lines) {
        $t = $line.Trim()
        if (-not $t) { continue }
        if ($t.StartsWith("data:")) {
            $payload = $t.Substring(5).Trim()
            if (-not $payload -or $payload -eq "[DONE]") { continue }
            $obj = $null
            try { $obj = $payload | ConvertFrom-Json -ErrorAction Stop } catch { $obj = $null }
            if (-not $obj) { continue }
            $text = Extract-CopilotText $obj
            if (-not $text) { continue }
            if ($lastText -and $text.StartsWith($lastText)) {
                $delta = $text.Substring($lastText.Length)
                if ($delta) { Write-Host -NoNewline $delta }
            } else {
                if ($lastText) { Write-Host "" }
                Write-Host $text
            }
            $lastText = $text
        }
    }
    if ($lastText) { Write-Host "" }
}


function Handle-CopilotChatCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: copilot chat create|send|ask|stream"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($action) {
        "create" {
            $resp = New-CopilotConversation
            if ($resp) { $resp | ConvertTo-Json -Depth 8 }
        }
        "stream" {
            Handle-CopilotChatCommand (@("send","--stream") + $rest)
        }
        "ask" {
            $text = Get-ArgValue $parsed.Map "text"
            if (-not $text) {
                $text = $parsed.Positionals | Select-Object -First 1
            }
            if (-not $text) {
                Write-Warn "Usage: copilot chat ask --text <message>"
                return
            }
            $conv = New-CopilotConversation
            if (-not $conv -or -not $conv.id) { return }
            $parsed.Map["id"] = $conv.id
            Handle-CopilotChatCommand (@("send") + $rest)
        }
        "send" {
            $id = Get-ArgValue $parsed.Map "id"
            if (-not $id) {
                $id = $parsed.Positionals | Select-Object -First 1
            }
            $id = Get-CopilotConversationId $id
            if (-not $id) {
                Write-Warn "Conversation id required. Use: copilot chat create OR copilot chat send --id <id> --text <msg>"
                return
            }

            $jsonRaw = Get-ArgValue $parsed.Map "json"
            $bodyFile = Get-ArgValue $parsed.Map "bodyFile"
            if ($jsonRaw -or $bodyFile) {
                $body = Read-JsonPayload $jsonRaw $bodyFile $null
            } else {
                $text = Get-ArgValue $parsed.Map "text"
                if (-not $text) { $text = ($parsed.Positionals | Select-Object -Skip 1 -First 1) }
                if (-not $text) {
                    Write-Warn "Usage: copilot chat send <conversationId> --text <message> [--files url1,url2]"
                    return
                }
                $tz = Get-ArgValue $parsed.Map "tz"
                if (-not $tz) { $tz = Get-ArgValue $parsed.Map "timezone" }
                if (-not $tz) { $tz = Get-CopilotTimeZone }

                $body = [ordered]@{
                    message      = @{ text = $text }
                    locationHint = @{ timeZone = $tz }
                }

            $files = Parse-CommaList (Get-ArgValue $parsed.Map "files")
            $fileSingle = Get-ArgValue $parsed.Map "file"
            if ($fileSingle) { $files += $fileSingle }
            if ($files.Count -gt 0) {
                if (-not $body.contextualResources) { $body.contextualResources = @{} }
                $body.contextualResources.files = @($files | ForEach-Object { @{ uri = $_ } })
            }

            $useSearch = Get-ArgValue $parsed.Map "usesearch"
            if ($useSearch) { Add-CopilotFilesFromHits -Body $body -Hits $global:CopilotLastSearchHits -IndexRaw $useSearch }
            $useSearchTop = Get-ArgValue $parsed.Map "usesearchtop"
            if ($useSearchTop) { Add-CopilotFilesFromHitsTop -Body $body -Hits $global:CopilotLastSearchHits -TopRaw $useSearchTop }
            $useRetrieve = Get-ArgValue $parsed.Map "useretrieve"
            if (-not $useRetrieve) { $useRetrieve = Get-ArgValue $parsed.Map "useretrieval" }
            if ($useRetrieve) { Add-CopilotFilesFromHits -Body $body -Hits $global:CopilotLastRetrievalHits -IndexRaw $useRetrieve }
            $useRetrieveTop = Get-ArgValue $parsed.Map "useretrievetop"
            if ($useRetrieveTop) { Add-CopilotFilesFromHitsTop -Body $body -Hits $global:CopilotLastRetrievalHits -TopRaw $useRetrieveTop }

            $webEnabled = Get-ArgValue $parsed.Map "web"
            if ($null -ne $webEnabled) {
                $we = Parse-Bool $webEnabled $true
                if (-not $body.contextualResources) { $body.contextualResources = @{} }
                $body.contextualResources.webContext = @{ isWebEnabled = $we }
                }

                $ctxJson = Get-ArgValue $parsed.Map "contextual"
                $ctxFile = Get-ArgValue $parsed.Map "contextualFile"
                if ($ctxJson -or $ctxFile) {
                    $ctx = Read-JsonPayload $ctxJson $ctxFile $null
                    if ($ctx) {
                        if (-not $body.contextualResources) { $body.contextualResources = @{} }
                        foreach ($p in $ctx.PSObject.Properties) {
                            $body.contextualResources[$p.Name] = $p.Value
                        }
                    }
                }

                $addJson = Get-ArgValue $parsed.Map "additional"
                $addFile = Get-ArgValue $parsed.Map "additionalFile"
                if ($addJson -or $addFile) {
                    $add = Read-JsonPayload $addJson $addFile $null
                    if ($add) {
                        if ($add -is [array]) {
                            Add-CopilotAdditionalContext -Body $body -Items $add
                        } else {
                            $body.additionalContext = $add
                        }
                    }
                }

                $ctxMaxRaw = Get-ArgValue $parsed.Map "ctxmax"
                $ctxMax = 2000
                if ($ctxMaxRaw) { try { $ctxMax = [int]$ctxMaxRaw } catch {} }

                $ctxItems = @()
                $ctxText = Parse-CommaList (Get-ArgValue $parsed.Map "ctx")
                foreach ($t in $ctxText) {
                    if ($t) { $ctxItems += (New-CopilotContextMessage -Text (Truncate-Text $t $ctxMax)) }
                }
                $ctxFilePath = Get-ArgValue $parsed.Map "ctxfile"
                if ($ctxFilePath -and (Test-Path $ctxFilePath)) {
                    $raw = Get-Content -Raw -Path $ctxFilePath
                    $ctxItems += (New-CopilotContextMessage -Text (Truncate-Text $raw $ctxMax) -Description ("local file: " + $ctxFilePath))
                }
                $mailIds = Parse-CommaList (Get-ArgValue $parsed.Map "mail")
                $eventIds = Parse-CommaList (Get-ArgValue $parsed.Map "event")
                $meetingIds = Parse-CommaList (Get-ArgValue $parsed.Map "meeting")
                $personIds = Parse-CommaList (Get-ArgValue $parsed.Map "person")

                $userSeg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
                if (-not $userSeg) { $userSeg = "/me" }

                foreach ($m in $mailIds) {
                    $text = Get-CopilotMailContext -MessageId $m -UserSegment $userSeg
                    if ($text) { $ctxItems += (New-CopilotContextMessage -Text (Truncate-Text $text $ctxMax) -Description ("mail:" + $m)) }
                }
                foreach ($e in $eventIds) {
                    $text = Get-CopilotEventContext -EventId $e -UserSegment $userSeg
                    if ($text) { $ctxItems += (New-CopilotContextMessage -Text (Truncate-Text $text $ctxMax) -Description ("event:" + $e)) }
                }
                foreach ($e in $meetingIds) {
                    $text = Get-CopilotEventContext -EventId $e -UserSegment $userSeg
                    if ($text) { $ctxItems += (New-CopilotContextMessage -Text (Truncate-Text $text $ctxMax) -Description ("meeting:" + $e)) }
                }
                foreach ($p in $personIds) {
                    $text = Get-CopilotPersonContext -Identity $p
                    if ($text) { $ctxItems += (New-CopilotContextMessage -Text (Truncate-Text $text $ctxMax) -Description ("person:" + $p)) }
                }
                if ($ctxItems.Count -gt 0) {
                    Add-CopilotAdditionalContext -Body $body -Items $ctxItems
                }
            }

            $stream = $parsed.Map.ContainsKey("stream")
            if ($stream) {
                if (-not (Invoke-CopilotChatStreamLive -ConversationId $id -Body $body)) {
                    Write-Warn "Falling back to non-streaming chat."
                    $resp = Invoke-GraphRequest -Method "POST" -Uri ("/copilot/conversations/" + $id + "/chat") -Body $body -Beta
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                }
            } else {
                $resp = Invoke-GraphRequest -Method "POST" -Uri ("/copilot/conversations/" + $id + "/chat") -Body $body -Beta
                if (-not $resp) { return }
                if ($parsed.Map.ContainsKey("text")) {
                    $outText = Extract-CopilotText $resp
                    if ($outText) { Write-Host $outText }
                } else {
                    $resp | ConvertTo-Json -Depth 8
                }
            }
        }
        default {
            Write-Warn "Usage: copilot chat create|send|ask|stream"
        }
    }
}


function Handle-CopilotSearchCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: copilot search --query <text> [--path <url>] [--pageSize <n>]"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    if ($action -eq "next") {
        $parsed = Parse-NamedArgs $rest
        $url = Get-ArgValue $parsed.Map "url"
        if (-not $url) {
            $url = $parsed.Positionals | Select-Object -First 1
        }
        if (-not $url) {
            Write-Warn "Usage: copilot search next --url <nextLink>"
            return
        }
        $resp = Invoke-GraphRequest -Method "GET" -Uri $url -Beta
        if ($resp) { $resp | ConvertTo-Json -Depth 8 }
        return
    }
    if ($action -in @("list","open","download")) {
        $parsed = Parse-NamedArgs $rest
        $hits = $global:CopilotLastSearchHits
        if (-not $hits) {
            Write-Warn "No cached search hits. Run: copilot search --query <text>"
            return
        }
        if ($action -eq "list") {
            Show-CopilotHits $hits
            return
        }
        $idx = Resolve-CopilotHitIndex $hits $parsed
        if ($idx -lt 0) {
            Write-Warn "Usage: copilot search open|download --index <n> [--out <file>]"
            return
        }
        $hit = $hits[$idx]
        if ($action -eq "open") {
            $info = Get-DriveItemInfoFromHit $hit
            if ($info -and $info.WebUrl) { Write-Host $info.WebUrl } else { Write-Warn "WebUrl not available." }
            return
        }
        if ($action -eq "download") {
            $out = Get-ArgValue $parsed.Map "out"
            Download-CopilotHit -Hit $hit -OutFile $out
            return
        }
    }

    $parsed = Parse-NamedArgs $Args
    $query = Get-ArgValue $parsed.Map "query"
    if (-not $query) {
        $query = $parsed.Positionals | Select-Object -First 1
    }
    $jsonRaw = Get-ArgValue $parsed.Map "json"
    $bodyFile = Get-ArgValue $parsed.Map "bodyFile"
    if ($jsonRaw -or $bodyFile) {
        $body = Read-JsonPayload $jsonRaw $bodyFile $null
    } else {
        if (-not $query) {
            Write-Warn "Usage: copilot search --query <text> [--path <url>] [--pageSize <n>]"
            return
        }
        $body = [ordered]@{ query = $query }
        $pageSize = Get-ArgValue $parsed.Map "pagesize"
        if ($pageSize) {
            try { $body.pageSize = [int]$pageSize } catch {}
        }
        $paths = Parse-CommaList (Get-ArgValue $parsed.Map "paths")
        $singlePath = Get-ArgValue $parsed.Map "path"
        if ($singlePath) { $paths += $singlePath }
        $meta = Parse-CommaList (Get-ArgValue $parsed.Map "metadata")

        $filter = Get-ArgValue $parsed.Map "filter"
        $dataSource = Get-ArgValue $parsed.Map "datasource"
        if (-not $dataSource) { $dataSource = "oneDrive" }

        if ($paths.Count -gt 0 -or $meta.Count -gt 0 -or $filter) {
            $ds = @{}
            if ($paths.Count -gt 0) {
                $clauses = @()
                foreach ($p in $paths) { $clauses += ("path:`"" + $p + "`"") }
                $ds.filterExpression = ($clauses -join " OR ")
            }
            if ($filter) {
                $ds.filterExpression = $filter
            }
            if ($meta.Count -gt 0) {
                $ds.resourceMetadataNames = $meta
            }
            $body.dataSources = @{}
            $body.dataSources[$dataSource] = $ds
        }
    }

    $resp = Invoke-GraphRequest -Method "POST" -Uri "/copilot/search" -Body $body -Beta
    if (-not $resp) { return }
    $hits = @()
    if ($resp.searchHits) { $hits = @($resp.searchHits) }
    $global:CopilotLastSearchHits = $hits
    if ($parsed.Map.ContainsKey("hits")) {
        if ($resp.searchHits) {
            $resp.searchHits | ConvertTo-Json -Depth 8
        } else {
            $resp | ConvertTo-Json -Depth 8
        }
        return
    }
    $resp | ConvertTo-Json -Depth 8
}


function Handle-CopilotRetrieveCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: copilot retrieve --query <text> --source sharePoint|oneDriveBusiness|externalItem [--max <n>]"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $Args[0].ToLowerInvariant()
    if ($action -in @("list","open","download")) {
        $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
        $parsed = Parse-NamedArgs $rest
        $hits = $global:CopilotLastRetrievalHits
        if (-not $hits) {
            Write-Warn "No cached retrieval hits. Run: copilot retrieve --query <text> ..."
            return
        }
        if ($action -eq "list") {
            Show-CopilotHits $hits
            return
        }
        $idx = Resolve-CopilotHitIndex $hits $parsed
        if ($idx -lt 0) {
            Write-Warn "Usage: copilot retrieve open|download --index <n> [--out <file>]"
            return
        }
        $hit = $hits[$idx]
        if ($action -eq "open") {
            $info = Get-DriveItemInfoFromHit $hit
            if ($info -and $info.WebUrl) { Write-Host $info.WebUrl } else { Write-Warn "WebUrl not available." }
            return
        }
        if ($action -eq "download") {
            $out = Get-ArgValue $parsed.Map "out"
            Download-CopilotHit -Hit $hit -OutFile $out
            return
        }
    }

    if ($action -in @("ask","chat")) {
        $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
        $parsed2 = Parse-NamedArgs $rest
        $prompt = Get-ArgValue $parsed2.Map "text"
        if (-not $prompt) { $prompt = Get-ArgValue $parsed2.Map "prompt" }
        if (-not $prompt) {
            Write-Warn "Usage: copilot retrieve ask --query <text> --source <source> --prompt <text> [--top <n>] [--stream]"
            return
        }
        $query = Get-ArgValue $parsed2.Map "query"
        if (-not $query) { $query = $parsed2.Positionals | Select-Object -First 1 }
        $source = Get-ArgValue $parsed2.Map "source"
        if (-not $query -or -not $source) {
            Write-Warn "Usage: copilot retrieve ask --query <text> --source <source> --prompt <text> [--top <n>] [--stream]"
            return
        }
        $top = Get-ArgValue $parsed2.Map "top"
        $stream = $parsed2.Map.ContainsKey("stream")
        $body = [ordered]@{
            queryString = $query
            dataSource  = $source
        }
        if ($top) { try { $body.maximumNumberOfResults = [int]$top } catch {} }
        $filter = Get-ArgValue $parsed2.Map "filter"
        if ($filter) { $body.filterExpression = $filter }
        $meta = Parse-CommaList (Get-ArgValue $parsed2.Map "metadata")
        if ($meta.Count -gt 0) { $body.resourceMetadata = $meta }
        if ($source.ToLowerInvariant() -eq "externalitem") {
            $conns = Parse-CommaList (Get-ArgValue $parsed2.Map "connections")
            if ($conns.Count -gt 0) {
                $body.dataSourceConfiguration = @{
                    externalItem = @{
                        connections = @($conns | ForEach-Object { @{ connectionId = $_ } })
                    }
                }
            }
        }
        $resp = Invoke-GraphRequestAuto -Method "POST" -Uri "/copilot/retrieval" -Body $body -Api $api -AllowFallback:$allowFallback
        if (-not $resp) { return }
        $hits = @()
        if ($resp.retrievalHits) { $hits = @($resp.retrievalHits) }
        $global:CopilotLastRetrievalHits = $hits
        $idxRaw = if ($top) { (1..([int]$top) -join ",") } else { "" }
        $chatArgs = @("chat","ask","--text",$prompt)
        if ($stream) { $chatArgs += "--stream" }
        if ($idxRaw) { $chatArgs += @("--useRetrieve",$idxRaw) }
        Handle-CopilotChatCommand $chatArgs
        return
    }

    $parsed = Parse-NamedArgs $Args
    $jsonRaw = Get-ArgValue $parsed.Map "json"
    $bodyFile = Get-ArgValue $parsed.Map "bodyFile"
    $useBeta = $parsed.Map.ContainsKey("beta")
    $useV1 = $parsed.Map.ContainsKey("v1")
    $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
    if ($parsed.Map.ContainsKey("auto")) { $allowFallback = $true }
    if ($useBeta -or $useV1) { $allowFallback = $false }
    $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }

    if ($jsonRaw -or $bodyFile) {
        $body = Read-JsonPayload $jsonRaw $bodyFile $null
    } else {
        $query = Get-ArgValue $parsed.Map "query"
        if (-not $query) {
            $query = $parsed.Positionals | Select-Object -First 1
        }
        $source = Get-ArgValue $parsed.Map "source"
        if (-not $query -or -not $source) {
            Write-Warn "Usage: copilot retrieve --query <text> --source sharePoint|oneDriveBusiness|externalItem [--max <n>]"
            return
        }
        $source = $source
        $body = [ordered]@{
            queryString = $query
            dataSource  = $source
        }
        $filter = Get-ArgValue $parsed.Map "filter"
        if ($filter) { $body.filterExpression = $filter }
        $max = Get-ArgValue $parsed.Map "max"
        if ($max) {
            try { $body.maximumNumberOfResults = [int]$max } catch {}
        }
        $meta = Parse-CommaList (Get-ArgValue $parsed.Map "metadata")
        if ($meta.Count -gt 0) {
            $body.resourceMetadata = $meta
        }
        if ($source.ToLowerInvariant() -eq "externalitem") {
            $conns = Parse-CommaList (Get-ArgValue $parsed.Map "connections")
            if ($conns.Count -gt 0) {
                $body.dataSourceConfiguration = @{
                    externalItem = @{
                        connections = @($conns | ForEach-Object { @{ connectionId = $_ } })
                    }
                }
            }
        }
    }

    $resp = Invoke-GraphRequestAuto -Method "POST" -Uri "/copilot/retrieval" -Body $body -Api $api -AllowFallback:$allowFallback
    if (-not $resp) { return }
    $hits = @()
    if ($resp.retrievalHits) { $hits = @($resp.retrievalHits) }
    $global:CopilotLastRetrievalHits = $hits
    if ($parsed.Map.ContainsKey("hits")) {
        if ($resp.retrievalHits) {
            $resp.retrievalHits | ConvertTo-Json -Depth 8
        } else {
            $resp | ConvertTo-Json -Depth 8
        }
        return
    }
    $resp | ConvertTo-Json -Depth 8
}


function Handle-CopilotCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: copilot chat|search|retrieve ..."
        return
    }
    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    switch ($sub) {
        "chat"     { Handle-CopilotChatCommand $rest }
        "search"   { Handle-CopilotSearchCommand $rest }
        "retrieve" { Handle-CopilotRetrieveCommand $rest }
        default    { Write-Warn "Usage: copilot chat|search|retrieve ..." }
    }
}
