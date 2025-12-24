# Handler: Auditfeed
# Purpose: Auditfeed command handlers.
function Resolve-O365PublisherId {
    $pid = $global:Config.o365.publisherId
    if ($pid) { return $pid }
    return $null
}

function Handle-AuditFeedCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: auditfeed list|start|stop|content|notifications"
        return
    }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $tenantId = Resolve-TenantGuid
    if (-not $tenantId) {
        Write-Warn "Tenant ID not set. Use: /tenant set id <tenantGuid>"
        return
    }
    $baseUrl = $global:Config.o365.manageApiBase
    $scope = "https://manage.office.com/.default"
    $base = "/api/v1.0/" + $tenantId + "/activity/feed"
    $publisherId = Resolve-O365PublisherId

    switch ($sub) {
        "list" {
            $resp = Invoke-ExternalApiRequest -Method "GET" -Url ($base + "/subscriptions/list") -Scope $scope -BaseUrl $baseUrl
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "start" {
            $contentType = Get-ArgValue $parsed.Map "type"
            if (-not $contentType) {
                Write-Warn "Usage: auditfeed start --type <contentType> [--webhook <url>]"
                return
            }
            $qs = "contentType=" + (Encode-QueryValue $contentType)
            if ($publisherId) { $qs += "&PublisherIdentifier=" + (Encode-QueryValue $publisherId) }
            $body = $null
            $webhook = Get-ArgValue $parsed.Map "webhook"
            if ($webhook) {
                $authId = Get-ArgValue $parsed.Map "authId"
                $exp = Get-ArgValue $parsed.Map "expiration"
                $body = @{
                    webhook = @{
                        address = $webhook
                    }
                }
                if ($authId) { $body.webhook.authId = $authId }
                if ($exp) { $body.webhook.expiration = $exp }
            }
            $resp = Invoke-ExternalApiRequest -Method "POST" -Url ($base + "/subscriptions/start?" + $qs) -Body $body -Scope $scope -BaseUrl $baseUrl
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "stop" {
            $contentType = Get-ArgValue $parsed.Map "type"
            if (-not $contentType) {
                Write-Warn "Usage: auditfeed stop --type <contentType>"
                return
            }
            $qs = "contentType=" + (Encode-QueryValue $contentType)
            if ($publisherId) { $qs += "&PublisherIdentifier=" + (Encode-QueryValue $publisherId) }
            $resp = Invoke-ExternalApiRequest -Method "POST" -Url ($base + "/subscriptions/stop?" + $qs) -Scope $scope -BaseUrl $baseUrl
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "content" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: auditfeed content list|get ..."
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest2
            switch ($action) {
                "list" {
                    $contentType = Get-ArgValue $parsed2.Map "type"
                    $start = Get-ArgValue $parsed2.Map "start"
                    $end = Get-ArgValue $parsed2.Map "end"
                    if (-not $contentType -or -not $start -or -not $end) {
                        Write-Warn "Usage: auditfeed content list --type <contentType> --start <iso> --end <iso>"
                        return
                    }
                    $qs = "contentType=" + (Encode-QueryValue $contentType) + "&startTime=" + (Encode-QueryValue $start) + "&endTime=" + (Encode-QueryValue $end)
                    if ($publisherId) { $qs += "&PublisherIdentifier=" + (Encode-QueryValue $publisherId) }
                    $resp = Invoke-ExternalApiRequest -Method "GET" -Url ($base + "/subscriptions/content?" + $qs) -Scope $scope -BaseUrl $baseUrl
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "get" {
                    $uri = Get-ArgValue $parsed2.Map "uri"
                    if (-not $uri) {
                        Write-Warn "Usage: auditfeed content get --uri <contentUri> [--out <file>]"
                        return
                    }
                    $resp = Invoke-ExternalApiRequest -Method "GET" -Url $uri -Scope $scope -BaseUrl $baseUrl
                    if ($resp) {
                        $out = Get-ArgValue $parsed2.Map "out"
                        if ($out) {
                            $resp | ConvertTo-Json -Depth 10 | Set-Content -Path $out -Encoding ASCII
                            Write-Info ("Saved: " + $out)
                        } else {
                            $resp | ConvertTo-Json -Depth 10
                        }
                    }
                }
                default {
                    Write-Warn "Usage: auditfeed content list|get ..."
                }
            }
        }
        "notifications" {
            $action = if ($rest.Count -gt 0) { $rest[0].ToLowerInvariant() } else { "list" }
            $parsed2 = Parse-NamedArgs ($rest | Select-Object -Skip 1)
            if ($action -ne "list") {
                Write-Warn "Usage: auditfeed notifications list --type <contentType> [--start <iso>] [--end <iso>]"
                return
            }
            $contentType = Get-ArgValue $parsed2.Map "type"
            if (-not $contentType) {
                Write-Warn "Usage: auditfeed notifications list --type <contentType> [--start <iso>] [--end <iso>]"
                return
            }
            $qs = "contentType=" + (Encode-QueryValue $contentType)
            $start = Get-ArgValue $parsed2.Map "start"
            $end = Get-ArgValue $parsed2.Map "end"
            if ($start) { $qs += "&startTime=" + (Encode-QueryValue $start) }
            if ($end) { $qs += "&endTime=" + (Encode-QueryValue $end) }
            if ($publisherId) { $qs += "&PublisherIdentifier=" + (Encode-QueryValue $publisherId) }
            $resp = Invoke-ExternalApiRequest -Method "GET" -Url ($base + "/subscriptions/notifications?" + $qs) -Scope $scope -BaseUrl $baseUrl
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        default {
            Write-Warn "Usage: auditfeed list|start|stop|content|notifications"
        }
    }
}

