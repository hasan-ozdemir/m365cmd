# Handler: AdaptiveCard
# Purpose: Adaptive card webhook helpers.
function Expand-AdaptiveCardTemplate {
    param(
        [string]$CardJson,
        [object]$CardData
    )
    if (-not $CardJson -or -not $CardData) { return $CardJson }
    $map = @{}
    if ($CardData -is [hashtable]) {
        $map = $CardData
    } else {
        foreach ($p in $CardData.PSObject.Properties) {
            $map[$p.Name] = $p.Value
        }
    }
    $text = $CardJson
    foreach ($k in $map.Keys) {
        $v = $map[$k]
        if ($null -eq $v) { continue }
        $text = $text -replace ([regex]::Escape('${' + $k + '}')), [string]$v
    }
    return $text
}


function Handle-AdaptiveCardCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: adaptivecard send --url <webhookUrl> [--title <t>] [--description <d>] [--imageUrl <url>] [--actionUrl <url>] [--card <json>] [--cardData <json>] [--facts k=v,k=v]"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "send" {
            $url = Get-ArgValue $parsed.Map "url"
            if (-not $url) {
                Write-Warn "Usage: adaptivecard send --url <webhookUrl> ..."
                return
            }
            $title = Get-ArgValue $parsed.Map "title"
            if (-not $title) { $title = Get-ArgValue $parsed.Map "t" }
            $desc = Get-ArgValue $parsed.Map "description"
            if (-not $desc) { $desc = Get-ArgValue $parsed.Map "d" }
            $img = Get-ArgValue $parsed.Map "imageUrl"
            if (-not $img) { $img = Get-ArgValue $parsed.Map "i" }
            $actionUrl = Get-ArgValue $parsed.Map "actionUrl"
            if (-not $actionUrl) { $actionUrl = Get-ArgValue $parsed.Map "a" }
            $cardRaw = Get-ArgValue $parsed.Map "card"
            $cardDataRaw = Get-ArgValue $parsed.Map "cardData"
            $factsRaw = Get-ArgValue $parsed.Map "facts"
            $factRaw = Get-ArgValue $parsed.Map "fact"

            $card = $null
            if ($cardRaw) {
                $cardData = $null
                if ($cardDataRaw) { $cardData = Parse-Value $cardDataRaw }
                $expanded = if ($cardData) { Expand-AdaptiveCardTemplate $cardRaw $cardData } else { $cardRaw }
                $card = Parse-Value $expanded
                if ($card -is [string]) {
                    Write-Warn "Card JSON is invalid."
                    return
                }
            } else {
                $card = @{
                    type    = "AdaptiveCard"
                    '$schema' = "http://adaptivecards.io/schemas/adaptive-card.json"
                    version = "1.2"
                    body    = @()
                }
                if ($title) {
                    $card.body += @{
                        type   = "TextBlock"
                        size   = "Medium"
                        weight = "Bolder"
                        text   = $title
                    }
                }
                if ($img) {
                    $card.body += @{
                        type = "Image"
                        url  = $img
                        size = "Stretch"
                    }
                }
                if ($desc) {
                    $card.body += @{
                        type = "TextBlock"
                        text = $desc
                        wrap = $true
                    }
                }
                $facts = @{}
                if ($factsRaw) {
                    $facts = Parse-KvPairs $factsRaw
                } elseif ($factRaw) {
                    $facts = Parse-KvPairs $factRaw
                }
                if ($facts.Keys.Count -gt 0) {
                    $card.body += @{
                        type  = "FactSet"
                        facts = @($facts.Keys | ForEach-Object { @{ title = ($_ + ":"); value = [string]$facts[$_] } })
                    }
                }
                if ($actionUrl) {
                    $card.actions = @(
                        @{
                            type  = "Action.OpenUrl"
                            title = "View"
                            url   = $actionUrl
                        }
                    )
                }
            }

            $payload = @{
                type        = "message"
                attachments = @(
                    @{
                        contentType = "application/vnd.microsoft.card.adaptive"
                        content     = $card
                    }
                )
            }

            try {
                $resp = Invoke-RestMethod -Method Post -Uri $url -Headers @{ "content-type" = "application/json" } -Body ($payload | ConvertTo-Json -Depth 10)
                if ($resp -is [string] -and $resp.ToLowerInvariant().Contains("failed")) {
                    Write-Err $resp
                    return
                }
                if ($resp) { $resp | ConvertTo-Json -Depth 6 }
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        default {
            Write-Warn "Usage: adaptivecard send --url <webhookUrl> ..."
        }
    }
}
