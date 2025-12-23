# Handler: Security
# Purpose: Security command handlers.
function Resolve-SecurityPath {
    param([string]$Type)
    $t = if ($Type) { $Type.ToLowerInvariant() } else { "alerts" }
    switch ($t) {
        "alerts" { return "/security/alerts" }
        "alerts_v2" { return "/security/alerts_v2" }
        "alertsv2" { return "/security/alerts_v2" }
        "incidents" { return "/security/incidents" }
        "securescores" { return "/security/secureScores" }
        "securescorecontrolprofiles" { return "/security/secureScoreControlProfiles" }
        "ti" { return "/security/tiIndicators" }
        "tiindicators" { return "/security/tiIndicators" }
        default { return "/security/alerts" }
    }
}

function Handle-SecurityCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: security list|get [--type alerts|alerts_v2|incidents|secureScores|secureScoreControlProfiles] OR security alert|incident|hunt|ti ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $type = Get-ArgValue $parsed.Map "type"
    $path = Resolve-SecurityPath $type
    $useBeta = $parsed.Map.ContainsKey("beta")
    $useV1 = $parsed.Map.ContainsKey("v1")
    $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
    if ($parsed.Map.ContainsKey("auto")) { $allowFallback = $true }
    if ($useBeta -or $useV1) { $allowFallback = $false }
    $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }

    switch ($sub) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @()
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($path + $qh.Query) -Headers $qh.Headers -Api $api -AllowFallback:$allowFallback
            if ($resp -and $resp.value) {
                $resp.value | ConvertTo-Json -Depth 10
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: security get <id> [--type alerts|alerts_v2|incidents|secureScores|secureScoreControlProfiles] [--beta|--auto]"
                return
            }
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($path + "/" + $id) -Api $api -AllowFallback:$allowFallback
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        default {
            Write-Warn "Usage: security list|get [--type alerts|alerts_v2|incidents|secureScores|secureScoreControlProfiles] OR security ti ..."
        }
    }
}
