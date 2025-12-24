# Handler: Learning
# Purpose: Learning command handlers.
function Handle-LearningCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: learning provider|content|activity ..."
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    switch ($sub) {
        "provider" { Handle-VivaCommand (@("provider") + $rest) }
        "content" { Handle-VivaCommand (@("content") + $rest) }
        "activity" { Handle-LearningActivityCommand $rest }
        default { Write-Warn "Usage: learning provider|content|activity ..." }
    }
}

function Handle-LearningActivityCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: learning activity list|get|create|delete [--user <upn|id>] [--beta|--auto]"
        return
    }
    if (-not (Require-GraphConnection)) { return }
    $action = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $user = Get-ArgValue $parsed.Map "user"
    $useBeta = $parsed.Map.ContainsKey("beta")
    $useV1 = $parsed.Map.ContainsKey("v1")
    $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
    if ($parsed.Map.ContainsKey("auto")) { $allowFallback = $true }
    if ($useBeta -or $useV1) { $allowFallback = $false }
    $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
    $seg = if ($user) { "/users/" + $user } else { "/me" }
    $base = $seg + "/employeeExperience/learningCourseActivities"

    switch ($action) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @()
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Api $api -AllowFallback:$allowFallback
            if ($resp) {
                if ($resp.value) { $resp.value | ConvertTo-Json -Depth 8 } else { $resp | ConvertTo-Json -Depth 8 }
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: learning activity get <id> [--user <upn|id>]"
                return
            }
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
            if ($resp) { $resp | ConvertTo-Json -Depth 8 }
        }
        "create" {
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $body) {
                Write-Warn "Usage: learning activity create --json <payload> [--user <upn|id>]"
                return
            }
            $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $api -AllowFallback:$allowFallback
            if ($resp) { $resp | ConvertTo-Json -Depth 8 }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: learning activity delete <id> [--force] [--user <upn|id>]"
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
            if ($resp -ne $null) { Write-Info "Learning activity deleted." }
        }
        default {
            Write-Warn "Usage: learning activity list|get|create|delete [--user <upn|id>] [--beta|--auto]"
        }
    }
}

