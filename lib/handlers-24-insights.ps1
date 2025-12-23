# Handler: Insights
# Purpose: Insights command handlers.
function Handle-InsightsCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: insights list|get --type shared|trending|used [--user <upn|id>]"
        return
    }
    if (-not (Require-GraphConnection)) { return }
    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $type = (Get-ArgValue $parsed.Map "type")
    if (-not $type) { $type = "shared" }
    $user = Get-ArgValue $parsed.Map "user"
    $seg = if ($user) { "/users/" + $user } else { "/me" }
    $base = $seg + "/insights/" + $type

    switch ($action) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @("id","lastShared","lastUsed","resourceReference")
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
            if ($resp -and $resp.value) {
                $resp.value | Select-Object Id, LastShared, LastUsed, ResourceReference | Format-Table -AutoSize
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 8
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: insights get <id> --type shared|trending|used [--user <upn|id>]"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 8 }
        }
        default {
            Write-Warn "Usage: insights list|get --type shared|trending|used [--user <upn|id>]"
        }
    }
}
