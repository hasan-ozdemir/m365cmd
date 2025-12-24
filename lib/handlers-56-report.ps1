# Handler: Report
# Purpose: Report command handlers.
function Handle-ReportCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: report list|run"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    $knownReports = @(
        "getOffice365ActiveUserDetail",
        "getOffice365ActiveUserCounts",
        "getOffice365ActiveUserUserDetail",
        "getOffice365GroupsActivityDetail",
        "getOffice365GroupsActivityCounts",
        "getOffice365GroupsActivityGroupDetail",
        "getOffice365GroupsActivityStorage",
        "getOffice365GroupsActivityFileCounts",
        "getMailboxUsageDetail",
        "getMailboxUsageMailboxCounts",
        "getMailboxUsageQuotaStatusMailboxCounts",
        "getMailboxUsageStorage",
        "getEmailActivityUserDetail",
        "getEmailActivityCounts",
        "getEmailActivityUserCounts",
        "getEmailAppUsageUserDetail",
        "getEmailAppUsageAppsUserCounts",
        "getEmailAppUsageUserCounts",
        "getMailboxUsageMailboxCounts",
        "getSharePointSiteUsageDetail",
        "getSharePointSiteUsageFileCounts",
        "getSharePointSiteUsageSiteCounts",
        "getSharePointSiteUsageStorage",
        "getSharePointActivityUserDetail",
        "getSharePointActivityFileCounts",
        "getSharePointActivityUserCounts",
        "getOneDriveUsageAccountDetail",
        "getOneDriveUsageAccountCounts",
        "getOneDriveUsageFileCounts",
        "getOneDriveUsageStorage",
        "getOneDriveActivityUserDetail",
        "getOneDriveActivityUserCounts",
        "getOneDriveActivityFileCounts",
        "getTeamsUserActivityUserDetail",
        "getTeamsUserActivityCounts",
        "getTeamsUserActivityUserCounts",
        "getTeamsDeviceUsageUserDetail",
        "getTeamsDeviceUsageDistributionUserCounts",
        "getTeamsDeviceUsageUserCounts",
        "getTeamsTeamActivityDetail",
        "getTeamsTeamActivityCounts",
        "getTeamsTeamActivityDistributionCounts",
        "getTeamsTeamActivityUserCounts",
        "getTeamsTeamActivityFileCounts",
        "getYammerActivityUserDetail",
        "getYammerActivityCounts",
        "getYammerActivityUserCounts",
        "getYammerDeviceUsageUserDetail",
        "getYammerDeviceUsageDistributionUserCounts",
        "getYammerDeviceUsageUserCounts",
        "getYammerGroupsActivityDetail",
        "getYammerGroupsActivityGroupDetail",
        "getYammerGroupsActivityCounts"
    )

    switch ($sub) {
        "list" {
            $knownReports | Sort-Object | Format-Wide -Column 3
        }
        "run" {
            $name = $parsed.Positionals | Select-Object -First 1
            if (-not $name) {
                Write-Warn "Usage: report run <name> [--period D7|D30|D90|D180] [--date YYYY-MM-DD] [--format csv|json] [--out <file>] [--beta|--auto]"
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

            $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
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
            Write-Warn "Usage: report list|run"
        }
    }
}

