# Handler: Health
# Purpose: Health command handlers.
function Handle-HealthCommand {
    param([string[]]$Args)
    if (-not (Require-GraphConnection)) { return }
    $sub = if ($Args -and $Args.Count -gt 0) { $Args[0].ToLowerInvariant() } else { "list" }
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "list" { }
        "status" { }
        "issues" { }
        default { }
    }

    if ($sub -eq "issues") {
        $serviceId = $parsed.Positionals | Select-Object -First 1
        if (-not $serviceId) {
            Write-Warn "Usage: health issues <serviceId>"
            return
        }
        try {
            $issues = Get-MgServiceAnnouncementHealthOverviewIssue -ServiceHealthId $serviceId -All
            $issues | Select-Object Id, Title, Classification, Status, StartDateTime, LastModifiedDateTime | Format-Table -AutoSize
        } catch {
            Write-Err $_.Exception.Message
        }
        return
    }

    try {
        $health = Get-MgServiceAnnouncementHealthOverview -All
        $health | Select-Object Id, Service, Status | Format-Table -AutoSize
    } catch {
        Write-Err $_.Exception.Message
    }
}
