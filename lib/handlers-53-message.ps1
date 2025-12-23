# Handler: Message
# Purpose: Message command handlers.
function Handle-MessageCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: message list|get"
        return
    }
    if (-not (Require-GraphConnection)) { return }
    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "list" {
            $filter = Get-ArgValue $parsed.Map "filter"
            $top = Get-ArgValue $parsed.Map "top"
            try {
                if ($top) {
                    $msgs = if ($filter) { Get-MgServiceAnnouncementMessage -Filter $filter -Top ([int]$top) } else { Get-MgServiceAnnouncementMessage -Top ([int]$top) }
                } else {
                    $msgs = if ($filter) { Get-MgServiceAnnouncementMessage -Filter $filter -All } else { Get-MgServiceAnnouncementMessage -All }
                }
                $msgs | Select-Object Id, Title, Category, Severity, LastModifiedDateTime | Format-Table -AutoSize
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: message get <messageId>"
                return
            }
            try {
                Get-MgServiceAnnouncementMessage -ServiceUpdateMessageId $id | Format-List *
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        default {
            Write-Warn "Usage: message list|get"
        }
    }
}
