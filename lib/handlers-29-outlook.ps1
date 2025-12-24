# Handler: Outlook
# Purpose: Outlook command handlers.
function Handle-OutlookCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: outlook mail|calendar|contacts|people|todo|meeting"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    switch ($sub) {
        "mail" { Handle-MailCommand $rest }
        "calendar" { Handle-CalendarCommand $rest }
        "contacts" { Handle-ContactsCommand $rest }
        "people" { Handle-PeopleCommand $rest }
        "todo" { Handle-TodoCommand $rest }
        "meeting" { Handle-MeetingCommand $rest }
        default {
            Write-Warn "Usage: outlook mail|calendar|contacts|people|todo|meeting"
        }
    }
}

