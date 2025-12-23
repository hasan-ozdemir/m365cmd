# Handler: Outlook
# Purpose: Outlook command handlers.
function Handle-OutlookCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: outlook mail|calendar|contacts|people|todo|meeting"
        return
    }
    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
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
