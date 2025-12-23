# Handler: Word
# Purpose: Word command handlers.
function Handle-WordCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: word list|search|get|download|upload|create|update|delete|convert|preview|share"
        return
    }
    if (-not (Require-GraphConnection)) { return }
    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $fileOps = @("get","download","upload","create","update","delete","convert","preview","share","copy","move")
    if ($fileOps -contains $action) {
        Handle-FileCommand (@($action) + $rest)
        return
    }
    switch ($action) {
        "list" { Invoke-FileTypeSearch -Types "doc,docx,rtf,odt" -Map (Parse-NamedArgs $rest).Map }
        "search" { Invoke-FileTypeSearch -Types "doc,docx,rtf,odt" -Map (Parse-NamedArgs $rest).Map }
        default { Write-Warn "Usage: word list|search|get|download|upload|create|update|delete|convert|preview|share" }
    }
}
