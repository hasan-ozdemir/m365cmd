# Handler: Powerpoint
# Purpose: Powerpoint command handlers.
function Handle-PowerPointCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: powerpoint list|search|get|download|upload|create|update|delete|convert|preview|share"
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
        "list" { Invoke-FileTypeSearch -Types "ppt,pptx,pptm,pps,ppsx,pot,potx" -Map (Parse-NamedArgs $rest).Map }
        "search" { Invoke-FileTypeSearch -Types "ppt,pptx,pptm,pps,ppsx,pot,potx" -Map (Parse-NamedArgs $rest).Map }
        default { Write-Warn "Usage: powerpoint list|search|get|download|upload|create|update|delete|convert|preview|share" }
    }
}
