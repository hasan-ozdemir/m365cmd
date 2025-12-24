# Handler: Powerpoint
# Purpose: Powerpoint command handlers.
function Handle-PowerPointCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: powerpoint list|search|get|download|upload|create|update|delete|convert|preview|share"
        return
    }
    if (-not (Require-GraphConnection)) { return }
    $action = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
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

