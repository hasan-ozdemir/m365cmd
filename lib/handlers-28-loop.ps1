# Handler: Loop
# Purpose: Loop command handlers.
function Handle-LoopCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: loop open|info|list|search|get|download|upload|create|update|delete|share"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }

    if ($sub -in @("open","info")) {
        switch ($sub) {
            "open" { Write-Host "https://loop.microsoft.com/" }
            "info" { Write-Warn "Loop components are stored as .loop files in OneDrive. Use loop file operations or file/stream commands." }
        }
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $fileOps = @("get","download","upload","create","update","delete","convert","preview","share","copy","move")
    if ($fileOps -contains $sub) {
        Handle-FileCommand (@($sub) + $rest)
        return
    }

    switch ($sub) {
        "list" { Invoke-FileTypeSearch -Types "loop,fluid" -Map (Parse-NamedArgs $rest).Map }
        "search" { Invoke-FileTypeSearch -Types "loop,fluid" -Map (Parse-NamedArgs $rest).Map }
        default { Write-Warn "Usage: loop open|info|list|search|get|download|upload|create|update|delete|share" }
    }
}

