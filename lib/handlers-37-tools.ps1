# Handler: Tools
# Purpose: Utility command handlers.
function Handle-M365Command {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: m365 status|login|logout|request|search|version|docs"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    switch ($sub) {
        "status" {
            Show-Status
        }
        "login" {
            Invoke-Login ($rest | Select-Object -First 1)
        }
        "logout" {
            Invoke-Logout
        }
        "request" {
            if (-not $rest -or $rest.Count -lt 2) {
                Write-Warn "Usage: m365 request <get|post|patch|put|delete> <url|path> [--body <json>] [--bodyFile <path>] [--headers <json>] [--beta|--v1|--auto] [--out <file>]"
                return
            }
            Handle-GraphCommand (@("req") + $rest)
        }
        "search" {
            Handle-SearchCommand $rest
        }
        "version" {
            Write-Host "m365cmd (native CLI port)"
        }
        "docs" {
            Write-Host "Use /help or /help <topic> for command reference."
        }
        default {
            Write-Warn "Unknown m365 command. Use: m365 status|login|logout|request|search|version|docs"
        }
    }
}

