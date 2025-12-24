# Handler: Powerautomate
# Purpose: Powerautomate command handlers.
function Handle-PowerAutomateCommand {
    param([string[]]$InputArgs)
    Handle-PPCommand (@("flow") + $InputArgs)
}

