# Handler: Powerautomate
# Purpose: Powerautomate command handlers.
function Handle-PowerAutomateCommand {
    param([string[]]$Args)
    Handle-PPCommand (@("flow") + $Args)
}
