# Handler: Powerapps
# Purpose: Powerapps command handlers.
function Handle-PowerAppsCommand {
    param([string[]]$InputArgs)
    Handle-PPCommand (@("app") + $InputArgs)
}

