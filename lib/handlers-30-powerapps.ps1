# Handler: Powerapps
# Purpose: Powerapps command handlers.
function Handle-PowerAppsCommand {
    param([string[]]$Args)
    Handle-PPCommand (@("app") + $Args)
}
