# Handler: Lists
# Purpose: Lists command handlers.
function Handle-ListsCommand {
    param([string[]]$Args)
    Handle-SPListCommand $Args
}
