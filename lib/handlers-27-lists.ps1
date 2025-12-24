# Handler: Lists
# Purpose: Lists command handlers.
function Handle-ListsCommand {
    param([string[]]$InputArgs)
    Handle-SPListCommand $InputArgs
}

