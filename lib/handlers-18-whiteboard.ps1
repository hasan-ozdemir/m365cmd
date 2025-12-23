# Handler: Whiteboard
# Purpose: Whiteboard command handlers.
function Handle-WhiteboardCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        $Args = @("list")
    }
    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $path = Get-ArgValue $parsed.Map "path"
    $item = $parsed.Positionals | Select-Object -First 1

    if (-not $path -and $action -eq "list") {
        $path = "Whiteboards"
    } elseif ($path -and -not $path.TrimStart("/", "\").ToLowerInvariant().StartsWith("whiteboards")) {
        $path = "Whiteboards/" + $path.TrimStart("/", "\")
    }
    if ($path) { $parsed.Map["path"] = $path }

    $newArgs = @($action)
    foreach ($k in $parsed.Map.Keys) {
        $v = $parsed.Map[$k]
        if ($v -is [bool] -and $v) {
            $newArgs += ("--" + $k)
        } elseif ($null -ne $v) {
            $newArgs += ("--" + $k)
            $newArgs += ($v.ToString())
        }
    }
    if ($item) {
        $newArgs += $item
    }
    $extraPos = @($parsed.Positionals | Select-Object -Skip 1)
    if ($extraPos.Count -gt 0) { $newArgs += $extraPos }
    Handle-FileCommand $newArgs
}
