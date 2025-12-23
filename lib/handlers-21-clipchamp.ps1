# Handler: Clipchamp
# Purpose: Clipchamp command handlers.
function Handle-ClipchampCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: clipchamp open|info|list|search|project|file"
        return
    }
    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }

    if ($sub -in @("open","info")) {
        switch ($sub) {
            "open" { Write-Host "https://app.clipchamp.com/" }
            "info" { Write-Warn "Clipchamp stores project assets in OneDrive/SharePoint. Use 'clipchamp project' or file/stream commands." }
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
        "list" {
            Invoke-FileTypeSearch -Types "clipchamp" -Map (Parse-NamedArgs $rest).Map
        }
        "search" {
            Invoke-FileTypeSearch -Types "clipchamp" -Map (Parse-NamedArgs $rest).Map
        }
        "project" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: clipchamp project list|get|open|assets|exports [--path <projectFolder>]"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed = Parse-NamedArgs $rest2
            switch ($action) {
                "list" { Invoke-FileTypeSearch -Types "clipchamp" -Map $parsed.Map }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: clipchamp project get <itemId>"
                        return
                    }
                    Handle-FileCommand @("get", $id) + (Build-FileArgsFromMap $parsed.Map)
                }
                "open" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $path = Get-ArgValue $parsed.Map "path"
                    if (-not $id -and -not $path) {
                        Write-Warn "Usage: clipchamp project open <itemId> OR --path <path>"
                        return
                    }
                    $base = Resolve-DriveBase $parsed.Map
                    if (-not $base) { $base = "/me/drive" }
                    $itemId = Resolve-DriveItemId -Base $base -ItemId $id -Path $path
                    if (-not $itemId) {
                        Write-Warn "File not found."
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/items/" + $itemId)
                    if ($resp -and $resp.webUrl) {
                        Write-Host $resp.webUrl
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 6
                    }
                }
                "assets" {
                    $path = Get-ArgValue $parsed.Map "path"
                    if (-not $path) { $path = $parsed.Positionals | Select-Object -First 1 }
                    if (-not $path) {
                        Write-Warn "Usage: clipchamp project assets --path <projectFolder> [--folder Assets]"
                        return
                    }
                    $folder = Get-ArgValue $parsed.Map "folder"
                    if (-not $folder) { $folder = "Assets" }
                    $full = (Normalize-DrivePath $path).TrimEnd("/") + "/" + $folder
                    $args2 = @("list","--path", $full) + (Build-FileArgsFromMap $parsed.Map)
                    Handle-FileCommand $args2
                }
                "exports" {
                    $path = Get-ArgValue $parsed.Map "path"
                    if (-not $path) { $path = $parsed.Positionals | Select-Object -First 1 }
                    if (-not $path) {
                        Write-Warn "Usage: clipchamp project exports --path <projectFolder> [--folder Export|Exports]"
                        return
                    }
                    $folder = Get-ArgValue $parsed.Map "folder"
                    if (-not $folder) { $folder = "Export" }
                    $full = (Normalize-DrivePath $path).TrimEnd("/") + "/" + $folder
                    $args2 = @("list","--path", $full) + (Build-FileArgsFromMap $parsed.Map)
                    Handle-FileCommand $args2
                }
                default {
                    Write-Warn "Usage: clipchamp project list|get|open|assets|exports [--path <projectFolder>]"
                }
            }
        }
        "file" {
            Handle-FileCommand $rest
        }
        default {
            Write-Warn "Usage: clipchamp open|info|list|search|project|file"
        }
    }
}
