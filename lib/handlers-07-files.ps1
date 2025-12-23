# Handler: Files
# Purpose: Files command handlers.
function Handle-DriveCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: drive list|get|delta"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "list" {
            $user = Get-ArgValue $parsed.Map "user"
            $site = Get-ArgValue $parsed.Map "site"
            $group = Get-ArgValue $parsed.Map "group"
            if ($site) {
                $uri = "/sites/" + $site + "/drives"
            } elseif ($group) {
                $uri = "/groups/" + $group + "/drives"
            } else {
                $seg = Resolve-UserSegment $user
                if (-not $seg) { return }
                $uri = $seg + "/drives"
            }
            $qh = Build-QueryAndHeaders $parsed.Map @("id", "name", "driveType", "webUrl")
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($uri + $qh.Query) -Headers $qh.Headers
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Id", "Name", "DriveType", "WebUrl")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 6
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: drive get <driveId>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ("/drives/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 6 }
        }
        "delta" {
            $token = Get-ArgValue $parsed.Map "token"
            $useBeta = $parsed.Map.ContainsKey("beta")
            if ($token) {
                $resp = Invoke-GraphRequest -Method "GET" -Uri $token -Beta:$useBeta
                if ($resp) { $resp | ConvertTo-Json -Depth 6 }
                return
            }
            $base = Resolve-DriveBase $parsed.Map
            if (-not $base) { return }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/root/delta") -Beta:$useBeta
            if ($resp) { $resp | ConvertTo-Json -Depth 6 }
        }
        default {
            Write-Warn "Usage: drive list|get|delta"
        }
    }
}



function Handle-FileCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: file list|get|create|update|delete|download|convert|preview|upload|copy|move|share"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $base = Resolve-DriveBase $parsed.Map
    if (-not $base) { return }
    $useBeta = $parsed.Map.ContainsKey("beta")

    switch ($sub) {
        "list" {
            $path = Get-ArgValue $parsed.Map "path"
            $item = Get-ArgValue $parsed.Map "item"
            if (-not $path -and -not $item) {
                $item = $parsed.Positionals | Select-Object -First 1
            }
            if ($path) {
                $p = Normalize-DrivePath $path
                $uri = $base + "/root:/" + $p + ":/children"
            } elseif ($item) {
                $uri = $base + "/items/" + $item + "/children"
            } else {
                $uri = $base + "/root/children"
            }
            $qh = Build-QueryAndHeaders $parsed.Map @("id", "name", "size", "lastModifiedDateTime", "webUrl")
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($uri + $qh.Query) -Headers $qh.Headers -Beta:$useBeta
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Name", "Id", "Size", "LastModifiedDateTime", "WebUrl")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 6
            }
        }
        "get" {
            $path = Get-ArgValue $parsed.Map "path"
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $path -and -not $id) {
                Write-Warn "Usage: file get <itemId> [--path <path>]"
                return
            }
            if ($path) {
                $p = Normalize-DrivePath $path
                $uri = $base + "/root:/" + $p
            } else {
                $uri = $base + "/items/" + $id
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri $uri -Beta:$useBeta
            if ($resp) { $resp | ConvertTo-Json -Depth 6 }
        }
        "create" {
            $name = Get-ArgValue $parsed.Map "name"
            $path = Get-ArgValue $parsed.Map "path"
            $local = Get-ArgValue $parsed.Map "local"
            $content = Get-ArgValue $parsed.Map "content"
            $isFolder = $parsed.Map.ContainsKey("folder")

            if (-not $name -and -not $local) {
                Write-Warn "Usage: file create --name <name> [--path <parentPath>] [--folder] [--content <text>|--local <file>]"
                return
            }
            $parent = Normalize-DrivePath $path
            if ($isFolder) {
                $folderUri = if ($parent) { $base + "/root:/" + $parent + ":/children" } else { $base + "/root/children" }
                $body = @{
                    name = $name
                    folder = @{}
                    "@microsoft.graph.conflictBehavior" = "rename"
                }
                $resp = Invoke-GraphRequest -Method "POST" -Uri $folderUri -Body $body -Beta:$useBeta
                if ($resp) { Write-Info "Folder created." }
                return
            }

            if ($local) {
                $dest = $null
                if ($name) {
                    $dest = if ($parent) { $parent.TrimEnd("/") + "/" + $name } else { $name }
                } elseif ($parent) {
                    $dest = $parent
                }
                Invoke-DriveUpload -Base $base -LocalPath $local -DestPath $dest -Beta:$useBeta
                return
            }

            if (-not $content) {
                Write-Warn "Usage: file create --name <name> [--path <parentPath>] --content <text>"
                return
            }
            $target = if ($parent) { $parent.TrimEnd("/") + "/" + $name } else { $name }
            $uri = $base + "/root:/" + $target + ":/content"
            $resp = Invoke-GraphRequest -Method "PUT" -Uri $uri -Body $content -ContentType "text/plain" -Beta:$useBeta
            if ($resp) { Write-Info "File created." }
        }
        "update" {
            $path = Get-ArgValue $parsed.Map "path"
            $id = $parsed.Positionals | Select-Object -First 1
            $setRaw = Get-ArgValue $parsed.Map "set"
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            $raw = if ($setRaw) { $setRaw } else { $jsonRaw }
            if (-not $raw -or (-not $path -and -not $id)) {
                Write-Warn "Usage: file update <itemId> --set key=value[,key=value] OR --set '{\"name\":\"New\"}'"
                return
            }
            $body = Parse-Value $raw
            if ($body -is [string]) { $body = Parse-KvPairs $raw }
            if ($body.Keys.Count -eq 0) {
                Write-Warn "No properties to update."
                return
            }
            $uri = if ($path) {
                $p = Normalize-DrivePath $path
                $base + "/root:/" + $p
            } else {
                $base + "/items/" + $id
            }
            $resp = Invoke-GraphRequest -Method "PATCH" -Uri $uri -Body $body -Beta:$useBeta
            if ($resp) { Write-Info "Item updated." }
        }
        "delete" {
            $path = Get-ArgValue $parsed.Map "path"
            $id = $parsed.Positionals | Select-Object -First 1
            $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
            if (-not $path -and -not $id) {
                Write-Warn "Usage: file delete <itemId> [--force] OR file delete --path <path>"
                return
            }
            if (-not $force) {
                $confirm = Read-Host "Type DELETE to confirm"
                if ($confirm -ne "DELETE") {
                    Write-Info "Canceled."
                    return
                }
            }
            $uri = if ($path) {
                $p = Normalize-DrivePath $path
                $base + "/root:/" + $p
            } else {
                $base + "/items/" + $id
            }
            $resp = Invoke-GraphRequest -Method "DELETE" -Uri $uri -Beta:$useBeta
            if ($resp -ne $null) { Write-Info "Item deleted." }
        }
        "download" {
            $path = Get-ArgValue $parsed.Map "path"
            $id = $parsed.Positionals | Select-Object -First 1
            $out = Get-ArgValue $parsed.Map "out"
            if (-not $out -or (-not $path -and -not $id)) {
                Write-Warn "Usage: file download <itemId> --out <file> OR file download --path <path> --out <file>"
                return
            }
            $uri = if ($path) {
                $p = Normalize-DrivePath $path
                $base + "/root:/" + $p + ":/content"
            } else {
                $base + "/items/" + $id + "/content"
            }
            Invoke-GraphDownload -Uri $uri -OutFile $out -Beta:$useBeta
        }
        "convert" {
            $path = Get-ArgValue $parsed.Map "path"
            $id = $parsed.Positionals | Select-Object -First 1
            $out = Get-ArgValue $parsed.Map "out"
            $format = Get-ArgValue $parsed.Map "format"
            if (-not $format) { $format = "pdf" }
            if (-not $out -or (-not $path -and -not $id)) {
                Write-Warn "Usage: file convert <itemId> --out <file> [--format pdf|html|txt] OR file convert --path <path> --out <file>"
                return
            }
            $itemId = Resolve-DriveItemId -Base $base -ItemId $id -Path $path -Beta:$useBeta
            if (-not $itemId) {
                Write-Warn "Usage: file convert <itemId> --out <file> [--format pdf|html|txt]"
                return
            }
            $uri = $base + "/items/" + $itemId + "/content?format=" + $format
            Invoke-GraphDownload -Uri $uri -OutFile $out -Beta:$useBeta
        }
        "preview" {
            $path = Get-ArgValue $parsed.Map "path"
            $id = $parsed.Positionals | Select-Object -First 1
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            $itemId = Resolve-DriveItemId -Base $base -ItemId $id -Path $path -Beta:$useBeta
            if (-not $itemId) {
                Write-Warn "Usage: file preview <itemId> [--path <path>] [--json <payload>]"
                return
            }
            if (-not $body) { $body = @{} }
            $resp = Invoke-GraphRequest -Method "POST" -Uri ($base + "/items/" + $itemId + "/preview") -Body $body -Beta:$useBeta
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "upload" {
            $local = Get-ArgValue $parsed.Map "local"
            $dest = Get-ArgValue $parsed.Map "dest"
            $path = Get-ArgValue $parsed.Map "path"
            $name = Get-ArgValue $parsed.Map "name"
            if (-not $local) {
                Write-Warn "Usage: file upload --local <file> [--dest <path>] [--path <folder>] [--name <name>]"
                return
            }
            if ($name) {
                $parent = Normalize-DrivePath $path
                $dest = if ($parent) { $parent.TrimEnd("/") + "/" + $name } else { $name }
            } elseif (-not $dest -and $path) {
                $dest = Normalize-DrivePath $path
            }
            Invoke-DriveUpload -Base $base -LocalPath $local -DestPath $dest -Beta:$useBeta
        }
        "copy" {
            $path = Get-ArgValue $parsed.Map "path"
            $id = $parsed.Positionals | Select-Object -First 1
            $dest = Get-ArgValue $parsed.Map "dest"
            $name = Get-ArgValue $parsed.Map "name"
            if (-not $dest) {
                Write-Warn "Usage: file copy <itemId> [--path <path>] --dest <folderPath> [--name <newName>]"
                return
            }
            $itemId = Resolve-DriveItemId -Base $base -ItemId $id -Path $path -Beta:$useBeta
            if (-not $itemId) {
                Write-Warn "Usage: file copy <itemId> [--path <path>] --dest <folderPath> [--name <newName>]"
                return
            }
            $destPath = Normalize-DrivePath $dest
            $parentRef = @{ path = "/drive/root:/" + $destPath }
            $body = @{ parentReference = $parentRef }
            if ($name) { $body.name = $name }
            $resp = Invoke-GraphRequest -Method "POST" -Uri ($base + "/items/" + $itemId + "/copy") -Body $body -Beta:$useBeta
            if ($resp -ne $null) { Write-Info "Copy requested." }
        }
        "move" {
            $path = Get-ArgValue $parsed.Map "path"
            $id = $parsed.Positionals | Select-Object -First 1
            $dest = Get-ArgValue $parsed.Map "dest"
            $name = Get-ArgValue $parsed.Map "name"
            if (-not $dest) {
                Write-Warn "Usage: file move <itemId> [--path <path>] --dest <folderPath> [--name <newName>]"
                return
            }
            $itemId = Resolve-DriveItemId -Base $base -ItemId $id -Path $path -Beta:$useBeta
            if (-not $itemId) {
                Write-Warn "Usage: file move <itemId> [--path <path>] --dest <folderPath> [--name <newName>]"
                return
            }
            $destPath = Normalize-DrivePath $dest
            $body = @{ parentReference = @{ path = "/drive/root:/" + $destPath } }
            if ($name) { $body.name = $name }
            $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/items/" + $itemId) -Body $body -Beta:$useBeta
            if ($resp -ne $null) { Write-Info "Move completed." }
        }
        "share" {
            $action = $parsed.Positionals | Select-Object -First 1
            $item = $parsed.Positionals | Select-Object -Skip 1 -First 1
            $path = Get-ArgValue $parsed.Map "path"
            if (-not $action) {
                Write-Warn "Usage: file share list|get|link|invite|update|delete <itemId> [--path <path>]"
                return
            }
            $itemId = Resolve-DriveItemId -Base $base -ItemId $item -Path $path -Beta:$useBeta
            if (-not $itemId) {
                Write-Warn "Usage: file share list|get|link|invite|update|delete <itemId> [--path <path>]"
                return
            }

            switch ($action) {
                "list" {
                    $qh = Build-QueryAndHeaders $parsed.Map @("id", "roles", "shareId", "link")
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/items/" + $itemId + "/permissions" + $qh.Query) -Headers $qh.Headers -Beta:$useBeta
                    if ($resp -and $resp.value) {
                        $resp.value | Select-Object Id, Roles, ShareId, Link | Format-Table -AutoSize
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 8
                    }
                }
                "get" {
                    $permId = Get-ArgValue $parsed.Map "perm"
                    if (-not $permId) { $permId = Get-ArgValue $parsed.Map "permission" }
                    if (-not $permId) {
                        Write-Warn "Usage: file share get <itemId> --perm <permissionId>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/items/" + $itemId + "/permissions/" + $permId) -Beta:$useBeta
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                }
                "link" {
                    $linkAction = $parsed.Positionals | Select-Object -Skip 2 -First 1
                    if (-not $linkAction) { $linkAction = "create" }
                    switch ($linkAction) {
                        "create" {
                            $type = Get-ArgValue $parsed.Map "type"
                            $scope = Get-ArgValue $parsed.Map "scope"
                            $password = Get-ArgValue $parsed.Map "password"
                            $expiration = Get-ArgValue $parsed.Map "expiration"
                            if (-not $type) { $type = "view" }
                            $t = $type.ToLowerInvariant()
                            switch ($t) {
                                "read" { $type = "view" }
                                "view" { $type = "view" }
                                "can-view" { $type = "view" }
                                "write" { $type = "edit" }
                                "edit" { $type = "edit" }
                                "can-edit" { $type = "edit" }
                                default { $type = $t }
                            }
                            if (-not $scope) { $scope = "anonymous" }
                            $body = @{
                                type  = $type
                                scope = $scope
                            }
                            if ($password) { $body.password = $password }
                            if ($expiration) { $body.expirationDateTime = $expiration }
                            $resp = Invoke-GraphRequest -Method "POST" -Uri ($base + "/items/" + $itemId + "/createLink") -Body $body -Beta:$useBeta
                            if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                        }
                        "update" {
                            $permId = Get-ArgValue $parsed.Map "perm"
                            $rolesRaw = Get-ArgValue $parsed.Map "roles"
                            if (-not $permId -or -not $rolesRaw) {
                                Write-Warn "Usage: file share link update <itemId> --perm <permissionId> --roles read|write"
                                return
                            }
                            $roles = Normalize-ShareRoles $rolesRaw
                            $body = @{ roles = $roles }
                            $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/items/" + $itemId + "/permissions/" + $permId) -Body $body -Beta:$useBeta
                            if ($resp) { Write-Info "Link permission updated." }
                        }
                        "delete" {
                            $permId = Get-ArgValue $parsed.Map "perm"
                            if (-not $permId) {
                                Write-Warn "Usage: file share link delete <itemId> --perm <permissionId> [--force]"
                                return
                            }
                            $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                            if (-not $force) {
                                $confirm = Read-Host "Type DELETE to confirm"
                                if ($confirm -ne "DELETE") {
                                    Write-Info "Canceled."
                                    return
                                }
                            }
                            $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($base + "/items/" + $itemId + "/permissions/" + $permId) -Beta:$useBeta
                            if ($resp -ne $null) { Write-Info "Link deleted." } else { Write-Info "Delete requested." }
                        }
                        default {
                            Write-Warn "Usage: file share link create|update|delete <itemId> ..."
                        }
                    }
                }
                "invite" {
                    $to = Get-ArgValue $parsed.Map "to"
                    $rolesRaw = Get-ArgValue $parsed.Map "roles"
                    if (-not $rolesRaw) { $rolesRaw = Get-ArgValue $parsed.Map "role" }
                    if (-not $to) {
                        Write-Warn "Usage: file share invite <itemId> --to a@b.com,b@b.com [--roles read|write] [--message text] [--requireSignIn true|false] [--send true|false]"
                        return
                    }
                    $roles = if ($rolesRaw) { Normalize-ShareRoles $rolesRaw } else { @("read") }
                    $message = Get-ArgValue $parsed.Map "message"
                    $requireSignIn = Parse-Bool (Get-ArgValue $parsed.Map "requireSignIn") $true
                    $send = Parse-Bool (Get-ArgValue $parsed.Map "send") $true
                    $body = @{
                        recipients = Build-RecipientList $to
                        roles      = $roles
                        message    = $message
                        requireSignIn = $requireSignIn
                        sendInvitation = $send
                    }
                    $resp = Invoke-GraphRequest -Method "POST" -Uri ($base + "/items/" + $itemId + "/invite") -Body $body -Beta:$useBeta
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                }
                "update" {
                    $permId = Get-ArgValue $parsed.Map "perm"
                    $rolesRaw = Get-ArgValue $parsed.Map "roles"
                    if (-not $permId -or -not $rolesRaw) {
                        Write-Warn "Usage: file share update <itemId> --perm <permissionId> --roles read|write"
                        return
                    }
                    $roles = Normalize-ShareRoles $rolesRaw
                    $body = @{ roles = $roles }
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/items/" + $itemId + "/permissions/" + $permId) -Body $body -Beta:$useBeta
                    if ($resp) { Write-Info "Permission updated." }
                }
                "delete" {
                    $permId = Get-ArgValue $parsed.Map "perm"
                    if (-not $permId) {
                        Write-Warn "Usage: file share delete <itemId> --perm <permissionId> [--force]"
                        return
                    }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host "Type DELETE to confirm"
                        if ($confirm -ne "DELETE") {
                            Write-Info "Canceled."
                            return
                        }
                    }
                    $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($base + "/items/" + $itemId + "/permissions/" + $permId) -Beta:$useBeta
                    if ($resp -ne $null) { Write-Info "Permission deleted." } else { Write-Info "Delete requested." }
                }
                default {
                    Write-Warn "Usage: file share list|get|link|invite|update|delete <itemId> [--path <path>]"
                }
            }
        }
        default {
            Write-Warn "Usage: file list|get|create|update|delete|download|upload|share"
        }
    }
}



function Handle-MailCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: mail folder|message ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }

    if ($sub -in @("folder", "folders")) {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: mail folder list|get|create|update|delete"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed = Parse-NamedArgs $rest2
        $seg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
        if (-not $seg) { return }

        switch ($action) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed.Map @("id", "displayName", "totalItemCount", "unreadItemCount")
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/mailFolders" + $qh.Query) -Headers $qh.Headers
                if ($resp -and $resp.value) {
                    Write-GraphTable $resp.value @("DisplayName", "Id", "TotalItemCount", "UnreadItemCount")
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 6
                }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: mail folder get <folderId>"
                    return
                }
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/mailFolders/" + $id)
                if ($resp) { $resp | ConvertTo-Json -Depth 6 }
            }
            "create" {
                $name = Get-ArgValue $parsed.Map "name"
                if (-not $name) { $name = Get-ArgValue $parsed.Map "displayName" }
                if (-not $name) {
                    Write-Warn "Usage: mail folder create --name <displayName>"
                    return
                }
                $body = @{ displayName = $name }
                $resp = Invoke-GraphRequest -Method "POST" -Uri ($seg + "/mailFolders") -Body $body
                if ($resp) { Write-Info "Mail folder created." }
            }
            "update" {
                $id = $parsed.Positionals | Select-Object -First 1
                $raw = Get-ArgValue $parsed.Map "set"
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                $name = Get-ArgValue $parsed.Map "name"
                if (-not $id) {
                    Write-Warn "Usage: mail folder update <folderId> --set key=value[,key=value]"
                    return
                }
                if (-not $raw -and -not $jsonRaw -and $name) {
                    $body = @{ displayName = $name }
                } else {
                    $rawBody = if ($raw) { $raw } else { $jsonRaw }
                    $body = Parse-Value $rawBody
                    if ($body -is [string]) { $body = Parse-KvPairs $rawBody }
                }
                if (-not $body -or $body.Keys.Count -eq 0) {
                    Write-Warn "No properties to update."
                    return
                }
                $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($seg + "/mailFolders/" + $id) -Body $body
                if ($resp) { Write-Info "Mail folder updated." }
            }
            "delete" {
                $id = $parsed.Positionals | Select-Object -First 1
                $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                if (-not $id) {
                    Write-Warn "Usage: mail folder delete <folderId> [--force]"
                    return
                }
                if (-not $force) {
                    $confirm = Read-Host "Type DELETE to confirm"
                    if ($confirm -ne "DELETE") {
                        Write-Info "Canceled."
                        return
                    }
                }
                $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($seg + "/mailFolders/" + $id)
                if ($resp -ne $null) { Write-Info "Mail folder deleted." }
            }
            default {
                Write-Warn "Usage: mail folder list|get|create|update|delete"
            }
        }
        return
    }

    if ($sub -in @("message", "messages", "msg")) {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: mail message list|get|create|update|delete|send"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed = Parse-NamedArgs $rest2
        $seg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
        if (-not $seg) { return }
        $folder = Get-ArgValue $parsed.Map "folder"

        switch ($action) {
            "list" {
                $base = if ($folder) { $seg + "/mailFolders/" + $folder + "/messages" } else { $seg + "/messages" }
                $qh = Build-QueryAndHeaders $parsed.Map @("id", "subject", "receivedDateTime")
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
                if ($resp -and $resp.value) {
                    Write-GraphTable $resp.value @("Subject", "Id", "ReceivedDateTime")
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 6
                }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: mail message get <messageId>"
                    return
                }
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/messages/" + $id)
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
            }
            "create" {
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                $body = if ($jsonRaw) { Parse-Value $jsonRaw } else { Build-MailMessage $parsed.Map }
                if (-not $body -or $body.Keys.Count -eq 0) {
                    Write-Warn "Usage: mail message create --subject <text> --to <a@b.com> [--body <text>]"
                    return
                }
                $base = if ($folder) { $seg + "/mailFolders/" + $folder + "/messages" } else { $seg + "/messages" }
                $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
                if ($resp) { Write-Info "Message created (draft)." }
            }
            "update" {
                $id = $parsed.Positionals | Select-Object -First 1
                $raw = Get-ArgValue $parsed.Map "set"
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                if (-not $id -or (-not $raw -and -not $jsonRaw)) {
                    Write-Warn "Usage: mail message update <messageId> --set key=value[,key=value]"
                    return
                }
                $rawBody = if ($raw) { $raw } else { $jsonRaw }
                $body = Parse-Value $rawBody
                if ($body -is [string]) { $body = Parse-KvPairs $rawBody }
                if ($body.Keys.Count -eq 0) {
                    Write-Warn "No properties to update."
                    return
                }
                $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($seg + "/messages/" + $id) -Body $body
                if ($resp) { Write-Info "Message updated." }
            }
            "delete" {
                $id = $parsed.Positionals | Select-Object -First 1
                $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                if (-not $id) {
                    Write-Warn "Usage: mail message delete <messageId> [--force]"
                    return
                }
                if (-not $force) {
                    $confirm = Read-Host "Type DELETE to confirm"
                    if ($confirm -ne "DELETE") {
                        Write-Info "Canceled."
                        return
                    }
                }
                $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($seg + "/messages/" + $id)
                if ($resp -ne $null) { Write-Info "Message deleted." }
            }
            "send" {
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                $message = if ($jsonRaw) { Parse-Value $jsonRaw } else { Build-MailMessage $parsed.Map }
                if (-not $message -or $message.Keys.Count -eq 0) {
                    Write-Warn "Usage: mail message send --subject <text> --to <a@b.com> [--body <text>] [--save true|false]"
                    return
                }
                $save = Parse-Bool (Get-ArgValue $parsed.Map "save") $true
                $body = @{
                    message = $message
                    saveToSentItems = $save
                }
                $resp = Invoke-GraphRequest -Method "POST" -Uri ($seg + "/sendMail") -Body $body
                if ($resp -ne $null) { Write-Info "Message sent." }
            }
            default {
                Write-Warn "Usage: mail message list|get|create|update|delete|send"
            }
        }
        return
    }

    Write-Warn "Usage: mail folder|message ..."
}



function Handle-CalendarCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: calendar list|get|create|update|delete|view|event"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }

    if ($sub -in @("event", "events")) {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: calendar event list|get|create|update|delete"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed = Parse-NamedArgs $rest2
        $seg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
        if (-not $seg) { return }
        $calendarId = Get-ArgValue $parsed.Map "calendar"
        $base = if ($calendarId) { $seg + "/calendars/" + $calendarId + "/events" } else { $seg + "/events" }

        switch ($action) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed.Map @("id", "subject", "start", "end")
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
                if ($resp -and $resp.value) {
                    Write-GraphTable $resp.value @("Subject", "Id", "Start", "End")
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 6
                }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: calendar event get <eventId>"
                    return
                }
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/events/" + $id)
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
            }
            "create" {
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                $body = $null
                if ($jsonRaw) {
                    $body = Parse-Value $jsonRaw
                } else {
                    $subject = Get-ArgValue $parsed.Map "subject"
                    $start = Get-ArgValue $parsed.Map "start"
                    $end = Get-ArgValue $parsed.Map "end"
                    $tz = Get-ArgValue $parsed.Map "tz"
                    if (-not $tz) { $tz = "UTC" }
                    if (-not $subject -or -not $start -or -not $end) {
                        Write-Warn "Usage: calendar event create --subject <text> --start <iso> --end <iso> [--tz <tz>]"
                        return
                    }
                    $bodyText = Get-ArgValue $parsed.Map "body"
                    $contentType = Get-ArgValue $parsed.Map "contentType"
                    $location = Get-ArgValue $parsed.Map "location"
                    $body = @{
                        subject = $subject
                        start = @{ dateTime = $start; timeZone = $tz }
                        end   = @{ dateTime = $end; timeZone = $tz }
                    }
                    if ($bodyText) { $body.body = New-ContentBody $bodyText $contentType }
                    if ($location) { $body.location = @{ displayName = $location } }
                }
                $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
                if ($resp) { Write-Info "Event created." }
            }
            "update" {
                $id = $parsed.Positionals | Select-Object -First 1
                $raw = Get-ArgValue $parsed.Map "set"
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                if (-not $id -or (-not $raw -and -not $jsonRaw)) {
                    Write-Warn "Usage: calendar event update <eventId> --set key=value[,key=value]"
                    return
                }
                $rawBody = if ($raw) { $raw } else { $jsonRaw }
                $body = Parse-Value $rawBody
                if ($body -is [string]) { $body = Parse-KvPairs $rawBody }
                if ($body.Keys.Count -eq 0) {
                    Write-Warn "No properties to update."
                    return
                }
                $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($seg + "/events/" + $id) -Body $body
                if ($resp) { Write-Info "Event updated." }
            }
            "delete" {
                $id = $parsed.Positionals | Select-Object -First 1
                $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                if (-not $id) {
                    Write-Warn "Usage: calendar event delete <eventId> [--force]"
                    return
                }
                if (-not $force) {
                    $confirm = Read-Host "Type DELETE to confirm"
                    if ($confirm -ne "DELETE") {
                        Write-Info "Canceled."
                        return
                    }
                }
                $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($seg + "/events/" + $id)
                if ($resp -ne $null) { Write-Info "Event deleted." }
            }
            default {
                Write-Warn "Usage: calendar event list|get|create|update|delete"
            }
        }
        return
    }

    $parsed = Parse-NamedArgs $rest
    $seg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
    if (-not $seg) { return }

    switch ($sub) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @("id", "name", "canEdit", "canShare")
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/calendars" + $qh.Query) -Headers $qh.Headers
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Name", "Id", "CanEdit", "CanShare")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 6
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: calendar get <calendarId>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/calendars/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 6 }
        }
        "create" {
            $name = Get-ArgValue $parsed.Map "name"
            if (-not $name) {
                Write-Warn "Usage: calendar create --name <displayName>"
                return
            }
            $body = @{ name = $name }
            $resp = Invoke-GraphRequest -Method "POST" -Uri ($seg + "/calendars") -Body $body
            if ($resp) { Write-Info "Calendar created." }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $raw = Get-ArgValue $parsed.Map "set"
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            if (-not $id -or (-not $raw -and -not $jsonRaw)) {
                Write-Warn "Usage: calendar update <calendarId> --set key=value[,key=value]"
                return
            }
            $rawBody = if ($raw) { $raw } else { $jsonRaw }
            $body = Parse-Value $rawBody
            if ($body -is [string]) { $body = Parse-KvPairs $rawBody }
            if ($body.Keys.Count -eq 0) {
                Write-Warn "No properties to update."
                return
            }
            $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($seg + "/calendars/" + $id) -Body $body
            if ($resp) { Write-Info "Calendar updated." }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
            if (-not $id) {
                Write-Warn "Usage: calendar delete <calendarId> [--force]"
                return
            }
            if (-not $force) {
                $confirm = Read-Host "Type DELETE to confirm"
                if ($confirm -ne "DELETE") {
                    Write-Info "Canceled."
                    return
                }
            }
            $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($seg + "/calendars/" + $id)
            if ($resp -ne $null) { Write-Info "Calendar deleted." }
        }
        "view" {
            $startRaw = Get-ArgValue $parsed.Map "start"
            $endRaw = Get-ArgValue $parsed.Map "end"
            if ($startRaw -and $endRaw) {
                $start = [datetime]::Parse($startRaw)
                $end = [datetime]::Parse($endRaw)
            } else {
                $range = Resolve-CalendarRange (Get-ArgValue $parsed.Map "date") (Get-ArgValue $parsed.Map "range")
                $start = $range.Start
                $end = $range.End
            }
            $qs = "?startDateTime=" + (Encode-QueryValue ($start.ToString("o"))) + "&endDateTime=" + (Encode-QueryValue ($end.ToString("o")))
            $headers = @{}
            $tz = Get-ArgValue $parsed.Map "tz"
            if ($tz) { $headers["Prefer"] = 'outlook.timezone="' + $tz + '"' }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/calendarView" + $qs) -Headers $headers
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Subject", "Id", "Start", "End")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 6
            }
        }
        default {
            Write-Warn "Usage: calendar list|get|create|update|delete|view|event"
        }
    }
}



function Handle-TodoCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: todo list|get|create|update|delete|task"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }

    if ($sub -in @("task", "tasks")) {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: todo task list|get|create|update|delete --list <listId>"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed = Parse-NamedArgs $rest2
        $seg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
        if (-not $seg) { return }
        $listId = Get-ArgValue $parsed.Map "list"
        if (-not $listId) {
            Write-Warn "Usage: todo task list|get|create|update|delete --list <listId>"
            return
        }
        $base = $seg + "/todo/lists/" + $listId + "/tasks"

        switch ($action) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed.Map @("id", "title", "status", "dueDateTime")
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
                if ($resp -and $resp.value) {
                    Write-GraphTable $resp.value @("Title", "Id", "Status", "DueDateTime")
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 6
                }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: todo task get <taskId> --list <listId>"
                    return
                }
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
                if ($resp) { $resp | ConvertTo-Json -Depth 6 }
            }
            "create" {
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                $body = $null
                if ($jsonRaw) {
                    $body = Parse-Value $jsonRaw
                } else {
                    $title = Get-ArgValue $parsed.Map "title"
                    if (-not $title) {
                        Write-Warn "Usage: todo task create --list <listId> --title <text> [--body <text>] [--status <status>] [--due <iso>] [--tz <tz>]"
                        return
                    }
                    $bodyText = Get-ArgValue $parsed.Map "body"
                    $status = Get-ArgValue $parsed.Map "status"
                    $due = Get-ArgValue $parsed.Map "due"
                    $tz = Get-ArgValue $parsed.Map "tz"
                    if (-not $tz) { $tz = "UTC" }
                    $body = @{ title = $title }
                    if ($bodyText) { $body.body = New-ContentBody $bodyText "text" }
                    if ($status) { $body.status = $status }
                    if ($due) { $body.dueDateTime = @{ dateTime = $due; timeZone = $tz } }
                }
                $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
                if ($resp) { Write-Info "Task created." }
            }
            "update" {
                $id = $parsed.Positionals | Select-Object -First 1
                $raw = Get-ArgValue $parsed.Map "set"
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                if (-not $id -or (-not $raw -and -not $jsonRaw)) {
                    Write-Warn "Usage: todo task update <taskId> --list <listId> --set key=value[,key=value]"
                    return
                }
                $rawBody = if ($raw) { $raw } else { $jsonRaw }
                $body = Parse-Value $rawBody
                if ($body -is [string]) { $body = Parse-KvPairs $rawBody }
                if ($body.Keys.Count -eq 0) {
                    Write-Warn "No properties to update."
                    return
                }
                $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
                if ($resp) { Write-Info "Task updated." }
            }
            "delete" {
                $id = $parsed.Positionals | Select-Object -First 1
                $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                if (-not $id) {
                    Write-Warn "Usage: todo task delete <taskId> --list <listId> [--force]"
                    return
                }
                if (-not $force) {
                    $confirm = Read-Host "Type DELETE to confirm"
                    if ($confirm -ne "DELETE") {
                        Write-Info "Canceled."
                        return
                    }
                }
                $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($base + "/" + $id)
                if ($resp -ne $null) { Write-Info "Task deleted." }
            }
            default {
                Write-Warn "Usage: todo task list|get|create|update|delete --list <listId>"
            }
        }
        return
    }

    $parsed = Parse-NamedArgs $rest
    $seg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
    if (-not $seg) { return }

    switch ($sub) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @("id", "displayName", "isOwner")
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/todo/lists" + $qh.Query) -Headers $qh.Headers
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("DisplayName", "Id", "IsOwner")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 6
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: todo get <listId>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/todo/lists/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 6 }
        }
        "create" {
            $name = Get-ArgValue $parsed.Map "name"
            if (-not $name) {
                Write-Warn "Usage: todo create --name <displayName>"
                return
            }
            $body = @{ displayName = $name }
            $resp = Invoke-GraphRequest -Method "POST" -Uri ($seg + "/todo/lists") -Body $body
            if ($resp) { Write-Info "List created." }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $raw = Get-ArgValue $parsed.Map "set"
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            if (-not $id -or (-not $raw -and -not $jsonRaw)) {
                Write-Warn "Usage: todo update <listId> --set key=value[,key=value]"
                return
            }
            $rawBody = if ($raw) { $raw } else { $jsonRaw }
            $body = Parse-Value $rawBody
            if ($body -is [string]) { $body = Parse-KvPairs $rawBody }
            if ($body.Keys.Count -eq 0) {
                Write-Warn "No properties to update."
                return
            }
            $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($seg + "/todo/lists/" + $id) -Body $body
            if ($resp) { Write-Info "List updated." }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
            if (-not $id) {
                Write-Warn "Usage: todo delete <listId> [--force]"
                return
            }
            if (-not $force) {
                $confirm = Read-Host "Type DELETE to confirm"
                if ($confirm -ne "DELETE") {
                    Write-Info "Canceled."
                    return
                }
            }
            $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($seg + "/todo/lists/" + $id)
            if ($resp -ne $null) { Write-Info "List deleted." }
        }
        default {
            Write-Warn "Usage: todo list|get|create|update|delete|task"
        }
    }
}



function Handle-OneNoteCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: onenote notebook|section|page ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }

    if ($sub -in @("notebook", "notebooks")) {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: onenote notebook list|get|create|update|delete"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed = Parse-NamedArgs $rest2
        $seg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
        if (-not $seg) { return }

        switch ($action) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed.Map @("id", "displayName")
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/onenote/notebooks" + $qh.Query) -Headers $qh.Headers
                if ($resp -and $resp.value) {
                    Write-GraphTable $resp.value @("DisplayName", "Id")
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 6
                }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: onenote notebook get <notebookId>"
                    return
                }
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/onenote/notebooks/" + $id)
                if ($resp) { $resp | ConvertTo-Json -Depth 6 }
            }
            "create" {
                $name = Get-ArgValue $parsed.Map "name"
                if (-not $name) {
                    Write-Warn "Usage: onenote notebook create --name <displayName>"
                    return
                }
                $body = @{ displayName = $name }
                $resp = Invoke-GraphRequest -Method "POST" -Uri ($seg + "/onenote/notebooks") -Body $body
                if ($resp) { Write-Info "Notebook created." }
            }
            "update" {
                $id = $parsed.Positionals | Select-Object -First 1
                $raw = Get-ArgValue $parsed.Map "set"
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                if (-not $id -or (-not $raw -and -not $jsonRaw)) {
                    Write-Warn "Usage: onenote notebook update <notebookId> --set key=value[,key=value]"
                    return
                }
                $rawBody = if ($raw) { $raw } else { $jsonRaw }
                $body = Parse-Value $rawBody
                if ($body -is [string]) { $body = Parse-KvPairs $rawBody }
                if ($body.Keys.Count -eq 0) {
                    Write-Warn "No properties to update."
                    return
                }
                $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($seg + "/onenote/notebooks/" + $id) -Body $body
                if ($resp) { Write-Info "Notebook updated." }
            }
            "delete" {
                $id = $parsed.Positionals | Select-Object -First 1
                $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                if (-not $id) {
                    Write-Warn "Usage: onenote notebook delete <notebookId> [--force]"
                    return
                }
                if (-not $force) {
                    $confirm = Read-Host "Type DELETE to confirm"
                    if ($confirm -ne "DELETE") {
                        Write-Info "Canceled."
                        return
                    }
                }
                $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($seg + "/onenote/notebooks/" + $id)
                if ($resp -ne $null) { Write-Info "Notebook deleted." }
            }
            default {
                Write-Warn "Usage: onenote notebook list|get|create|update|delete"
            }
        }
        return
    }

    if ($sub -in @("section", "sections")) {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: onenote section list|get|create|update|delete"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed = Parse-NamedArgs $rest2
        $seg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
        if (-not $seg) { return }
        $notebookId = Get-ArgValue $parsed.Map "notebook"

        switch ($action) {
            "list" {
                $uri = if ($notebookId) { $seg + "/onenote/notebooks/" + $notebookId + "/sections" } else { $seg + "/onenote/sections" }
                $qh = Build-QueryAndHeaders $parsed.Map @("id", "displayName")
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($uri + $qh.Query) -Headers $qh.Headers
                if ($resp -and $resp.value) {
                    Write-GraphTable $resp.value @("DisplayName", "Id")
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 6
                }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: onenote section get <sectionId>"
                    return
                }
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/onenote/sections/" + $id)
                if ($resp) { $resp | ConvertTo-Json -Depth 6 }
            }
            "create" {
                $name = Get-ArgValue $parsed.Map "name"
                if (-not $name -or -not $notebookId) {
                    Write-Warn "Usage: onenote section create --notebook <id> --name <displayName>"
                    return
                }
                $body = @{ displayName = $name }
                $resp = Invoke-GraphRequest -Method "POST" -Uri ($seg + "/onenote/notebooks/" + $notebookId + "/sections") -Body $body
                if ($resp) { Write-Info "Section created." }
            }
            "update" {
                $id = $parsed.Positionals | Select-Object -First 1
                $raw = Get-ArgValue $parsed.Map "set"
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                if (-not $id -or (-not $raw -and -not $jsonRaw)) {
                    Write-Warn "Usage: onenote section update <sectionId> --set key=value[,key=value]"
                    return
                }
                $rawBody = if ($raw) { $raw } else { $jsonRaw }
                $body = Parse-Value $rawBody
                if ($body -is [string]) { $body = Parse-KvPairs $rawBody }
                if ($body.Keys.Count -eq 0) {
                    Write-Warn "No properties to update."
                    return
                }
                $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($seg + "/onenote/sections/" + $id) -Body $body
                if ($resp) { Write-Info "Section updated." }
            }
            "delete" {
                $id = $parsed.Positionals | Select-Object -First 1
                $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                if (-not $id) {
                    Write-Warn "Usage: onenote section delete <sectionId> [--force]"
                    return
                }
                if (-not $force) {
                    $confirm = Read-Host "Type DELETE to confirm"
                    if ($confirm -ne "DELETE") {
                        Write-Info "Canceled."
                        return
                    }
                }
                $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($seg + "/onenote/sections/" + $id)
                if ($resp -ne $null) { Write-Info "Section deleted." }
            }
            default {
                Write-Warn "Usage: onenote section list|get|create|update|delete"
            }
        }
        return
    }

    if ($sub -in @("page", "pages")) {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: onenote page list|get|create|update|delete|content"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed = Parse-NamedArgs $rest2
        $seg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
        if (-not $seg) { return }
        $sectionId = Get-ArgValue $parsed.Map "section"

        switch ($action) {
            "list" {
                $uri = if ($sectionId) { $seg + "/onenote/sections/" + $sectionId + "/pages" } else { $seg + "/onenote/pages" }
                $qh = Build-QueryAndHeaders $parsed.Map @("id", "title", "createdDateTime", "lastModifiedDateTime")
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($uri + $qh.Query) -Headers $qh.Headers
                if ($resp -and $resp.value) {
                    Write-GraphTable $resp.value @("Title", "Id", "CreatedDateTime", "LastModifiedDateTime")
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 6
                }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: onenote page get <pageId>"
                    return
                }
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/onenote/pages/" + $id)
                if ($resp) { $resp | ConvertTo-Json -Depth 6 }
            }
            "content" {
                $id = $parsed.Positionals | Select-Object -First 1
                $out = Get-ArgValue $parsed.Map "out"
                if (-not $id) {
                    Write-Warn "Usage: onenote page content <pageId> [--out <file>]"
                    return
                }
                if ($out) {
                    Invoke-GraphDownload -Uri ($seg + "/onenote/pages/" + $id + "/content") -OutFile $out
                } else {
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/onenote/pages/" + $id + "/content")
                    if ($resp) { Write-Host $resp }
                }
            }
            "create" {
                $title = Get-ArgValue $parsed.Map "title"
                $content = Get-ArgValue $parsed.Map "content"
                $contentFile = Get-ArgValue $parsed.Map "file"
                if (-not $sectionId -or -not $title) {
                    Write-Warn "Usage: onenote page create --section <id> --title <text> [--content <html>] [--file <htmlFile>]"
                    return
                }
                if ($contentFile) {
                    if (-not (Test-Path $contentFile)) {
                        Write-Warn "Content file not found."
                        return
                    }
                    $html = Get-Content -Raw -Path $contentFile
                } else {
                    if (-not $content) { $content = "" }
                    $html = "<html><head><title>" + $title + "</title></head><body>" + $content + "</body></html>"
                }
                $resp = Invoke-GraphRequest -Method "POST" -Uri ($seg + "/onenote/sections/" + $sectionId + "/pages") -Body $html -ContentType "text/html"
                if ($resp) { Write-Info "Page created." }
            }
            "update" {
                $id = $parsed.Positionals | Select-Object -First 1
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                if (-not $id -or -not $jsonRaw) {
                    Write-Warn "Usage: onenote page update <pageId> --json <patchArray>"
                    return
                }
                $body = Parse-Value $jsonRaw
                $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($seg + "/onenote/pages/" + $id + "/content") -Body $body -ContentType "application/json"
                if ($resp) { Write-Info "Page updated." }
            }
            "delete" {
                $id = $parsed.Positionals | Select-Object -First 1
                $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                if (-not $id) {
                    Write-Warn "Usage: onenote page delete <pageId> [--force]"
                    return
                }
                if (-not $force) {
                    $confirm = Read-Host "Type DELETE to confirm"
                    if ($confirm -ne "DELETE") {
                        Write-Info "Canceled."
                        return
                    }
                }
                $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($seg + "/onenote/pages/" + $id)
                if ($resp -ne $null) { Write-Info "Page deleted." }
            }
            default {
                Write-Warn "Usage: onenote page list|get|create|update|delete|content"
            }
        }
        return
    }

    Write-Warn "Usage: onenote notebook|section|page ..."
}



function Handle-ChatCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: chat list|get|create|message"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }

    if ($sub -in @("message", "messages", "msg")) {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: chat message list|get|create|update|delete --chat <chatId>"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed = Parse-NamedArgs $rest2
        $chatId = Get-ArgValue $parsed.Map "chat"
        if (-not $chatId) {
            Write-Warn "Usage: chat message list|get|create|update|delete --chat <chatId>"
            return
        }
        $base = "/chats/" + $chatId + "/messages"

        switch ($action) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed.Map @("id", "createdDateTime")
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
                if ($resp -and $resp.value) {
                    Write-GraphTable $resp.value @("Id", "CreatedDateTime")
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 6
                }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: chat message get <messageId> --chat <chatId>"
                    return
                }
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
                if ($resp) { $resp | ConvertTo-Json -Depth 6 }
            }
            "create" {
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                $content = Get-ArgValue $parsed.Map "body"
                $contentType = Get-ArgValue $parsed.Map "contentType"
                if (-not $jsonRaw -and -not $content) {
                    Write-Warn "Usage: chat message create --chat <chatId> --body <text> [--contentType text|html]"
                    return
                }
                $body = if ($jsonRaw) { Parse-Value $jsonRaw } else { @{ body = (New-ContentBody $content $contentType) } }
                $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
                if ($resp) { Write-Info "Chat message created." }
            }
            "update" {
                $id = $parsed.Positionals | Select-Object -First 1
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                $content = Get-ArgValue $parsed.Map "body"
                $contentType = Get-ArgValue $parsed.Map "contentType"
                if (-not $id) {
                    Write-Warn "Usage: chat message update <messageId> --chat <chatId> [--body <text>|--json <json>]"
                    return
                }
                if (-not $jsonRaw -and -not $content) {
                    Write-Warn "Usage: chat message update <messageId> --chat <chatId> --body <text> OR --json <json>"
                    return
                }
                $body = if ($jsonRaw) { Parse-Value $jsonRaw } else { @{ body = (New-ContentBody $content $contentType) } }
                $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
                if ($resp) { Write-Info "Chat message updated." }
            }
            "delete" {
                $id = $parsed.Positionals | Select-Object -First 1
                $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                if (-not $id) {
                    Write-Warn "Usage: chat message delete <messageId> --chat <chatId> [--force]"
                    return
                }
                if (-not $force) {
                    $confirm = Read-Host "Type DELETE to confirm"
                    if ($confirm -ne "DELETE") {
                        Write-Info "Canceled."
                        return
                    }
                }
                $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($base + "/" + $id)
                if ($resp -ne $null) { Write-Info "Chat message deleted." }
            }
            default {
                Write-Warn "Usage: chat message list|get|create|update|delete --chat <chatId>"
            }
        }
        return
    }

    $parsed = Parse-NamedArgs $rest
    $seg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
    if (-not $seg) { return }

    switch ($sub) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @("id", "topic", "chatType")
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/chats" + $qh.Query) -Headers $qh.Headers
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Id", "ChatType", "Topic")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 6
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: chat get <chatId>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ("/chats/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 6 }
        }
        "create" {
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            if (-not $jsonRaw) {
                Write-Warn "Usage: chat create --json <chatPayload>"
                return
            }
            $body = Parse-Value $jsonRaw
            $resp = Invoke-GraphRequest -Method "POST" -Uri "/chats" -Body $body
            if ($resp) { Write-Info "Chat created." }
        }
        default {
            Write-Warn "Usage: chat list|get|create|message"
        }
    }
}



function Handle-ChannelMessageCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: channelmsg list|get|create|update|delete --team <teamId> --channel <channelId>"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $team = Get-ArgValue $parsed.Map "team"
    $channel = Get-ArgValue $parsed.Map "channel"
    if (-not $team -or -not $channel) {
        Write-Warn "Usage: channelmsg list|get|create|update|delete --team <teamId> --channel <channelId>"
        return
    }
    $base = "/teams/" + $team + "/channels/" + $channel + "/messages"

    switch ($sub) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @("id", "createdDateTime")
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Id", "CreatedDateTime")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 6
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: channelmsg get <messageId> --team <teamId> --channel <channelId>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 6 }
        }
        "create" {
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            $content = Get-ArgValue $parsed.Map "body"
            $contentType = Get-ArgValue $parsed.Map "contentType"
            if (-not $jsonRaw -and -not $content) {
                Write-Warn "Usage: channelmsg create --team <teamId> --channel <channelId> --body <text> [--contentType text|html]"
                return
            }
            $body = if ($jsonRaw) { Parse-Value $jsonRaw } else { @{ body = (New-ContentBody $content $contentType) } }
            $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
            if ($resp) { Write-Info "Channel message created." }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            $content = Get-ArgValue $parsed.Map "body"
            $contentType = Get-ArgValue $parsed.Map "contentType"
            if (-not $id) {
                Write-Warn "Usage: channelmsg update <messageId> --team <teamId> --channel <channelId> [--body <text>|--json <json>]"
                return
            }
            if (-not $jsonRaw -and -not $content) {
                Write-Warn "Usage: channelmsg update <messageId> --team <teamId> --channel <channelId> --body <text> OR --json <json>"
                return
            }
            $body = if ($jsonRaw) { Parse-Value $jsonRaw } else { @{ body = (New-ContentBody $content $contentType) } }
            $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
            if ($resp) { Write-Info "Channel message updated." }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
            if (-not $id) {
                Write-Warn "Usage: channelmsg delete <messageId> --team <teamId> --channel <channelId> [--force]"
                return
            }
            if (-not $force) {
                $confirm = Read-Host "Type DELETE to confirm"
                if ($confirm -ne "DELETE") {
                    Write-Info "Canceled."
                    return
                }
            }
            $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($base + "/" + $id)
            if ($resp -ne $null) { Write-Info "Channel message deleted." }
        }
        default {
            Write-Warn "Usage: channelmsg list|get|create|update|delete --team <teamId> --channel <channelId>"
        }
    }
}


function Handle-ContactsCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: contacts folder|item ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    if (-not $rest -or $rest.Count -eq 0) {
        Write-Warn "Usage: contacts folder|item ..."
        return
    }
    $action = $rest[0].ToLowerInvariant()
    $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest2
    $seg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
    if (-not $seg) { return }

    switch ($sub) {
        "folder" {
            $base = $seg + "/contactFolders"
            switch ($action) {
                "list" {
                    $qh = Build-QueryAndHeaders $parsed.Map @("id", "displayName")
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
                    if ($resp -and $resp.value) {
                        Write-GraphTable $resp.value @("Id", "DisplayName")
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 10
                    }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: contacts folder get <id> [--user <upn|id>]"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "create" {
                    $name = Get-ArgValue $parsed.Map "name"
                    $jsonRaw = Get-ArgValue $parsed.Map "json"
                    $body = if ($jsonRaw) { Parse-Value $jsonRaw } else { @{ displayName = $name } }
                    if (-not $body -or (-not $body.displayName -and -not $body.DisplayName)) {
                        Write-Warn "Usage: contacts folder create --name <text> OR --json <payload>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: contacts folder update <id> --json <payload> OR --set key=value"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
                    if ($resp -ne $null) { Write-Info "Contact folder updated." }
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: contacts folder delete <id> [--force]"
                        return
                    }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host "Type DELETE to confirm"
                        if ($confirm -ne "DELETE") {
                            Write-Info "Canceled."
                            return
                        }
                    }
                    $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($base + "/" + $id)
                    if ($resp -ne $null) { Write-Info "Contact folder deleted." }
                }
                default {
                    Write-Warn "Usage: contacts folder list|get|create|update|delete"
                }
            }
        }
        "item" {
            $folder = Get-ArgValue $parsed.Map "folder"
            $base = if ($folder) { $seg + "/contactFolders/" + $folder + "/contacts" } else { $seg + "/contacts" }
            switch ($action) {
                "list" {
                    $qh = Build-QueryAndHeaders $parsed.Map @("id", "displayName", "emailAddresses")
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
                    if ($resp -and $resp.value) {
                        $resp.value | Select-Object Id, DisplayName, EmailAddresses | Format-Table -AutoSize
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 10
                    }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: contacts item get <id> [--folder <id>]"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "create" {
                    $jsonRaw = Get-ArgValue $parsed.Map "json"
                    if (-not $jsonRaw) {
                        Write-Warn "Usage: contacts item create --json <payload>"
                        return
                    }
                    $body = Parse-Value $jsonRaw
                    $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: contacts item update <id> --json <payload> OR --set key=value"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
                    if ($resp -ne $null) { Write-Info "Contact updated." }
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: contacts item delete <id> [--force]"
                        return
                    }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host "Type DELETE to confirm"
                        if ($confirm -ne "DELETE") {
                            Write-Info "Canceled."
                            return
                        }
                    }
                    $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($base + "/" + $id)
                    if ($resp -ne $null) { Write-Info "Contact deleted." }
                }
                default {
                    Write-Warn "Usage: contacts item list|get|create|update|delete"
                }
            }
        }
        default {
            Write-Warn "Usage: contacts folder|item ..."
        }
    }
}

function Handle-PeopleCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: people list|get|create|update|delete [--user <upn|id>]"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }

    if ($sub -in @("create", "update", "delete")) {
        Write-Warn "People API is read-only. Redirecting to contacts item CRUD."
        Handle-ContactsCommand (@("item", $sub) + $rest)
        return
    }

    $parsed = Parse-NamedArgs $rest
    $seg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
    if (-not $seg) { return }
    $base = $seg + "/people"

    switch ($sub) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @("id", "displayName")
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Id", "DisplayName")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 8
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: people get <id> [--user <upn|id>]"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 8 }
        }
        default {
            Write-Warn "Usage: people list|get|create|update|delete [--user <upn|id>]"
        }
    }
}



function Handle-PlannerCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: planner plan|bucket|task ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    if (-not $rest -or $rest.Count -eq 0) {
        Write-Warn "Usage: planner plan|bucket|task ..."
        return
    }
    $action = $rest[0].ToLowerInvariant()
    $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest2

    function Get-PlannerEtagLocal {
        param([string]$Path)
        return (Get-GraphEtag $Path)
    }

    switch ($sub) {
        "plan" {
            switch ($action) {
                "list" {
                    $group = Get-ArgValue $parsed.Map "group"
                    if (-not $group) {
                        Write-Warn "Usage: planner plan list --group <groupId>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/groups/" + $group + "/planner/plans")
                    if ($resp -and $resp.value) {
                        Write-GraphTable $resp.value @("Id", "Title", "Owner")
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 10
                    }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: planner plan get <id>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/planner/plans/" + $id)
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "create" {
                    $title = Get-ArgValue $parsed.Map "title"
                    $group = Get-ArgValue $parsed.Map "group"
                    $jsonRaw = Get-ArgValue $parsed.Map "json"
                    $body = if ($jsonRaw) { Parse-Value $jsonRaw } else { @{ title = $title; owner = $group } }
                    if (-not $body -or -not $body.title -or -not $body.owner) {
                        Write-Warn "Usage: planner plan create --group <groupId> --title <text> OR --json <payload>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "POST" -Uri "/planner/plans" -Body $body
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: planner plan update <id> --json <payload> OR --set key=value"
                        return
                    }
                    $etag = Get-ArgValue $parsed.Map "etag"
                    if (-not $etag) { $etag = Get-PlannerEtagLocal ("/planner/plans/" + $id) }
                    $headers = @{}
                    if ($etag) { $headers["If-Match"] = $etag }
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri ("/planner/plans/" + $id) -Body $body -Headers $headers
                    if ($resp -ne $null) { Write-Info "Plan updated." }
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: planner plan delete <id> [--force]"
                        return
                    }
                    $etag = Get-ArgValue $parsed.Map "etag"
                    if (-not $etag) { $etag = Get-PlannerEtagLocal ("/planner/plans/" + $id) }
                    $headers = @{}
                    if ($etag) { $headers["If-Match"] = $etag }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host "Type DELETE to confirm"
                        if ($confirm -ne "DELETE") {
                            Write-Info "Canceled."
                            return
                        }
                    }
                    $resp = Invoke-GraphRequest -Method "DELETE" -Uri ("/planner/plans/" + $id) -Headers $headers
                    if ($resp -ne $null) { Write-Info "Plan deleted." }
                }
                default {
                    Write-Warn "Usage: planner plan list|get|create|update|delete"
                }
            }
        }
        "bucket" {
            switch ($action) {
                "list" {
                    $plan = Get-ArgValue $parsed.Map "plan"
                    if (-not $plan) {
                        Write-Warn "Usage: planner bucket list --plan <planId>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/planner/plans/" + $plan + "/buckets")
                    if ($resp -and $resp.value) {
                        Write-GraphTable $resp.value @("Id", "Name", "OrderHint")
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 10
                    }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: planner bucket get <id>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/planner/buckets/" + $id)
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "create" {
                    $name = Get-ArgValue $parsed.Map "name"
                    $plan = Get-ArgValue $parsed.Map "plan"
                    $order = Get-ArgValue $parsed.Map "orderHint"
                    if (-not $order) { $order = " !" }
                    $jsonRaw = Get-ArgValue $parsed.Map "json"
                    $body = if ($jsonRaw) { Parse-Value $jsonRaw } else { @{ name = $name; planId = $plan; orderHint = $order } }
                    if (-not $body -or -not $body.name -or -not $body.planId) {
                        Write-Warn "Usage: planner bucket create --plan <planId> --name <text> [--orderHint <text>] OR --json <payload>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "POST" -Uri "/planner/buckets" -Body $body
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: planner bucket update <id> --json <payload> OR --set key=value"
                        return
                    }
                    $etag = Get-ArgValue $parsed.Map "etag"
                    if (-not $etag) { $etag = Get-PlannerEtagLocal ("/planner/buckets/" + $id) }
                    $headers = @{}
                    if ($etag) { $headers["If-Match"] = $etag }
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri ("/planner/buckets/" + $id) -Body $body -Headers $headers
                    if ($resp -ne $null) { Write-Info "Bucket updated." }
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: planner bucket delete <id> [--force]"
                        return
                    }
                    $etag = Get-ArgValue $parsed.Map "etag"
                    if (-not $etag) { $etag = Get-PlannerEtagLocal ("/planner/buckets/" + $id) }
                    $headers = @{}
                    if ($etag) { $headers["If-Match"] = $etag }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host "Type DELETE to confirm"
                        if ($confirm -ne "DELETE") {
                            Write-Info "Canceled."
                            return
                        }
                    }
                    $resp = Invoke-GraphRequest -Method "DELETE" -Uri ("/planner/buckets/" + $id) -Headers $headers
                    if ($resp -ne $null) { Write-Info "Bucket deleted." }
                }
                default {
                    Write-Warn "Usage: planner bucket list|get|create|update|delete"
                }
            }
        }
        "task" {
            switch ($action) {
                "list" {
                    $plan = Get-ArgValue $parsed.Map "plan"
                    $bucket = Get-ArgValue $parsed.Map "bucket"
                    if (-not $plan -and -not $bucket) {
                        Write-Warn "Usage: planner task list --plan <planId> OR --bucket <bucketId>"
                        return
                    }
                    $path = if ($bucket) { "/planner/buckets/" + $bucket + "/tasks" } else { "/planner/plans/" + $plan + "/tasks" }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri $path
                    if ($resp -and $resp.value) {
                        Write-GraphTable $resp.value @("Id", "Title", "BucketId", "PercentComplete")
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 10
                    }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: planner task get <id>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/planner/tasks/" + $id)
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "create" {
                    $title = Get-ArgValue $parsed.Map "title"
                    $plan = Get-ArgValue $parsed.Map "plan"
                    $bucket = Get-ArgValue $parsed.Map "bucket"
                    $assign = Get-ArgValue $parsed.Map "assign"
                    $jsonRaw = Get-ArgValue $parsed.Map "json"
                    if ($jsonRaw) {
                        $body = Parse-Value $jsonRaw
                    } else {
                        $body = @{ title = $title; planId = $plan }
                        if ($bucket) { $body.bucketId = $bucket }
                        if ($assign) {
                            $assignments = @{}
                            foreach ($u in (Parse-CommaList $assign)) {
                                $assignments[$u] = @{ orderHint = " !" }
                            }
                            $body.assignments = $assignments
                        }
                    }
                    if (-not $body -or -not $body.title -or -not $body.planId) {
                        Write-Warn "Usage: planner task create --plan <planId> --title <text> [--bucket <bucketId>] OR --json <payload>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "POST" -Uri "/planner/tasks" -Body $body
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: planner task update <id> --json <payload> OR --set key=value"
                        return
                    }
                    $etag = Get-ArgValue $parsed.Map "etag"
                    if (-not $etag) { $etag = Get-PlannerEtagLocal ("/planner/tasks/" + $id) }
                    $headers = @{}
                    if ($etag) { $headers["If-Match"] = $etag }
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri ("/planner/tasks/" + $id) -Body $body -Headers $headers
                    if ($resp -ne $null) { Write-Info "Task updated." }
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: planner task delete <id> [--force]"
                        return
                    }
                    $etag = Get-ArgValue $parsed.Map "etag"
                    if (-not $etag) { $etag = Get-PlannerEtagLocal ("/planner/tasks/" + $id) }
                    $headers = @{}
                    if ($etag) { $headers["If-Match"] = $etag }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host "Type DELETE to confirm"
                        if ($confirm -ne "DELETE") {
                            Write-Info "Canceled."
                            return
                        }
                    }
                    $resp = Invoke-GraphRequest -Method "DELETE" -Uri ("/planner/tasks/" + $id) -Headers $headers
                    if ($resp -ne $null) { Write-Info "Task deleted." }
                }
                default {
                    Write-Warn "Usage: planner task list|get|create|update|delete"
                }
            }
        }
        default {
            Write-Warn "Usage: planner plan|bucket|task ..."
        }
    }
}


function Resolve-ExcelItemPath {
    param(
        [hashtable]$Map
    )
    $item = Get-ArgValue $Map "item"
    $path = Get-ArgValue $Map "path"
    $user = Get-ArgValue $Map "user"
    if ($item) {
        if ($item.StartsWith("/")) { return $item }
        $baseItem = Resolve-DriveBase $Map
        if (-not $baseItem) {
            if ($user) {
                $seg = Resolve-UserSegment $user
                if ($seg) { $baseItem = $seg + "/drive" }
            }
        }
        if (-not $baseItem) { $baseItem = "/me/drive" }
        return ($baseItem + "/items/" + $item)
    }
    if (-not $path) { return $null }
    $base = Resolve-DriveBase $Map
    if (-not $base) {
        if ($user) {
            $seg = Resolve-UserSegment $user
            if ($seg) { $base = $seg + "/drive" }
        }
    }
    if (-not $base) { $base = "/me/drive" }
    $p = Normalize-DrivePath $path
    return ($base + "/root:/" + $p + ":")
}



function Handle-ExcelCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: excel workbook|worksheet|table|range|cell ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    if (-not $rest -or $rest.Count -eq 0) {
        Write-Warn "Usage: excel workbook|worksheet|table|range|cell ..."
        return
    }
    $action = $rest[0].ToLowerInvariant()
    $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest2
    $itemPath = Resolve-ExcelItemPath $parsed.Map

    if ($sub -eq "workbook") {
        if ($action -ne "list" -and $action -ne "get") {
            Write-Warn "Usage: excel workbook list|get ..."
            return
        }
        $base = Resolve-DriveBase $parsed.Map
        if (-not $base) { $base = "/me/drive" }
        if ($action -eq "list") {
            $qh = Build-QueryAndHeaders $parsed.Map @("id","name","webUrl","file")
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/root/children" + $qh.Query) -Headers $qh.Headers
            if ($resp -and $resp.value) {
                $resp.value | Select-Object Id, Name, WebUrl, File | Format-Table -AutoSize
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
            return
        }
        if ($action -eq "get") {
            $item = Resolve-DriveItemId -Base $base -ItemId (Get-ArgValue $parsed.Map "item") -Path (Get-ArgValue $parsed.Map "path")
            if (-not $item) {
                Write-Warn "Usage: excel workbook get --item <id> OR --path <path>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/items/" + $item + "/workbook")
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            return
        }
    }

    if (-not $itemPath) {
        Write-Warn "Workbook item required: use --item <id> or --path <path>"
        return
    }

    switch ($sub) {
        "worksheet" {
            $base = $itemPath + "/workbook/worksheets"
            switch ($action) {
                "list" {
                    $resp = Invoke-GraphRequest -Method "GET" -Uri $base
                    if ($resp -and $resp.value) {
                        Write-GraphTable $resp.value @("Id", "Name", "Position", "Visibility")
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 10
                    }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: excel worksheet get <id|name> --item <id>|--path <path>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "create" {
                    $name = Get-ArgValue $parsed.Map "name"
                    $jsonRaw = Get-ArgValue $parsed.Map "json"
                    $body = if ($jsonRaw) { Parse-Value $jsonRaw } else { @{ name = $name } }
                    if (-not $body -or -not $body.name) {
                        Write-Warn "Usage: excel worksheet create --name <text> OR --json <payload>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: excel worksheet update <id|name> --json <payload> OR --set key=value"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
                    if ($resp -ne $null) { Write-Info "Worksheet updated." }
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: excel worksheet delete <id|name> [--force]"
                        return
                    }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host "Type DELETE to confirm"
                        if ($confirm -ne "DELETE") {
                            Write-Info "Canceled."
                            return
                        }
                    }
                    $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($base + "/" + $id)
                    if ($resp -ne $null) { Write-Info "Worksheet deleted." }
                }
                default {
                    Write-Warn "Usage: excel worksheet list|get|create|update|delete"
                }
            }
        }
        "table" {
            $worksheet = Get-ArgValue $parsed.Map "worksheet"
            $base = if ($worksheet) { $itemPath + "/workbook/worksheets/" + $worksheet + "/tables" } else { $itemPath + "/workbook/tables" }
            switch ($action) {
                "list" {
                    $resp = Invoke-GraphRequest -Method "GET" -Uri $base
                    if ($resp -and $resp.value) {
                        Write-GraphTable $resp.value @("Id", "Name", "ShowHeaders", "ShowTotals")
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 10
                    }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: excel table get <id|name>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ($itemPath + "/workbook/tables/" + $id)
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "create" {
                    $jsonRaw = Get-ArgValue $parsed.Map "json"
                    if (-not $jsonRaw) {
                        Write-Warn "Usage: excel table create --json <payload>"
                        return
                    }
                    $body = Parse-Value $jsonRaw
                    $resp = Invoke-GraphRequest -Method "POST" -Uri ($itemPath + "/workbook/tables/add") -Body $body
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: excel table update <id|name> --json <payload> OR --set key=value"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($itemPath + "/workbook/tables/" + $id) -Body $body
                    if ($resp -ne $null) { Write-Info "Table updated." }
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: excel table delete <id|name> [--force]"
                        return
                    }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host "Type DELETE to confirm"
                        if ($confirm -ne "DELETE") {
                            Write-Info "Canceled."
                            return
                        }
                    }
                    $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($itemPath + "/workbook/tables/" + $id)
                    if ($resp -ne $null) { Write-Info "Table deleted." }
                }
                default {
                    Write-Warn "Usage: excel table list|get|create|update|delete"
                }
            }
        }
        "range" {
            $address = Get-ArgValue $parsed.Map "address"
            if (-not $address) {
                Write-Warn "Usage: excel range get|update --address <A1>"
                return
            }
            $base = $itemPath + "/workbook/worksheets/`$`($null)"
            $sheet = Get-ArgValue $parsed.Map "worksheet"
            $sheetSeg = if ($sheet) { "/workbook/worksheets/" + $sheet } else { "/workbook" }
            $rangePath = $itemPath + $sheetSeg + "/range(address='" + $address + "')"
            switch ($action) {
                "get" {
                    $resp = Invoke-GraphRequest -Method "GET" -Uri $rangePath
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "update" {
                    $jsonRaw = Get-ArgValue $parsed.Map "json"
                    $values = Get-ArgValue $parsed.Map "values"
                    $body = if ($jsonRaw) { Parse-Value $jsonRaw } elseif ($values) { @{ values = (Parse-Value $values) } } else { $null }
                    if (-not $body) {
                        Write-Warn "Usage: excel range update --address <A1> --values <jsonArray> OR --json <payload>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri $rangePath -Body $body
                    if ($resp -ne $null) { Write-Info "Range updated." }
                }
                default {
                    Write-Warn "Usage: excel range get|update --address <A1>"
                }
            }
        }
        "cell" {
            $row = Get-ArgValue $parsed.Map "row"
            $col = Get-ArgValue $parsed.Map "col"
            if (-not $row -or -not $col) {
                Write-Warn "Usage: excel cell get|update --row <n> --col <n>"
                return
            }
            $sheet = Get-ArgValue $parsed.Map "worksheet"
            $sheetSeg = if ($sheet) { "/workbook/worksheets/" + $sheet } else { "/workbook" }
            $rangePath = $itemPath + $sheetSeg + "/cell(row=" + $row + ",column=" + $col + ")"
            switch ($action) {
                "get" {
                    $resp = Invoke-GraphRequest -Method "GET" -Uri $rangePath
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "update" {
                    $jsonRaw = Get-ArgValue $parsed.Map "json"
                    $values = Get-ArgValue $parsed.Map "values"
                    $body = if ($jsonRaw) { Parse-Value $jsonRaw } elseif ($values) { @{ values = (Parse-Value $values) } } else { $null }
                    if (-not $body) {
                        Write-Warn "Usage: excel cell update --row <n> --col <n> --values <jsonArray> OR --json <payload>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri $rangePath -Body $body
                    if ($resp -ne $null) { Write-Info "Cell updated." }
                }
                default {
                    Write-Warn "Usage: excel cell get|update --row <n> --col <n>"
                }
            }
        }
        default {
            Write-Warn "Usage: excel workbook|worksheet|table|range|cell ..."
        }
    }
}



function Handle-OneDriveCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: onedrive share <same args as: file share ...>"
        return
    }
    $sub = $Args[0].ToLowerInvariant()
    if ($sub -ne "share") {
        Write-Warn "Usage: onedrive share <same args as: file share ...>"
        return
    }
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    Handle-FileCommand (@("share") + $rest)
}



