# Handler: Sharepoint
# Purpose: Sharepoint command handlers.
function Resolve-SiteId {
    param([string]$Identity)
    if (-not $Identity) { return $null }
    if ($Identity -match "^https?://") {
        try {
            $u = [System.Uri]$Identity
            $host = $u.Host
            $path = $u.AbsolutePath
            $apiPath = "/sites/" + $host + ":" + $path
            $resp = Invoke-GraphRequest -Method "GET" -Uri $apiPath
            if ($resp -and $resp.id) { return $resp.id }
        } catch {}
    }
    if ($Identity -match "^[^/]+:.+:") {
        $resp = Invoke-GraphRequest -Method "GET" -Uri ("/sites/" + $Identity)
        if ($resp -and $resp.id) { return $resp.id }
        return $null
    }
    return $Identity
}



function Handle-SiteCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: site list|get ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "list" {
            $search = Get-ArgValue $parsed.Map "search"
            if (-not $search) {
                Write-Warn "Usage: site list --search <text>"
                return
            }
            $top = Get-ArgValue $parsed.Map "top"
            $query = "?search=" + (Encode-QueryValue $search)
            if ($top) { $query += "&`$top=" + $top }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ("/sites" + $query)
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Id", "DisplayName", "WebUrl")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: site get <siteId|hostname:/path:>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ("/sites/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        default {
            Write-Warn "Usage: site list|get"
        }
    }
}



function Handle-SPListCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: splist list|get|create|update|delete|delta|item"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }

    if ($sub -eq "item" -or $sub -eq "items") {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: splist item list|get|create|update|delete --site <siteId> --list <listId>"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed = Parse-NamedArgs $rest2
        $siteId = Resolve-SiteId (Get-ArgValue $parsed.Map "site")
        $listId = Get-ArgValue $parsed.Map "list"
        if (-not $siteId -or -not $listId) {
            Write-Warn "Usage: splist item <action> --site <siteId|url|hostname:/path:> --list <listId>"
            return
        }
        $base = "/sites/" + $siteId + "/lists/" + $listId + "/items"
        switch ($action) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed.Map @()
                $expand = if ($qh.Query -match "\\?") { $qh.Query + "&`$expand=fields" } else { "?`$expand=fields" }
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $expand) -Headers $qh.Headers
                if ($resp -and $resp.value) {
                    $resp.value | ConvertTo-Json -Depth 10
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 10
                }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: splist item get <itemId> --site <siteId> --list <listId>"
                    return
                }
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id + "?`$expand=fields")
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "create" {
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                $fieldsRaw = Get-ArgValue $parsed.Map "fields"
                $body = $null
                if ($jsonRaw) {
                    $body = Parse-Value $jsonRaw
                } elseif ($fieldsRaw) {
                    $body = @{ fields = (Parse-KvPairs $fieldsRaw) }
                }
                if (-not $body) {
                    Write-Warn "Usage: splist item create --site <siteId> --list <listId> --fields key=value OR --json <payload>"
                    return
                }
                $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
            }
            "update" {
                $id = $parsed.Positionals | Select-Object -First 1
                $jsonRaw = Get-ArgValue $parsed.Map "json"
                $fieldsRaw = Get-ArgValue $parsed.Map "fields"
                $body = $null
                if ($jsonRaw) {
                    $body = Parse-Value $jsonRaw
                } elseif ($fieldsRaw) {
                    $body = Parse-KvPairs $fieldsRaw
                }
                if (-not $id -or -not $body) {
                    Write-Warn "Usage: splist item update <itemId> --site <siteId> --list <listId> --fields key=value OR --json <payload>"
                    return
                }
                if ($body.fields) {
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
                } else {
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id + "/fields") -Body $body
                }
                if ($resp -ne $null) { Write-Info "List item updated." }
            }
            "delete" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: splist item delete <itemId> --site <siteId> --list <listId>"
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
                if ($resp -ne $null) { Write-Info "List item deleted." }
            }
            default {
                Write-Warn "Usage: splist item list|get|create|update|delete --site <siteId> --list <listId>"
            }
        }
        return
    }

    if ($sub -eq "delta") {
        $parsed = Parse-NamedArgs $rest
        $siteId = Resolve-SiteId (Get-ArgValue $parsed.Map "site")
        $listId = Get-ArgValue $parsed.Map "list"
        $token = Get-ArgValue $parsed.Map "token"
        if (-not $siteId -or -not $listId) {
            Write-Warn "Usage: splist delta --site <siteId|url|hostname:/path:> --list <listId> [--token <deltaLink>]"
            return
        }
        $base = "/sites/" + $siteId + "/lists/" + $listId + "/items/delta"
        $uri = if ($token) { $token } else { $base }
        $qh = if (-not $token) { Build-QueryAndHeaders $parsed.Map @() } else { $null }
        $final = if ($qh) { $uri + $qh.Query } else { $uri }
        $resp = Invoke-GraphRequest -Method "GET" -Uri $final
        if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        return
    }

    $action = $sub
    $parsed = Parse-NamedArgs $rest
    $siteId = Resolve-SiteId (Get-ArgValue $parsed.Map "site")
    if (-not $siteId) {
        Write-Warn "Usage: splist <action> --site <siteId|url|hostname:/path:>"
        return
    }
    $base = "/sites/" + $siteId + "/lists"
    switch ($action) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @("id","displayName","list")
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
                Write-Warn "Usage: splist get <listId> --site <siteId>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "create" {
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            $name = Get-ArgValue $parsed.Map "name"
            $template = Get-ArgValue $parsed.Map "template"
            if (-not $template) { $template = "genericList" }
            $body = if ($jsonRaw) { Parse-Value $jsonRaw } else { @{ displayName = $name; list = @{ template = $template } } }
            if (-not $body -or -not $body.displayName) {
                Write-Warn "Usage: splist create --site <siteId> --name <text> [--template genericList] OR --json <payload>"
                return
            }
            $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $id -or -not $body) {
                Write-Warn "Usage: splist update <listId> --site <siteId> --json <payload> OR --set key=value"
                return
            }
            $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
            if ($resp -ne $null) { Write-Info "List updated." }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: splist delete <listId> --site <siteId>"
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
            if ($resp -ne $null) { Write-Info "List deleted." }
        }
        default {
            Write-Warn "Usage: splist list|get|create|update|delete|delta --site <siteId|url|hostname:/path:>"
        }
    }
}


function Handle-SPPageCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: spage list|get|create|update|delete|publish --site <siteId|url|hostname:/path:>"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $siteId = Resolve-SiteId (Get-ArgValue $parsed.Map "site")
    if (-not $siteId) {
        Write-Warn "Usage: spage <action> --site <siteId|url|hostname:/path:>"
        return
    }
    $base = "/sites/" + $siteId + "/pages"

    switch ($action) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @("id","name","title","webUrl","createdDateTime")
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Id","Name","Title","WebUrl","CreatedDateTime")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: spage get <pageId> --site <siteId>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "create" {
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            $name = Get-ArgValue $parsed.Map "name"
            $title = Get-ArgValue $parsed.Map "title"
            $news = Parse-Bool (Get-ArgValue $parsed.Map "news") $false
            $body = $null
            if ($jsonRaw) {
                $body = Parse-Value $jsonRaw
            } else {
                if (-not $name -or -not $title) {
                    Write-Warn "Usage: spage create --site <siteId> --name <page.aspx> --title <text> [--news true|false]"
                    return
                }
                $body = @{
                    "@odata.type" = "microsoft.graph.sitePage"
                    name          = $name
                    title         = $title
                }
                if ($news) { $body.promotionKind = "newsPost" }
            }
            $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $id -or -not $body) {
                Write-Warn "Usage: spage update <pageId> --site <siteId> --json <payload> OR --set key=value"
                return
            }
            $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
            if ($resp -ne $null) { Write-Info "Page updated." }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: spage delete <pageId> --site <siteId>"
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
            if ($resp -ne $null) { Write-Info "Page deleted." }
        }
        "publish" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: spage publish <pageId> --site <siteId>"
                return
            }
            $resp = Invoke-GraphRequest -Method "POST" -Uri ($base + "/" + $id + "/publish")
            if ($resp -ne $null) { Write-Info "Page published." }
        }
        default {
            Write-Warn "Usage: spage list|get|create|update|delete|publish --site <siteId|url|hostname:/path:>"
        }
    }
}


function Handle-SPColumnCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: spcolumn list|get|create|update|delete --site <siteId|url|hostname:/path:> [--list <listId>]"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $siteId = Resolve-SiteId (Get-ArgValue $parsed.Map "site")
    $listId = Get-ArgValue $parsed.Map "list"
    if (-not $siteId) {
        Write-Warn "Usage: spcolumn <action> --site <siteId|url|hostname:/path:> [--list <listId>]"
        return
    }
    $base = if ($listId) {
        "/sites/" + $siteId + "/lists/" + $listId + "/columns"
    } else {
        "/sites/" + $siteId + "/columns"
    }

    switch ($action) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @("id","name","displayName","hidden","required")
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Id","Name","DisplayName","Hidden","Required")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: spcolumn get <columnId> --site <siteId> [--list <listId>]"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "create" {
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $body) {
                Write-Warn "Usage: spcolumn create --site <siteId> [--list <listId>] --json <payload> OR --set key=value"
                return
            }
            $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $id -or -not $body) {
                Write-Warn "Usage: spcolumn update <columnId> --site <siteId> [--list <listId>] --json <payload> OR --set key=value"
                return
            }
            $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
            if ($resp -ne $null) { Write-Info "Column updated." }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: spcolumn delete <columnId> --site <siteId> [--list <listId>]"
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
            if ($resp -ne $null) { Write-Info "Column deleted." }
        }
        default {
            Write-Warn "Usage: spcolumn list|get|create|update|delete --site <siteId|url|hostname:/path:> [--list <listId>]"
        }
    }
}


function Handle-SPContentTypeCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: spctype list|get|create|update|delete --site <siteId|url|hostname:/path:> [--list <listId>]"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $siteId = Resolve-SiteId (Get-ArgValue $parsed.Map "site")
    $listId = Get-ArgValue $parsed.Map "list"
    if (-not $siteId) {
        Write-Warn "Usage: spctype <action> --site <siteId|url|hostname:/path:> [--list <listId>]"
        return
    }
    $base = if ($listId) {
        "/sites/" + $siteId + "/lists/" + $listId + "/contentTypes"
    } else {
        "/sites/" + $siteId + "/contentTypes"
    }

    switch ($action) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @("id","name","description","group","isHidden","isReadOnly")
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Id","Name","Group","IsHidden","IsReadOnly")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: spctype get <contentTypeId> --site <siteId> [--list <listId>]"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "create" {
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $body) {
                Write-Warn "Usage: spctype create --site <siteId> [--list <listId>] --json <payload> OR --set key=value"
                return
            }
            $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $id -or -not $body) {
                Write-Warn "Usage: spctype update <contentTypeId> --site <siteId> [--list <listId>] --json <payload> OR --set key=value"
                return
            }
            $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
            if ($resp -ne $null) { Write-Info "Content type updated." }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: spctype delete <contentTypeId> --site <siteId> [--list <listId>]"
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
            if ($resp -ne $null) { Write-Info "Content type deleted." }
        }
        default {
            Write-Warn "Usage: spctype list|get|create|update|delete --site <siteId|url|hostname:/path:> [--list <listId>]"
        }
    }
}


function Handle-SPPermissionCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: spperm list|get|grant|delete --site <siteId|url|hostname:/path:>"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $siteId = Resolve-SiteId (Get-ArgValue $parsed.Map "site")
    if (-not $siteId) {
        Write-Warn "Usage: spperm <action> --site <siteId|url|hostname:/path:>"
        return
    }
    $base = "/sites/" + $siteId + "/permissions"

    switch ($action) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @()
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: spperm get <permissionId> --site <siteId>"
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "grant" {
            $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if (-not $body) {
                Write-Warn "Usage: spperm grant --site <siteId> --json <payload> OR --set key=value"
                return
            }
            $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: spperm delete <permissionId> --site <siteId>"
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
            if ($resp -ne $null) { Write-Info "Permission deleted." }
        }
        default {
            Write-Warn "Usage: spperm list|get|grant|delete --site <siteId|url|hostname:/path:>"
        }
    }
}

