# Handler: Connections
# Purpose: Connections command handlers.
function Invoke-ConnectionsNewsList {
    param(
        [string]$SiteId,
        [hashtable]$Map
    )
    if (-not $SiteId) { return }
    if (-not $Map) { $Map = @{} }

    $select = Get-ArgValue $Map "select"
    if ($select) {
        if ($select -notmatch "(^|,)promotionKind(,|$)") {
            $Map["select"] = ($select + ",promotionKind")
        }
    }

    $qh = Build-QueryAndHeaders $Map @("id","name","title","webUrl","createdDateTime","promotionKind")
    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/sites/" + $SiteId + "/pages" + $qh.Query) -Headers $qh.Headers
    if ($resp -and $resp.value) {
        $news = @($resp.value | Where-Object { $_.promotionKind -eq "newsPost" })
        if (-not $news -or $news.Count -eq 0) {
            Write-Info "No news pages found."
            return
        }
        if ($Map.ContainsKey("json")) {
            $news | ConvertTo-Json -Depth 8
        } else {
            Write-GraphTable $news @("Id","Name","Title","WebUrl","CreatedDateTime")
        }
    } elseif ($resp) {
        $resp | ConvertTo-Json -Depth 8
    }
}

function Find-ConnectionsDashboardPages {
    param(
        [string]$SiteId,
        [hashtable]$Map
    )
    if (-not $SiteId) { return @() }
    if (-not $Map) { $Map = @{} }

    $select = Get-ArgValue $Map "select"
    if ($select) {
        if ($select -notmatch "(^|,)name(,|$)") {
            $Map["select"] = ($select + ",name,title,webUrl,createdDateTime")
        }
    }

    $qh = Build-QueryAndHeaders $Map @("id","name","title","webUrl","createdDateTime")
    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/sites/" + $SiteId + "/pages" + $qh.Query) -Headers $qh.Headers
    if (-not $resp -or -not $resp.value) { return @() }
    $pages = @($resp.value)
    $dash = @()
    foreach ($p in $pages) {
        if ($p.name -and $p.name.ToLowerInvariant() -eq "dashboard.aspx") {
            $dash += $p
        } elseif ($p.title -and $p.title.ToLowerInvariant() -eq "dashboard") {
            $dash += $p
        } elseif ($p.name -and $p.name.ToLowerInvariant() -like "*dashboard*") {
            $dash += $p
        }
    }
    return $dash
}

function Handle-ConnectionsCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: connections open|home|site|news|dashboard|page"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }

    switch ($sub) {
        "open" { Write-Host "https://www.microsoft365.com/" }
        "home" {
            if (-not (Require-GraphConnection)) { return }
            $resp = Invoke-GraphRequest -Method "GET" -Uri "/sites/root"
            if ($resp) { $resp | ConvertTo-Json -Depth 6 }
        }
        "site" { Handle-SiteCommand $rest }
        "news" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: connections news list|get|create|update|delete|publish --site <siteId|url|hostname:/path:>"
                return
            }
            if (-not (Require-GraphConnection)) { return }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed = Parse-NamedArgs $rest2
            $siteId = Resolve-SiteId (Get-ArgValue $parsed.Map "site")
            if (-not $siteId) {
                Write-Warn "Usage: connections news <action> --site <siteId|url|hostname:/path:>"
                return
            }

            switch ($action) {
                "list" {
                    Invoke-ConnectionsNewsList -SiteId $siteId -Map $parsed.Map
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: connections news get <pageId> --site <siteId>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/sites/" + $siteId + "/pages/" + $id)
                    if ($resp) {
                        if ($resp.promotionKind -and $resp.promotionKind -ne "newsPost") {
                            Write-Warn "Page is not marked as a news post."
                        }
                        $resp | ConvertTo-Json -Depth 8
                    }
                }
                "create" {
                    $rest2 += @("--news","true")
                    Handle-SPPageCommand (@($action) + $rest2)
                }
                "update" { Handle-SPPageCommand (@($action) + $rest2) }
                "delete" { Handle-SPPageCommand (@($action) + $rest2) }
                "publish" { Handle-SPPageCommand (@($action) + $rest2) }
                default {
                    Write-Warn "Usage: connections news list|get|create|update|delete|publish --site <siteId|url|hostname:/path:>"
                }
            }
        }
        "dashboard" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: connections dashboard list|get|open|update|publish --site <siteId|url|hostname:/path:>"
                return
            }
            if (-not (Require-GraphConnection)) { return }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed = Parse-NamedArgs $rest2
            $siteId = Resolve-SiteId (Get-ArgValue $parsed.Map "site")
            if (-not $siteId) {
                Write-Warn "Usage: connections dashboard <action> --site <siteId|url|hostname:/path:>"
                return
            }
            $dash = Find-ConnectionsDashboardPages -SiteId $siteId -Map $parsed.Map
            switch ($action) {
                "list" {
                    if (-not $dash -or $dash.Count -eq 0) {
                        Write-Info "No dashboard pages found."
                        return
                    }
                    if ($parsed.Map.ContainsKey("json")) {
                        $dash | ConvertTo-Json -Depth 8
                    } else {
                        Write-GraphTable $dash @("Id","Name","Title","WebUrl","CreatedDateTime")
                    }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        if ($dash -and $dash.Count -gt 0) {
                            $id = $dash[0].Id
                        }
                    }
                    if (-not $id) {
                        Write-Warn "Usage: connections dashboard get <pageId> --site <siteId>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/sites/" + $siteId + "/pages/" + $id)
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                }
                "open" {
                    $target = $null
                    if ($dash -and $dash.Count -gt 0) { $target = $dash[0].webUrl }
                    if ($target) {
                        Write-Host $target
                    } else {
                        Write-Warn "No dashboard page found."
                    }
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id -and $dash -and $dash.Count -gt 0) { $id = $dash[0].Id }
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: connections dashboard update <pageId> --site <siteId> --json <payload> OR --set key=value"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "PATCH" -Uri ("/sites/" + $siteId + "/pages/" + $id) -Body $body
                    if ($resp -ne $null) { Write-Info "Dashboard updated." }
                }
                "publish" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id -and $dash -and $dash.Count -gt 0) { $id = $dash[0].Id }
                    if (-not $id) {
                        Write-Warn "Usage: connections dashboard publish <pageId> --site <siteId>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "POST" -Uri ("/sites/" + $siteId + "/pages/" + $id + "/publish")
                    if ($resp -ne $null) { Write-Info "Dashboard published." }
                }
                default {
                    Write-Warn "Usage: connections dashboard list|get|open|update|publish --site <siteId|url|hostname:/path:>"
                }
            }
        }
        "page" { Handle-SPPageCommand $rest }
        default {
            Write-Warn "Usage: connections open|home|site|news|dashboard|page"
        }
    }
}

