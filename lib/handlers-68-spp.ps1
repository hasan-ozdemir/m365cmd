# Handler: SPP
# Purpose: SharePoint Premium helpers.
function Get-ServerRelativePath {
    param([string]$WebUrl, [string]$ChildUrl)
    try {
        $base = [System.Uri]$WebUrl
        $child = [System.Uri]$ChildUrl
        return $child.AbsolutePath
    } catch {
        return $ChildUrl
    }
}


function Get-SppTokenForSite {
    param([string]$SiteUrl)
    if (-not $SiteUrl) { return $null }
    try {
        $uri = [System.Uri]$SiteUrl
        $resource = $uri.Scheme + "://" + $uri.Host
    } catch {
        return $null
    }
    $scope = $resource.TrimEnd("/") + "/.default"
    $token = Get-DelegatedToken -Scope $scope
    if (-not $token) { $token = Get-AppToken -Scope $scope }
    return $token
}


function Invoke-SppRequest {
    param(
        [string]$Method,
        [string]$Url,
        [object]$Body,
        [hashtable]$Headers,
        [switch]$AllowNullResponse
    )
    $token = Get-SppTokenForSite $Url
    if (-not $token) {
        Write-Warn "SharePoint token missing. Configure auth.app.* or delegated token."
        return $null
    }
    $hdr = @{ Authorization = "Bearer " + $token; accept = "application/json;odata=nometadata" }
    if ($Headers) { foreach ($k in $Headers.Keys) { $hdr[$k] = $Headers[$k] } }
    $params = @{ Method = $Method; Uri = $Url; Headers = $hdr }
    if ($Body -ne $null) {
        $params.ContentType = "application/json;odata=verbose"
        $params.Body = ($Body | ConvertTo-Json -Depth 10)
    }
    try {
        return Invoke-RestMethod @params
    } catch {
        if (-not $AllowNullResponse) { Write-Err $_.Exception.Message }
        return $null
    }
}


function Handle-SppCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: spp contentcenter|model <args...>"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "contentcenter" {
            $action = if ($rest.Count -gt 0) { $rest[0].ToLowerInvariant() } else { "" }
            if ($action -ne "list") {
                Write-Warn "Usage: spp contentcenter list"
                return
            }
            if (-not (Require-SpoConnection)) { return }
            try {
                Get-SPOSite -Template "CONTENTCTR#0" | Select-Object Url, Title, Template | Format-Table -AutoSize
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "model" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: spp model list|get|apply|remove --siteUrl <url>"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest2
            $siteUrl = Get-ArgValue $parsed2.Map "siteUrl"
            if (-not $siteUrl -and $action -eq "apply") { $siteUrl = Get-ArgValue $parsed2.Map "contentCenterUrl" }
            if (-not $siteUrl) {
                Write-Warn "Usage: spp model <action> --siteUrl <url>"
                return
            }
            $siteUrl = $siteUrl.TrimEnd("/")
            switch ($action) {
                "list" {
                    $resp = Invoke-SppRequest -Method "GET" -Url ($siteUrl + "/_api/machinelearning/models")
                    if ($resp -and $resp.value) { $resp.value | ConvertTo-Json -Depth 8 }
                }
                "get" {
                    $id = Get-ArgValue $parsed2.Map "id"
                    $title = Get-ArgValue $parsed2.Map "title"
                    $withPub = Parse-Bool (Get-ArgValue $parsed2.Map "withPublications") $false
                    if (-not $id -and -not $title) {
                        Write-Warn "Usage: spp model get --siteUrl <url> --id <guid> OR --title <text>"
                        return
                    }
                    $resp = Invoke-SppRequest -Method "GET" -Url ($siteUrl + "/_api/machinelearning/models")
                    if (-not $resp -or -not $resp.value) { return }
                    $model = if ($id) { $resp.value | Where-Object { $_.UniqueId -eq $id } | Select-Object -First 1 } else { $resp.value | Where-Object { $_.ModelName -like ("*" + $title + "*") } | Select-Object -First 1 }
                    if (-not $model) { Write-Warn "Model not found."; return }
                    if ($withPub) {
                        $pub = Invoke-SppRequest -Method "GET" -Url ($siteUrl + "/_api/machinelearning/publications/getbymodeluniqueid('" + $model.UniqueId + "')")
                        $model | Add-Member -NotePropertyName Publications -NotePropertyValue $pub.value -Force
                    }
                    $model | ConvertTo-Json -Depth 8
                }
                "remove" {
                    $id = Get-ArgValue $parsed2.Map "id"
                    $title = Get-ArgValue $parsed2.Map "title"
                    if (-not $id -and -not $title) {
                        Write-Warn "Usage: spp model remove --siteUrl <url> --id <guid> OR --title <text> [--force]"
                        return
                    }
                    $force = Parse-Bool (Get-ArgValue $parsed2.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host "Remove model? (y/N)"
                        if ($confirm.ToLowerInvariant() -notin @("y","yes")) { Write-Info "Canceled."; return }
                    }
                    $reqUrl = $siteUrl + "/_api/machinelearning/models/"
                    if ($title) {
                        if (-not $title.ToLowerInvariant().EndsWith(".classifier")) { $title += ".classifier" }
                        $reqUrl += "getbytitle('" + (Encode-QueryValue $title) + "')"
                    } else {
                        $reqUrl += "getbyuniqueid('" + $id + "')"
                    }
                    $resp = Invoke-SppRequest -Method "DELETE" -Url $reqUrl -Headers @{ "if-match" = "*" } -AllowNullResponse
                    if ($resp -ne $null) { Write-Info "Model removed." }
                }
                "apply" {
                    $contentCenter = $siteUrl
                    $webUrl = Get-ArgValue $parsed2.Map "webUrl"
                    $modelId = Get-ArgValue $parsed2.Map "id"
                    $modelTitle = Get-ArgValue $parsed2.Map "title"
                    $listTitle = Get-ArgValue $parsed2.Map "listTitle"
                    $listId = Get-ArgValue $parsed2.Map "listId"
                    $listUrl = Get-ArgValue $parsed2.Map "listUrl"
                    if (-not $webUrl) {
                        Write-Warn "Usage: spp model apply --contentCenterUrl <url> --webUrl <url> --id <guid>|--title <name> --listTitle|--listId|--listUrl"
                        return
                    }
                    $models = Invoke-SppRequest -Method "GET" -Url ($contentCenter + "/_api/machinelearning/models")
                    $model = $null
                    if ($models -and $models.value) {
                        if ($modelId) { $model = $models.value | Where-Object { $_.UniqueId -eq $modelId } | Select-Object -First 1 }
                        elseif ($modelTitle) { $model = $models.value | Where-Object { $_.ModelName -like ("*" + $modelTitle + "*") } | Select-Object -First 1 }
                    }
                    if (-not $model) { Write-Warn "Model not found."; return }
                    $listInfoUrl = $webUrl.TrimEnd("/") + "/_api/web"
                    if ($listId) {
                        $listInfoUrl += "/lists(guid'" + $listId + "')"
                    } elseif ($listTitle) {
                        $listInfoUrl += "/lists/getByTitle('" + (Encode-QueryValue $listTitle) + "')"
                    } elseif ($listUrl) {
                        $rel = Get-ServerRelativePath $webUrl $listUrl
                        $listInfoUrl += "/GetList('" + (Encode-QueryValue $rel) + "')"
                    } else {
                        Write-Warn "Specify listTitle, listId, or listUrl."
                        return
                    }
                    $listInfoUrl += "?`$select=BaseType,RootFolder/ServerRelativeUrl&`$expand=RootFolder"
                    $listInfo = Invoke-SppRequest -Method "GET" -Url $listInfoUrl
                    if (-not $listInfo -or $listInfo.BaseType -ne 1) {
                        Write-Warn "Target list is not a document library."
                        return
                    }
                    $viewOpt = Get-ArgValue $parsed2.Map "viewOption"
                    if (-not $viewOpt) { $viewOpt = "NewViewAsDefault" }
                    $body = @{
                        __metadata  = @{ type = "Microsoft.Office.Server.ContentCenter.SPMachineLearningPublicationsEntityData" }
                        Publications = @{
                            results = @(
                                @{
                                    ModelUniqueId = $model.UniqueId
                                    TargetSiteUrl = $webUrl
                                    TargetWebServerRelativeUrl = ([System.Uri]$webUrl).AbsolutePath
                                    TargetLibraryServerRelativeUrl = $listInfo.RootFolder.ServerRelativeUrl
                                    ViewOption = $viewOpt
                                }
                            )
                        }
                    }
                    $resp = Invoke-SppRequest -Method "POST" -Url ($contentCenter + "/_api/machinelearning/publications") -Body $body
                    if ($resp) { $resp | ConvertTo-Json -Depth 6 }
                }
                default {
                    Write-Warn "Usage: spp model list|get|apply|remove"
                }
            }
        }
        default {
            Write-Warn "Usage: spp contentcenter|model <args...>"
        }
    }
}
