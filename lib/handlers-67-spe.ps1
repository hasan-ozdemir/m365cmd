# Handler: SPE
# Purpose: SharePoint Embedded helpers.
function Get-SpeContainerTypeIdByName {
    param([string]$Name)
    if (-not $Name) { return $null }
    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/storage/fileStorage/containerTypes?$filter=name eq '" + (Escape-ODataString $Name) + "'") -Beta
    if ($resp -and $resp.value) {
        return ($resp.value | Select-Object -First 1).id
    }
    return $null
}


function Handle-SpeCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: spe container|containertype <args...>"
        return
    }
    if (-not (Require-GraphConnection)) { return }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }

    switch ($sub) {
        "containertype" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: spe containertype list|get|add|remove"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed = Parse-NamedArgs $rest2
            switch ($action) {
                "list" {
                    $resp = Invoke-GraphRequest -Method "GET" -Uri "/storage/fileStorage/containerTypes" -Beta
                    if ($resp -and $resp.value) { $resp.value | ConvertTo-Json -Depth 8 }
                }
                "get" {
                    $id = Get-ArgValue $parsed.Map "id"
                    $name = Get-ArgValue $parsed.Map "name"
                    if ($name -and -not $id) { $id = Get-SpeContainerTypeIdByName $name }
                    if (-not $id) { Write-Warn "Usage: spe containertype get --id <id> OR --name <name>"; return }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/storage/fileStorage/containerTypes/" + $id) -Beta
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                }
                "add" {
                    $jsonRaw = Get-ArgValue $parsed.Map "json"
                    if (-not $jsonRaw) {
                        Write-Warn "Usage: spe containertype add --json <payload>"
                        return
                    }
                    $body = Parse-Value $jsonRaw
                    $resp = Invoke-GraphRequest -Method "POST" -Uri "/storage/fileStorage/containerTypes" -Body $body -Beta
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                }
                "remove" {
                    $id = Get-ArgValue $parsed.Map "id"
                    $name = Get-ArgValue $parsed.Map "name"
                    if ($name -and -not $id) { $id = Get-SpeContainerTypeIdByName $name }
                    if (-not $id) { Write-Warn "Usage: spe containertype remove --id <id> OR --name <name> [--force]"; return }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host "Type DELETE to confirm"
                        if ($confirm -ne "DELETE") { Write-Info "Canceled."; return }
                    }
                    $resp = Invoke-GraphRequest -Method "DELETE" -Uri ("/storage/fileStorage/containerTypes/" + $id) -Beta
                    if ($resp -ne $null) { Write-Info "Container type removed." }
                }
                default {
                    Write-Warn "Usage: spe containertype list|get|add|remove"
                }
            }
        }
        "container" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: spe container list|get|add|remove|activate|permission|recyclebinitem"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed = Parse-NamedArgs $rest2
            switch ($action) {
                "list" {
                    $typeId = Get-ArgValue $parsed.Map "containerTypeId"
                    $typeName = Get-ArgValue $parsed.Map "containerTypeName"
                    if (-not $typeId -and $typeName) { $typeId = Get-SpeContainerTypeIdByName $typeName }
                    if (-not $typeId) {
                        Write-Warn "Usage: spe container list --containerTypeId <id> OR --containerTypeName <name>"
                        return
                    }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/storage/fileStorage/containers?$filter=containerTypeId eq " + $typeId)
                    if ($resp -and $resp.value) { $resp.value | ConvertTo-Json -Depth 8 }
                }
                "get" {
                    $id = Get-ArgValue $parsed.Map "id"
                    if (-not $id) { $id = $parsed.Positionals | Select-Object -First 1 }
                    if (-not $id) { Write-Warn "Usage: spe container get <id>"; return }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/storage/fileStorage/containers/" + $id)
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                }
                "add" {
                    $jsonRaw = Get-ArgValue $parsed.Map "json"
                    if (-not $jsonRaw) { Write-Warn "Usage: spe container add --json <payload>"; return }
                    $body = Parse-Value $jsonRaw
                    $resp = Invoke-GraphRequest -Method "POST" -Uri "/storage/fileStorage/containers" -Body $body
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                }
                "remove" {
                    $id = Get-ArgValue $parsed.Map "id"
                    if (-not $id) { $id = $parsed.Positionals | Select-Object -First 1 }
                    if (-not $id) { Write-Warn "Usage: spe container remove <id> [--force]"; return }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host "Type DELETE to confirm"
                        if ($confirm -ne "DELETE") { Write-Info "Canceled."; return }
                    }
                    $resp = Invoke-GraphRequest -Method "DELETE" -Uri ("/storage/fileStorage/containers/" + $id)
                    if ($resp -ne $null) { Write-Info "Container removed." }
                }
                "activate" {
                    $id = Get-ArgValue $parsed.Map "id"
                    if (-not $id) { $id = $parsed.Positionals | Select-Object -First 1 }
                    if (-not $id) { Write-Warn "Usage: spe container activate --id <id>"; return }
                    $resp = Invoke-GraphRequest -Method "POST" -Uri ("/storage/fileStorage/containers/" + $id + "/activate")
                    if ($resp -ne $null) { Write-Info "Container activated." }
                }
                "permission" {
                    $action2 = $parsed.Positionals | Select-Object -First 1
                    if ($action2 -ne "list") {
                        Write-Warn "Usage: spe container permission list --containerId <id>"
                        return
                    }
                    $id = Get-ArgValue $parsed.Map "containerId"
                    if (-not $id) { $id = $parsed.Positionals | Select-Object -Skip 1 -First 1 }
                    if (-not $id) { Write-Warn "Usage: spe container permission list --containerId <id>"; return }
                    $resp = Invoke-GraphRequest -Method "GET" -Uri ("/storage/fileStorage/containers/" + $id + "/permissions")
                    if ($resp -and $resp.value) { $resp.value | ConvertTo-Json -Depth 8 }
                }
                "recyclebinitem" {
                    $action2 = $parsed.Positionals | Select-Object -First 1
                    if (-not $action2) { Write-Warn "Usage: spe container recyclebinitem list|restore ..."; return }
                    if ($action2 -eq "list") {
                        $typeId = Get-ArgValue $parsed.Map "containerTypeId"
                        $typeName = Get-ArgValue $parsed.Map "containerTypeName"
                        if (-not $typeId -and $typeName) { $typeId = Get-SpeContainerTypeIdByName $typeName }
                        if (-not $typeId) { Write-Warn "Usage: spe container recyclebinitem list --containerTypeId <id>|--containerTypeName <name>"; return }
                        $resp = Invoke-GraphRequest -Method "GET" -Uri ("/storage/fileStorage/deletedContainers?$filter=containerTypeId eq " + $typeId)
                        if ($resp -and $resp.value) { $resp.value | ConvertTo-Json -Depth 8 }
                    } elseif ($action2 -eq "restore") {
                        $id = Get-ArgValue $parsed.Map "id"
                        if (-not $id) { $id = Get-ArgValue $parsed.Map "containerId" }
                        if (-not $id) {
                            Write-Warn "Usage: spe container recyclebinitem restore --id <id>"
                            return
                        }
                        $resp = Invoke-GraphRequest -Method "POST" -Uri ("/storage/fileStorage/deletedContainers/" + $id + "/restore")
                        if ($resp -ne $null) { Write-Info "Container restored." }
                    } else {
                        Write-Warn "Usage: spe container recyclebinitem list|restore ..."
                    }
                }
                default {
                    Write-Warn "Usage: spe container list|get|add|remove|activate|permission|recyclebinitem"
                }
            }
        }
        default {
            Write-Warn "Usage: spe container|containertype <args...>"
        }
    }
}
