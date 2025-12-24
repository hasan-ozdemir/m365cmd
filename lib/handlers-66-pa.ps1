# Handler: Pa
# Purpose: Power Apps helpers.
function Get-PaResource {
    return "https://api.powerapps.com"
}


function Get-PaToken {
    $scope = "https://api.powerapps.com/.default"
    $token = Get-DelegatedToken -Scope $scope
    if (-not $token) { $token = Get-AppToken -Scope $scope }
    return $token
}


function Invoke-PaRequest {
    param(
        [string]$Method,
        [string]$Path,
        [object]$Body,
        [hashtable]$Headers,
        [switch]$AllowNullResponse
    )
    $token = Get-PaToken
    if (-not $token) {
        Write-Warn "Power Apps token missing. Configure auth.app.* or sign in for delegated token."
        return $null
    }
    $base = Get-PaResource
    $url = if ($Path -match "^https?://") { $Path } else { $base + $Path }
    $hdr = @{ Authorization = "Bearer " + $token; accept = "application/json" }
    if ($Headers) { foreach ($k in $Headers.Keys) { $hdr[$k] = $Headers[$k] } }
    $params = @{ Method = $Method; Uri = $url; Headers = $hdr }
    if ($Body -ne $null) {
        $params.ContentType = "application/json"
        $params.Body = ($Body | ConvertTo-Json -Depth 10)
    }
    try {
        return Invoke-RestMethod @params
    } catch {
        if (-not $AllowNullResponse) { Write-Err $_.Exception.Message }
        return $null
    }
}


function Handle-PaCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: pa app|connector|environment <args...>"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "environment" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: pa environment list|get --name <env> OR --default"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest2
            if ($action -eq "list") {
                $resp = Invoke-PaRequest -Method "GET" -Path "/providers/Microsoft.PowerApps/environments?api-version=2017-08-01"
                if ($resp -and $resp.value) { $resp.value | ConvertTo-Json -Depth 8 }
            } elseif ($action -eq "get") {
                $name = Get-ArgValue $parsed2.Map "name"
                $isDefault = Parse-Bool (Get-ArgValue $parsed2.Map "default") $false
                if (-not $name -and -not $isDefault) {
                    Write-Warn "Usage: pa environment get --name <env> OR --default"
                    return
                }
                $envName = if ($isDefault) { "~default" } else { $name }
                $resp = Invoke-PaRequest -Method "GET" -Path ("/providers/Microsoft.PowerApps/environments/" + $envName + "?api-version=2016-11-01")
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
            } else {
                Write-Warn "Usage: pa environment list|get"
            }
        }
        "app" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: pa app list|get|remove|export|permission|owner|consent ..."
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest2
            $env = Get-ArgValue $parsed2.Map "environmentName"
            $asAdmin = Parse-Bool (Get-ArgValue $parsed2.Map "asAdmin") $false
            switch ($action) {
                "list" {
                    $path = "/providers/Microsoft.PowerApps" + ($(if ($asAdmin) { "/scopes/admin" } else { "" })) + ($(if ($env) { "/environments/" + $env } else { "" })) + "/apps?api-version=2017-08-01"
                    $resp = Invoke-PaRequest -Method "GET" -Path $path
                    if ($resp -and $resp.value) { $resp.value | ConvertTo-Json -Depth 8 }
                }
                "get" {
                    $name = Get-ArgValue $parsed2.Map "name"
                    if (-not $name) { $name = $parsed2.Positionals | Select-Object -First 1 }
                    if (-not $name) {
                        Write-Warn "Usage: pa app get --name <appId> [--environmentName <env>] [--asAdmin]"
                        return
                    }
                    $path = "/providers/Microsoft.PowerApps" + ($(if ($asAdmin) { "/scopes/admin" } else { "" })) + ($(if ($env) { "/environments/" + $env } else { "" })) + "/apps/" + $name + "?api-version=2017-08-01"
                    $resp = Invoke-PaRequest -Method "GET" -Path $path
                    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
                }
                "remove" {
                    $name = Get-ArgValue $parsed2.Map "name"
                    if (-not $name) { $name = $parsed2.Positionals | Select-Object -First 1 }
                    if (-not $name) {
                        Write-Warn "Usage: pa app remove --name <appId> [--environmentName <env>] [--asAdmin] [--force]"
                        return
                    }
                    $force = Parse-Bool (Get-ArgValue $parsed2.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host "Remove Power App '$name'? (y/N)"
                        if ($confirm.ToLowerInvariant() -notin @("y","yes")) { Write-Info "Canceled."; return }
                    }
                    $path = "/providers/Microsoft.PowerApps" + ($(if ($asAdmin) { "/scopes/admin" } else { "" })) + ($(if ($env) { "/environments/" + $env } else { "" })) + "/apps/" + $name + "?api-version=2017-08-01"
                    $resp = Invoke-PaRequest -Method "DELETE" -Path $path -AllowNullResponse
                    if ($resp -ne $null) { Write-Info "App removed." }
                }
                "export" {
                    $name = Get-ArgValue $parsed2.Map "name"
                    if (-not $name) { $name = $parsed2.Positionals | Select-Object -First 1 }
                    if (-not $env -or -not $name) {
                        Write-Warn "Usage: pa app export --environmentName <env> --name <appId> [--path <folder>] [--packageDisplayName <name>]"
                        return
                    }
                    $details = @{
                        displayName       = Get-ArgValue $parsed2.Map "packageDisplayName"
                        description       = Get-ArgValue $parsed2.Map "packageDescription"
                        creator           = Get-ArgValue $parsed2.Map "packageCreatedBy"
                        sourceEnvironment = Get-ArgValue $parsed2.Map "packageSourceEnvironment"
                    }
                    $resourcesResp = Invoke-PpRequestRaw -Method "POST" -Path ("/providers/Microsoft.BusinessAppPlatform/environments/" + $env + "/listPackageResources") -ApiVersion "2016-11-01" -Body @{
                        baseResourceIds = @("/providers/Microsoft.PowerApps/apps/" + $name)
                    }
                    if (-not $resourcesResp -or -not $resourcesResp.Body) { return }
                    $resources = $resourcesResp.Body.resources
                    foreach ($k in $resources.Keys) { $resources[$k].suggestedCreationType = "Update" }
                    $export = Invoke-PpRequestRaw -Method "POST" -Path ("/providers/Microsoft.BusinessAppPlatform/environments/" + $env + "/exportPackage") -ApiVersion "2016-11-01" -Body @{
                        includedResourceIds = @("/providers/Microsoft.PowerApps/apps/" + $name)
                        details             = $details
                        resources           = $resources
                    }
                    if (-not $export -or -not $export.Headers) { return }
                    $location = $export.Headers.Location
                    if (-not $location) { Write-Warn "Export location missing."; return }
                    $status = "Running"
                    $packageLink = $null
                    for ($i = 0; $i -lt 120 -and $status -eq "Running"; $i++) {
                        Start-Sleep -Seconds 5
                        try {
                            $resp = Invoke-RestMethod -Method Get -Uri $location -Headers @{ Authorization = "Bearer " + (Get-PpToken); accept = "application/json" }
                            $status = $resp.properties.status
                            if ($status -eq "Succeeded") { $packageLink = $resp.properties.packageLink.value }
                        } catch {
                            Write-Err $_.Exception.Message
                            return
                        }
                    }
                    if (-not $packageLink) { Write-Warn "Package export did not complete."; return }
                    $fileName = if ($details.displayName) { $details.displayName } else { $name }
                    $target = Resolve-OutputFilePath (Get-ArgValue $parsed2.Map "path") $fileName ".zip"
                    Invoke-WebRequest -Uri $packageLink -OutFile $target -Headers @{ "x-anonymous" = "true" } | Out-Null
                    Write-Info ("Saved: " + $target)
                }
                "permission" {
                    if (-not $rest2 -or $rest2.Count -eq 0) {
                        Write-Warn "Usage: pa app permission list|ensure|remove --appName <id> ..."
                        return
                    }
                    $pAction = $rest2[0].ToLowerInvariant()
                    $rest3 = if ($rest2.Count -gt 1) { $rest2[1..($rest2.Count - 1)] } else { @() }
                    $pParsed = Parse-NamedArgs $rest3
                    $appName = Get-ArgValue $pParsed.Map "appName"
                    if (-not $appName) { $appName = Get-ArgValue $pParsed.Map "name" }
                    if (-not $appName) { $appName = $pParsed.Positionals | Select-Object -First 1 }
                    $roleName = Get-ArgValue $pParsed.Map "roleName"
                    if ($asAdmin -and -not $env) {
                        Write-Warn "When using --asAdmin, --environmentName is required."
                        return
                    }
                    $base = "/providers/Microsoft.PowerApps" + ($(if ($asAdmin) { "/scopes/admin/environments/" + $env } else { "" })) + "/apps/" + $appName
                    switch ($pAction) {
                        "list" {
                            if (-not $appName) { Write-Warn "Usage: pa app permission list --appName <id>"; return }
                            $resp = Invoke-PaRequest -Method "GET" -Path ($base + "/permissions?api-version=2022-11-01")
                            if ($resp -and $resp.value) {
                                $items = $resp.value
                                if ($roleName) { $items = $items | Where-Object { $_.properties.roleName -eq $roleName } }
                                $items | ConvertTo-Json -Depth 8
                            }
                        }
                        "ensure" {
                            $role = Get-ArgValue $pParsed.Map "roleName"
                            if (-not $role) { Write-Warn "Usage: pa app permission ensure --appName <id> --roleName CanView|CanEdit"; return }
                            $userId = Get-ArgValue $pParsed.Map "userId"
                            $userName = Get-ArgValue $pParsed.Map "userName"
                            $groupId = Get-ArgValue $pParsed.Map "groupId"
                            $groupName = Get-ArgValue $pParsed.Map "groupName"
                            $tenant = Parse-Bool (Get-ArgValue $pParsed.Map "tenant") $false
                            $sendInvite = Parse-Bool (Get-ArgValue $pParsed.Map "sendInvitationMail") $false
                            $principalId = $null
                            $principalType = $null
                            if ($userId) { $principalId = $userId; $principalType = "User" }
                            elseif ($userName) { $u = Resolve-UserObject $userName; if ($u) { $principalId = $u.Id; $principalType = "User" } }
                            elseif ($groupId) { $principalId = $groupId; $principalType = "Group" }
                            elseif ($groupName) { $g = Resolve-GroupObject $groupName; if ($g) { $principalId = $g.Id; $principalType = "Group" } }
                            elseif ($tenant) {
                                $tenantId = $global:Config.tenant.tenantId
                                if (-not $tenantId) { $ctx = Get-MgContextSafe; if ($ctx) { $tenantId = $ctx.TenantId } }
                                if ($tenantId) { $principalId = $tenantId; $principalType = "Tenant" }
                            }
                            if (-not $principalId) { Write-Warn "Specify user/group/tenant for permission ensure."; return }
                            $body = @{
                                put = @(
                                    @{
                                        properties = @{
                                            principal = @{ id = $principalId; type = $principalType }
                                            NotifyShareTargetOption = ($(if ($sendInvite) { "Notify" } else { "DoNotNotify" }))
                                            roleName  = $role
                                        }
                                    }
                                )
                            }
                            $resp = Invoke-PaRequest -Method "POST" -Path ($base + "/modifyPermissions?api-version=2022-11-01") -Body $body -AllowNullResponse
                            if ($resp -ne $null) { Write-Info "Permissions updated." }
                        }
                        "remove" {
                            $userId = Get-ArgValue $pParsed.Map "userId"
                            $userName = Get-ArgValue $pParsed.Map "userName"
                            $groupId = Get-ArgValue $pParsed.Map "groupId"
                            $groupName = Get-ArgValue $pParsed.Map "groupName"
                            $tenant = Parse-Bool (Get-ArgValue $pParsed.Map "tenant") $false
                            $principalId = $null
                            if ($userId) { $principalId = $userId }
                            elseif ($userName) { $u = Resolve-UserObject $userName; if ($u) { $principalId = $u.Id } }
                            elseif ($groupId) { $principalId = $groupId }
                            elseif ($groupName) { $g = Resolve-GroupObject $groupName; if ($g) { $principalId = $g.Id } }
                            elseif ($tenant) {
                                $tenantId = $global:Config.tenant.tenantId
                                if (-not $tenantId) { $ctx = Get-MgContextSafe; if ($ctx) { $tenantId = $ctx.TenantId } }
                                if ($tenantId) { $principalId = "tenant-" + $tenantId }
                            }
                            if (-not $principalId) { Write-Warn "Specify user/group/tenant for permission remove."; return }
                            $body = @{ delete = @(@{ id = $principalId }) }
                            $resp = Invoke-PaRequest -Method "POST" -Path ($base + "/modifyPermissions?api-version=2022-11-01") -Body $body -AllowNullResponse
                            if ($resp -ne $null) { Write-Info "Permissions removed." }
                        }
                        default {
                            Write-Warn "Usage: pa app permission list|ensure|remove ..."
                        }
                    }
                }
                "owner" {
                    if (-not $rest2 -or $rest2.Count -eq 0) {
                        Write-Warn "Usage: pa app owner set --environmentName <env> --appName <id> --userId <id>|--userName <upn>"
                        return
                    }
                    $oAction = $rest2[0].ToLowerInvariant()
                    if ($oAction -ne "set") { Write-Warn "Usage: pa app owner set ..."; return }
                    $rest3 = if ($rest2.Count -gt 1) { $rest2[1..($rest2.Count - 1)] } else { @() }
                    $oParsed = Parse-NamedArgs $rest3
                    $appName = Get-ArgValue $oParsed.Map "appName"
                    $envName = Get-ArgValue $oParsed.Map "environmentName"
                    $userId = Get-ArgValue $oParsed.Map "userId"
                    $userName = Get-ArgValue $oParsed.Map "userName"
                    $roleOld = Get-ArgValue $oParsed.Map "roleForOldAppOwner"
                    if (-not $appName -or -not $envName -or (-not $userId -and -not $userName)) {
                        Write-Warn "Usage: pa app owner set --environmentName <env> --appName <id> --userId <id>|--userName <upn> [--roleForOldAppOwner CanView|CanEdit]"
                        return
                    }
                    if (-not $userId) {
                        $u = Resolve-UserObject $userName
                        if ($u) { $userId = $u.Id }
                    }
                    if (-not $userId) { Write-Warn "User not found."; return }
                    $body = @{ roleForOldAppOwner = $roleOld; newAppOwner = $userId }
                    $resp = Invoke-PaRequest -Method "POST" -Path ("/providers/Microsoft.PowerApps/scopes/admin/environments/" + $envName + "/apps/" + $appName + "/modifyAppOwner?api-version=2022-11-01") -Body $body -AllowNullResponse
                    if ($resp -ne $null) { Write-Info "Owner updated." }
                }
                "consent" {
                    if (-not $rest2 -or $rest2.Count -eq 0) {
                        Write-Warn "Usage: pa app consent set --environmentName <env> --name <appId> --bypass true|false [--force]"
                        return
                    }
                    $cAction = $rest2[0].ToLowerInvariant()
                    if ($cAction -ne "set") { Write-Warn "Usage: pa app consent set ..."; return }
                    $rest3 = if ($rest2.Count -gt 1) { $rest2[1..($rest2.Count - 1)] } else { @() }
                    $cParsed = Parse-NamedArgs $rest3
                    $appName = Get-ArgValue $cParsed.Map "name"
                    $envName = Get-ArgValue $cParsed.Map "environmentName"
                    $bypass = Get-ArgValue $cParsed.Map "bypass"
                    $force = Parse-Bool (Get-ArgValue $cParsed.Map "force") $false
                    if (-not $appName -or -not $envName -or $null -eq $bypass) {
                        Write-Warn "Usage: pa app consent set --environmentName <env> --name <appId> --bypass true|false [--force]"
                        return
                    }
                    if (-not $force) {
                        $confirm = Read-Host "Set bypass consent to '$bypass' for app '$appName'? (y/N)"
                        if ($confirm.ToLowerInvariant() -notin @("y","yes")) { Write-Info "Canceled."; return }
                    }
                    $body = @{ bypassconsent = (Parse-Bool $bypass $false) }
                    $resp = Invoke-PaRequest -Method "POST" -Path ("/providers/Microsoft.PowerApps/scopes/admin/environments/" + $envName + "/apps/" + $appName + "/setPowerAppConnectionDirectConsentBypass?api-version=2021-02-01") -Body $body -AllowNullResponse
                    if ($resp -ne $null) { Write-Info "Consent updated." }
                }
                default {
                    Write-Warn "Usage: pa app list|get|remove|export|permission|owner|consent ..."
                }
            }
        }
        "connector" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: pa connector list|export --environmentName <env>"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest2
            $env = Get-ArgValue $parsed2.Map "environmentName"
            if (-not $env) {
                Write-Warn "Usage: pa connector list|export --environmentName <env>"
                return
            }
            if ($action -eq "list") {
                $path = "/providers/Microsoft.PowerApps/apis?api-version=2016-11-01&`$filter=environment%20eq%20%27" + (Encode-QueryValue $env) + "%27%20and%20IsCustomApi%20eq%20%27True%27"
                $resp = Invoke-PaRequest -Method "GET" -Path $path
                if ($resp -and $resp.value) { $resp.value | ConvertTo-Json -Depth 8 }
            } elseif ($action -eq "export") {
                $name = Get-ArgValue $parsed2.Map "name"
                if (-not $name) { $name = $parsed2.Positionals | Select-Object -First 1 }
                $outputFolder = Get-ArgValue $parsed2.Map "outputFolder"
                if (-not $name) {
                    Write-Warn "Usage: pa connector export --environmentName <env> --name <connectorName> [--outputFolder <path>]"
                    return
                }
                $baseFolder = if ($outputFolder) { $outputFolder } else { (Get-Location).Path }
                $targetFolder = Join-Path $baseFolder $name
                if (Test-Path $targetFolder) {
                    Write-Warn "Output folder already exists: $targetFolder"
                    return
                }
                $req = Invoke-PaRequest -Method "GET" -Path ("/providers/Microsoft.PowerApps/apis/" + (Encode-QueryValue $name) + "?api-version=2016-11-01&`$filter=environment%20eq%20%27" + (Encode-QueryValue $env) + "%27%20and%20IsCustomApi%20eq%20%27True%27")
                if (-not $req -or -not $req.properties) { Write-Warn "Connector not found."; return }
                New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
                $settings = @{
                    apiDefinition      = "apiDefinition.swagger.json"
                    apiProperties      = "apiProperties.json"
                    connectorId        = $name
                    environment        = $env
                    icon               = "icon.png"
                    powerAppsApiVersion = "2016-11-01"
                    powerAppsUrl       = "https://api.powerapps.com"
                }
                $settings | ConvertTo-Json -Depth 6 | Set-Content -Path (Join-Path $targetFolder "settings.json") -Encoding ASCII
                $props = @{ properties = $req.properties }
                $whitelist = @("connectionParameters","iconBrandColor","capabilities","policyTemplateInstances")
                $filtered = [ordered]@{}
                foreach ($k in $whitelist) {
                    if ($req.properties.PSObject.Properties.Name -contains $k) {
                        $filtered[$k] = $req.properties.$k
                    }
                }
                $props = @{ properties = $filtered }
                $props | ConvertTo-Json -Depth 8 | Set-Content -Path (Join-Path $targetFolder "apiProperties.json") -Encoding ASCII
                if ($req.properties.apiDefinitions -and $req.properties.apiDefinitions.originalSwaggerUrl) {
                    $swagger = Invoke-RestMethod -Method Get -Uri $req.properties.apiDefinitions.originalSwaggerUrl -Headers @{ "x-anonymous" = "true" }
                    if ($swagger) {
                        if ($swagger -is [string]) {
                            Set-Content -Path (Join-Path $targetFolder "apiDefinition.swagger.json") -Value $swagger -Encoding ASCII
                        } else {
                            $swagger | ConvertTo-Json -Depth 20 | Set-Content -Path (Join-Path $targetFolder "apiDefinition.swagger.json") -Encoding ASCII
                        }
                    }
                }
                if ($req.properties.iconUri) {
                    Invoke-WebRequest -Method Get -Uri $req.properties.iconUri -Headers @{ "x-anonymous" = "true" } -OutFile (Join-Path $targetFolder "icon.png") | Out-Null
                }
                Write-Info ("Exported connector to: " + $targetFolder)
            } else {
                Write-Warn "Usage: pa connector list|export --environmentName <env>"
            }
        }
        default {
            Write-Warn "Usage: pa app|connector|environment <args...>"
        }
    }
}
