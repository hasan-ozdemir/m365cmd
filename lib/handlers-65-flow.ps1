# Handler: Flow
# Purpose: Power Automate (flow) helpers.
function Get-FlowResource {
    return "https://api.flow.microsoft.com"
}


function Get-FlowToken {
    $scope = "https://api.flow.microsoft.com/.default"
    $token = Get-DelegatedToken -Scope $scope
    if (-not $token) { $token = Get-AppToken -Scope $scope }
    return $token
}


function Invoke-FlowRequest {
    param(
        [string]$Method,
        [string]$Path,
        [object]$Body,
        [hashtable]$Headers,
        [switch]$AllowNullResponse
    )
    $token = Get-FlowToken
    if (-not $token) {
        Write-Warn "Flow token missing. Configure auth.app.* or sign in for delegated token."
        return $null
    }
    $base = Get-FlowResource
    $url = if ($Path -match "^https?://") { $Path } else { $base + $Path }
    if ($url -notmatch "api-version=") {
        $join = if ($url -match "\\?") { "&" } else { "?" }
        $url = $url + $join + "api-version=2016-11-01"
    }
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


function Invoke-PpRequestRaw {
    param(
        [string]$Method,
        [string]$Path,
        [object]$Body,
        [string]$ApiVersion
    )
    $token = Get-PpToken
    if (-not $token) {
        Write-Warn "Power Platform token missing. Use pp login or configure auth.app.*"
        return $null
    }
    $base = $global:Config.pp.baseUrl.TrimEnd("/")
    $url = if ($Path -match "^https?://") { $Path } else { $base + $Path }
    $apiVer = if ($ApiVersion) { $ApiVersion } else { $global:Config.pp.apiVersion }
    if ($url -notmatch "api-version=") {
        $join = if ($url -match "\\?") { "&" } else { "?" }
        $url = $url + $join + "api-version=" + $apiVer
    }
    $hdr = @{ Authorization = "Bearer " + $token; accept = "application/json" }
    $params = @{ Method = $Method; Uri = $url; Headers = $hdr; ContentType = "application/json" }
    if ($Body -ne $null) { $params.Body = ($Body | ConvertTo-Json -Depth 10) }
    $respHeaders = $null
    try {
        $bodyOut = Invoke-RestMethod @params -ResponseHeadersVariable respHeaders
        return [pscustomobject]@{ Body = $bodyOut; Headers = $respHeaders }
    } catch {
        Write-Err $_.Exception.Message
        return $null
    }
}


function Resolve-OutputFilePath {
    param(
        [string]$Path,
        [string]$FileName,
        [string]$Extension
    )
    $safe = if ($FileName) { $FileName } else { "output" }
    $safe = $safe -replace '[\\/:*?"<>|]', '_'
    if ($Extension -and -not $safe.EndsWith($Extension)) { $safe += $Extension }
    if (-not $Path) { return (Join-Path (Get-Location) $safe) }
    if ($Path.EndsWith("\") -or $Path.EndsWith("/")) { return (Join-Path $Path $safe) }
    if ($Extension -and $Path.ToLowerInvariant().EndsWith($Extension.ToLowerInvariant())) { return $Path }
    if (Test-Path $Path -PathType Container) { return (Join-Path $Path $safe) }
    return $Path
}


function Handle-FlowCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: flow list|get|enable|disable|remove|export|run|owner|environment|recyclebinitem ..."
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "environment" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: flow environment list|get --name <env> OR --default"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest2
            if ($action -eq "list") {
                $resp = Invoke-FlowRequest -Method "GET" -Path "/providers/Microsoft.ProcessSimple/environments"
                if ($resp -and $resp.value) { $resp.value | ConvertTo-Json -Depth 8 }
            } elseif ($action -eq "get") {
                $name = Get-ArgValue $parsed2.Map "name"
                $isDefault = Parse-Bool (Get-ArgValue $parsed2.Map "default") $false
                if (-not $name -and -not $isDefault) {
                    Write-Warn "Usage: flow environment get --name <env> OR --default"
                    return
                }
                $envName = if ($isDefault) { "~default" } else { $name }
                $resp = Invoke-FlowRequest -Method "GET" -Path ("/providers/Microsoft.ProcessSimple/environments/" + $envName)
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
            } else {
                Write-Warn "Usage: flow environment list|get"
            }
        }
        "list" {
            $env = Get-ArgValue $parsed.Map "environmentName"
            if (-not $env) { $env = Get-ArgValue $parsed.Map "env" }
            if (-not $env) {
                Write-Warn "Usage: flow list --environmentName <env> [--sharingStatus personal|sharedWithMe|all] [--withSolutions] [--asAdmin]"
                return
            }
            $sharing = Get-ArgValue $parsed.Map "sharingStatus"
            $withSolutions = Parse-Bool (Get-ArgValue $parsed.Map "withSolutions") $false
            $asAdmin = Parse-Bool (Get-ArgValue $parsed.Map "asAdmin") $false
            $base = "/providers/Microsoft.ProcessSimple/" + ($(if ($asAdmin) { "scopes/admin/" } else { "" })) + "environments/" + $env + "/v2/flows"
            $query = @()
            if ($sharing -eq "personal") { $query += "`$filter=search('personal')" }
            if ($sharing -eq "sharedWithMe") { $query += "`$filter=search('team')" }
            if ($withSolutions) { $query += "include=includeSolutionCloudFlows" }
            $path = $base + (if ($query.Count -gt 0) { "?" + ($query -join "&") } else { "" })
            $resp = Invoke-FlowRequest -Method "GET" -Path $path
            if ($resp -and $resp.value) { $resp.value | ConvertTo-Json -Depth 8 }
        }
        "get" {
            $env = Get-ArgValue $parsed.Map "environmentName"
            if (-not $env) { $env = Get-ArgValue $parsed.Map "env" }
            $name = Get-ArgValue $parsed.Map "name"
            if (-not $name) { $name = $parsed.Positionals | Select-Object -First 1 }
            $asAdmin = Parse-Bool (Get-ArgValue $parsed.Map "asAdmin") $false
            if (-not $env -or -not $name) {
                Write-Warn "Usage: flow get --environmentName <env> --name <flowId> [--asAdmin]"
                return
            }
            $expand = "`$expand=swagger,properties.connectionreferences.apidefinition,properties.definitionsummary.operations.apioperation,operationDefinition,plan,properties.throttleData,properties.estimatedsuspensiondata,properties.licenseData"
            $path = "/providers/Microsoft.ProcessSimple/" + ($(if ($asAdmin) { "scopes/admin/" } else { "" })) + "environments/" + $env + "/flows/" + $name + "?" + $expand
            $resp = Invoke-FlowRequest -Method "GET" -Path $path
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "enable" {
            $env = Get-ArgValue $parsed.Map "environmentName"
            if (-not $env) { $env = Get-ArgValue $parsed.Map "env" }
            $name = Get-ArgValue $parsed.Map "name"
            if (-not $name) { $name = $parsed.Positionals | Select-Object -First 1 }
            $asAdmin = Parse-Bool (Get-ArgValue $parsed.Map "asAdmin") $false
            if (-not $env -or -not $name) {
                Write-Warn "Usage: flow enable --environmentName <env> --name <flowId> [--asAdmin]"
                return
            }
            $path = "/providers/Microsoft.ProcessSimple/" + ($(if ($asAdmin) { "scopes/admin/" } else { "" })) + "environments/" + $env + "/flows/" + $name + "/start"
            $resp = Invoke-FlowRequest -Method "POST" -Path $path -AllowNullResponse
            if ($resp -ne $null) { Write-Info "Flow enabled." }
        }
        "disable" {
            $env = Get-ArgValue $parsed.Map "environmentName"
            if (-not $env) { $env = Get-ArgValue $parsed.Map "env" }
            $name = Get-ArgValue $parsed.Map "name"
            if (-not $name) { $name = $parsed.Positionals | Select-Object -First 1 }
            $asAdmin = Parse-Bool (Get-ArgValue $parsed.Map "asAdmin") $false
            if (-not $env -or -not $name) {
                Write-Warn "Usage: flow disable --environmentName <env> --name <flowId> [--asAdmin]"
                return
            }
            $path = "/providers/Microsoft.ProcessSimple/" + ($(if ($asAdmin) { "scopes/admin/" } else { "" })) + "environments/" + $env + "/flows/" + $name + "/stop"
            $resp = Invoke-FlowRequest -Method "POST" -Path $path -AllowNullResponse
            if ($resp -ne $null) { Write-Info "Flow disabled." }
        }
        "remove" {
            $env = Get-ArgValue $parsed.Map "environmentName"
            if (-not $env) { $env = Get-ArgValue $parsed.Map "env" }
            $name = Get-ArgValue $parsed.Map "name"
            if (-not $name) { $name = $parsed.Positionals | Select-Object -First 1 }
            $asAdmin = Parse-Bool (Get-ArgValue $parsed.Map "asAdmin") $false
            if (-not $env -or -not $name) {
                Write-Warn "Usage: flow remove --environmentName <env> --name <flowId> [--asAdmin] [--force]"
                return
            }
            $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
            if (-not $force) {
                $confirm = Read-Host "Type DELETE to confirm"
                if ($confirm -ne "DELETE") { Write-Info "Canceled."; return }
            }
            $path = "/providers/Microsoft.ProcessSimple/" + ($(if ($asAdmin) { "scopes/admin/" } else { "" })) + "environments/" + $env + "/flows/" + $name
            $resp = Invoke-FlowRequest -Method "DELETE" -Path $path -AllowNullResponse
            if ($resp -ne $null) { Write-Info "Flow removed." }
        }
        "export" {
            $env = Get-ArgValue $parsed.Map "environmentName"
            if (-not $env) { $env = Get-ArgValue $parsed.Map "env" }
            $name = Get-ArgValue $parsed.Map "name"
            if (-not $name) { $name = $parsed.Positionals | Select-Object -First 1 }
            $format = Get-ArgValue $parsed.Map "format"
            $pathOut = Get-ArgValue $parsed.Map "path"
            if (-not $format) { $format = "zip" }
            if (-not $env -or -not $name) {
                Write-Warn "Usage: flow export --environmentName <env> --name <flowId> [--format json|zip] [--path <file|folder>]"
                return
            }

            if ($format.ToLowerInvariant() -eq "json") {
                $info = Invoke-FlowRequest -Method "GET" -Path ("/providers/Microsoft.ProcessSimple/environments/" + $env + "/flows/" + $name)
                $display = if ($info -and $info.properties -and $info.properties.displayName) { $info.properties.displayName } else { $name }
                $arm = Invoke-FlowRequest -Method "POST" -Path ("/providers/Microsoft.ProcessSimple/environments/" + $env + "/flows/" + $name + "/exportToARMTemplate")
                if (-not $arm) { return }
                $target = Resolve-OutputFilePath $pathOut $display ".json"
                $arm | ConvertTo-Json -Depth 20 | Set-Content -Path $target -Encoding ASCII
                Write-Info ("Saved: " + $target)
                return
            }

            $res = Invoke-PpRequestRaw -Method "POST" -Path ("/providers/Microsoft.BusinessAppPlatform/environments/" + $env + "/listPackageResources") -ApiVersion "2016-11-01" -Body @{
                baseResourceIds = @("/providers/Microsoft.Flow/flows/" + $name)
            }
            if (-not $res -or -not $res.Body) { return }
            $resources = $res.Body.resources
            if (-not $resources) { Write-Warn "Package resources missing."; return }
            foreach ($k in $resources.Keys) {
                if ($resources[$k].type -eq "Microsoft.Flow/flows") {
                    $resources[$k].suggestedCreationType = "Update"
                } else {
                    $resources[$k].suggestedCreationType = "Existing"
                }
            }
            $details = @{
                displayName       = Get-ArgValue $parsed.Map "packageDisplayName"
                description       = Get-ArgValue $parsed.Map "packageDescription"
                creator           = Get-ArgValue $parsed.Map "packageCreatedBy"
                sourceEnvironment = Get-ArgValue $parsed.Map "packageSourceEnvironment"
            }
            $export = Invoke-PpRequestRaw -Method "POST" -Path ("/providers/Microsoft.BusinessAppPlatform/environments/" + $env + "/exportPackage") -ApiVersion "2016-11-01" -Body @{
                includedResourceIds = @("/providers/Microsoft.Flow/flows/" + $name)
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
                    if ($status -eq "Succeeded") {
                        $packageLink = $resp.properties.packageLink.value
                        break
                    }
                } catch {
                    Write-Err $_.Exception.Message
                    return
                }
            }
            if (-not $packageLink) {
                Write-Warn "Package export did not complete."
                return
            }
            $fileName = Get-ArgValue $parsed.Map "packageDisplayName"
            if (-not $fileName) { $fileName = $name }
            $target = Resolve-OutputFilePath $pathOut $fileName ".zip"
            try {
                Invoke-WebRequest -Uri $packageLink -OutFile $target -Headers @{ "x-anonymous" = "true" } | Out-Null
                Write-Info ("Saved: " + $target)
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "run" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: flow run list|get|cancel|resubmit --environmentName <env> --flowName <id> ..."
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest2
            $env = Get-ArgValue $parsed2.Map "environmentName"
            if (-not $env) { $env = Get-ArgValue $parsed2.Map "env" }
            $flowId = Get-ArgValue $parsed2.Map "flowName"
            if (-not $flowId) { $flowId = Get-ArgValue $parsed2.Map "name" }
            if (-not $env -or -not $flowId) {
                Write-Warn "Usage: flow run <action> --environmentName <env> --flowName <id>"
                return
            }
            $asAdmin = Parse-Bool (Get-ArgValue $parsed2.Map "asAdmin") $false
            $base = "/providers/Microsoft.ProcessSimple/" + ($(if ($asAdmin) { "scopes/admin/" } else { "" })) + "environments/" + $env + "/flows/" + $flowId + "/runs"
            switch ($action) {
                "list" {
                    $filters = @()
                    $status = Get-ArgValue $parsed2.Map "status"
                    if ($status) { $filters += "status eq '" + $status + "'" }
                    $start = Get-ArgValue $parsed2.Map "triggerStartTime"
                    if ($start) { $filters += "startTime ge " + $start }
                    $end = Get-ArgValue $parsed2.Map "triggerEndTime"
                    if ($end) { $filters += "startTime lt " + $end }
                    $path = $base
                    if ($filters.Count -gt 0) { $path += "?`$filter=" + ($filters -join " and ") }
                    $resp = Invoke-FlowRequest -Method "GET" -Path $path
                    if ($resp -and $resp.value) { $resp.value | ConvertTo-Json -Depth 8 }
                }
                "get" {
                    $runId = $parsed2.Positionals | Select-Object -First 1
                    if (-not $runId) { $runId = Get-ArgValue $parsed2.Map "runId" }
                    if (-not $runId) {
                        Write-Warn "Usage: flow run get <runId> --environmentName <env> --flowName <id>"
                        return
                    }
                    $resp = Invoke-FlowRequest -Method "GET" -Path ($base + "/" + $runId)
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "cancel" {
                    $runId = $parsed2.Positionals | Select-Object -First 1
                    if (-not $runId) { $runId = Get-ArgValue $parsed2.Map "runId" }
                    if (-not $runId) {
                        Write-Warn "Usage: flow run cancel <runId> --environmentName <env> --flowName <id>"
                        return
                    }
                    $resp = Invoke-FlowRequest -Method "POST" -Path ($base + "/" + $runId + "/cancel") -AllowNullResponse
                    if ($resp -ne $null) { Write-Info "Run canceled." }
                }
                "resubmit" {
                    $runId = $parsed2.Positionals | Select-Object -First 1
                    if (-not $runId) { $runId = Get-ArgValue $parsed2.Map "runId" }
                    if (-not $runId) {
                        Write-Warn "Usage: flow run resubmit <runId> --environmentName <env> --flowName <id>"
                        return
                    }
                    $resp = Invoke-FlowRequest -Method "POST" -Path ($base + "/" + $runId + "/resubmit") -AllowNullResponse
                    if ($resp -ne $null) { Write-Info "Run resubmitted." }
                }
                default {
                    Write-Warn "Usage: flow run list|get|cancel|resubmit ..."
                }
            }
        }
        "owner" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: flow owner list|ensure|remove --environmentName <env> --flowName <id>"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest2
            $env = Get-ArgValue $parsed2.Map "environmentName"
            if (-not $env) { $env = Get-ArgValue $parsed2.Map "env" }
            $flowId = Get-ArgValue $parsed2.Map "flowName"
            if (-not $flowId) { $flowId = Get-ArgValue $parsed2.Map "name" }
            $asAdmin = Parse-Bool (Get-ArgValue $parsed2.Map "asAdmin") $false
            if (-not $env -or -not $flowId) {
                Write-Warn "Usage: flow owner <action> --environmentName <env> --flowName <id>"
                return
            }
            $base = "/providers/Microsoft.ProcessSimple/" + ($(if ($asAdmin) { "scopes/admin/" } else { "" })) + "environments/" + $env + "/flows/" + $flowId + "/permissions"
            switch ($action) {
                "list" {
                    $resp = Invoke-FlowRequest -Method "GET" -Path $base
                    if ($resp -and $resp.value) { $resp.value | ConvertTo-Json -Depth 8 }
                }
                "ensure" {
                    $role = Get-ArgValue $parsed2.Map "roleName"
                    if (-not $role) { $role = "CanEdit" }
                    $userId = Get-ArgValue $parsed2.Map "userId"
                    $userName = Get-ArgValue $parsed2.Map "userName"
                    $groupId = Get-ArgValue $parsed2.Map "groupId"
                    $groupName = Get-ArgValue $parsed2.Map "groupName"
                    $principalId = $null
                    $principalType = $null
                    if ($userId) { $principalId = $userId; $principalType = "User" }
                    elseif ($userName) { $u = Resolve-UserObject $userName; if ($u) { $principalId = $u.Id; $principalType = "User" } }
                    elseif ($groupId) { $principalId = $groupId; $principalType = "Group" }
                    elseif ($groupName) { $g = Resolve-GroupObject $groupName; if ($g) { $principalId = $g.Id; $principalType = "Group" } }
                    if (-not $principalId) {
                        Write-Warn "Usage: flow owner ensure --userId <id>|--userName <upn>|--groupId <id>|--groupName <name> --roleName CanView|CanEdit"
                        return
                    }
                    $body = @{
                        put = @(
                            @{
                                properties = @{
                                    principal = @{ id = $principalId; type = $principalType }
                                    roleName  = $role
                                }
                            }
                        )
                    }
                    $resp = Invoke-FlowRequest -Method "POST" -Path ($base.Replace("/permissions","/modifyPermissions")) -Body $body -AllowNullResponse
                    if ($resp -ne $null) { Write-Info "Owner updated." }
                }
                "remove" {
                    $principalId = Get-ArgValue $parsed2.Map "principalId"
                    if (-not $principalId) { $principalId = Get-ArgValue $parsed2.Map "id" }
                    if (-not $principalId) {
                        Write-Warn "Usage: flow owner remove --principalId <id>"
                        return
                    }
                    $body = @{ delete = @(@{ id = $principalId }) }
                    $resp = Invoke-FlowRequest -Method "POST" -Path ($base.Replace("/permissions","/modifyPermissions")) -Body $body -AllowNullResponse
                    if ($resp -ne $null) { Write-Info "Owner removed." }
                }
                default {
                    Write-Warn "Usage: flow owner list|ensure|remove ..."
                }
            }
        }
        "recyclebinitem" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: flow recyclebinitem list|restore --environmentName <env> [--flowName <id>]"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest2
            $env = Get-ArgValue $parsed2.Map "environmentName"
            if (-not $env) { $env = Get-ArgValue $parsed2.Map "env" }
            if (-not $env) {
                Write-Warn "Usage: flow recyclebinitem <action> --environmentName <env>"
                return
            }
            switch ($action) {
                "list" {
                    $path = "/providers/Microsoft.ProcessSimple/scopes/admin/environments/" + $env + "/v2/flows?include=softDeletedFlows"
                    $resp = Invoke-FlowRequest -Method "GET" -Path $path
                    if ($resp -and $resp.value) {
                        $deleted = @($resp.value | Where-Object { $_.properties.state -eq "Deleted" })
                        $deleted | ConvertTo-Json -Depth 8
                    }
                }
                "restore" {
                    $flowId = Get-ArgValue $parsed2.Map "flowName"
                    if (-not $flowId) { $flowId = Get-ArgValue $parsed2.Map "name" }
                    if (-not $flowId) {
                        Write-Warn "Usage: flow recyclebinitem restore --environmentName <env> --flowName <id>"
                        return
                    }
                    $path = "/providers/Microsoft.ProcessSimple/scopes/admin/environments/" + $env + "/flows/" + $flowId + "/restore"
                    $resp = Invoke-FlowRequest -Method "POST" -Path $path -AllowNullResponse
                    if ($resp -ne $null) { Write-Info "Flow restored." }
                }
                default {
                    Write-Warn "Usage: flow recyclebinitem list|restore ..."
                }
            }
        }
        "connector" {
            Write-Warn "flow connector is routed to pa connector. Use: pa connector list|export"
        }
        default {
            Write-Warn "Usage: flow list|get|enable|disable|remove|export|run|owner|environment|recyclebinitem ..."
        }
    }
}
