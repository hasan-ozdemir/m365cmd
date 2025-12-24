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
                    Write-Warn "pa app export is not implemented yet."
                }
                "permission" {
                    Write-Warn "pa app permission commands are not implemented yet."
                }
                "owner" {
                    Write-Warn "pa app owner commands are not implemented yet."
                }
                "consent" {
                    Write-Warn "pa app consent commands are not implemented yet."
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
                Write-Warn "pa connector export is not implemented yet."
            } else {
                Write-Warn "Usage: pa connector list|export --environmentName <env>"
            }
        }
        default {
            Write-Warn "Usage: pa app|connector|environment <args...>"
        }
    }
}
