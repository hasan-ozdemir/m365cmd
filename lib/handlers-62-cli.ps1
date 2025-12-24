# Handler: CLI
# Purpose: CLI-like helpers and config.
function Get-ConfigFlatMap {
    param([object]$Obj, [string]$Prefix = "")
    $map = @{}
    if ($null -eq $Obj) { return $map }
    foreach ($p in $Obj.PSObject.Properties) {
        $key = if ($Prefix) { $Prefix + "." + $p.Name } else { $p.Name }
        $val = $p.Value
        if ($val -and $val.PSObject.Properties.Count -gt 0 -and -not ($val -is [string])) {
            foreach ($kv in (Get-ConfigFlatMap $val $key).GetEnumerator()) {
                $map[$kv.Key] = $kv.Value
            }
        } else {
            $map[$key] = $val
        }
    }
    return $map
}


function Handle-CliCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: cli config|app|completion|consent|doctor|issue <args...>"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "config" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: cli config list|get|set|reset"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest2
            switch ($action) {
                "list" {
                    $flat = Get-ConfigFlatMap $global:Config
                    $flat.GetEnumerator() | Sort-Object Name | ForEach-Object {
                        [pscustomobject]@{ Key = $_.Key; Value = $_.Value }
                    } | Format-Table -AutoSize
                }
                "get" {
                    $key = Get-ArgValue $parsed2.Map "key"
                    if (-not $key) { $key = $parsed2.Positionals | Select-Object -First 1 }
                    if (-not $key) {
                        Write-Warn "Usage: cli config get --key <path>"
                        return
                    }
                    $val = Get-ConfigValue $key
                    if ($val -is [string]) { Write-Host $val } else { Write-Host ($val | ConvertTo-Json -Depth 8) }
                }
                "set" {
                    $key = Get-ArgValue $parsed2.Map "key"
                    $valRaw = Get-ArgValue $parsed2.Map "value"
                    if (-not $key) { $key = $parsed2.Positionals | Select-Object -First 1 }
                    if ($null -eq $valRaw -and $parsed2.Positionals.Count -ge 2) { $valRaw = $parsed2.Positionals[1] }
                    if (-not $key -or $null -eq $valRaw) {
                        Write-Warn "Usage: cli config set --key <path> --value <value>"
                        return
                    }
                    $val = Parse-Value $valRaw
                    Set-ConfigValue $key $val
                    Write-Info "Config updated."
                }
                "reset" {
                    $cfg = Get-DefaultConfig
                    Save-Config $cfg
                    $global:Config = Normalize-Config $cfg
                    Write-Info "Config reset to defaults."
                }
                default {
                    Write-Warn "Usage: cli config list|get|set|reset"
                }
            }
        }
        "app" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: cli app add|reconsent"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest2
            switch ($action) {
                "add" {
                    if (-not (Require-GraphConnection)) { return }
                    $name = Get-ArgValue $parsed2.Map "name"
                    if (-not $name) { $name = "CLI for M365" }
                    $scopesRaw = Get-ArgValue $parsed2.Map "scopes"
                    if (-not $scopesRaw) { $scopesRaw = "minimal" }
                    $save = Parse-Bool (Get-ArgValue $parsed2.Map "saveToConfig") $false
                    $scopes = @()
                    if ($scopesRaw -eq "all") {
                        $scopes = @($global:Config.auth.scopes)
                    } elseif ($scopesRaw -eq "minimal") {
                        $scopes = @("User.Read")
                    } else {
                        $scopes = Parse-CommaList $scopesRaw
                    }
                    $graphSp = Get-GraphServicePrincipal
                    $access = @()
                    foreach ($s in $scopes) {
                        $perm = Resolve-GraphPermission $graphSp $s "delegated"
                        if ($perm) { $access += @{ Id = $perm.Id; Type = $perm.Type } }
                    }
                    $reqAccess = @()
                    if ($access.Count -gt 0) {
                        $reqAccess += @{
                            ResourceAppId  = "00000003-0000-0000-c000-000000000000"
                            ResourceAccess = $access
                        }
                    }
                    $body = @{
                        DisplayName    = $name
                        SignInAudience = "AzureADMyOrg"
                        PublicClient   = @{ RedirectUris = @("http://localhost","https://localhost","https://login.microsoftonline.com/common/oauth2/nativeclient") }
                    }
                    if ($reqAccess.Count -gt 0) {
                        $body.RequiredResourceAccess = $reqAccess
                    }
                    try {
                        $app = New-MgApplication -BodyParameter $body
                        Write-Host ("AppId    : " + $app.AppId)
                        Write-Host ("ObjectId : " + $app.Id)
                        if ($save) {
                            Set-ConfigValue "auth.app.clientId" $app.AppId
                            $ctx = Get-MgContextSafe
                            if ($ctx -and $ctx.TenantId) { Set-ConfigValue "auth.app.tenantId" $ctx.TenantId }
                            Write-Info "Saved app id to config."
                        }
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "reconsent" {
                    $appId = Get-ArgValue $parsed2.Map "appId"
                    if (-not $appId) { $appId = $global:Config.auth.app.clientId }
                    if (-not $appId) {
                        Write-Warn "Usage: cli app reconsent --appId <appId>"
                        return
                    }
                    $tenant = $global:Config.tenant.tenantId
                    if (-not $tenant) { $tenant = $global:Config.tenant.defaultDomain }
                    $url = "https://login.microsoftonline.com/" + $tenant + "/adminconsent?client_id=" + $appId
                    Write-Host ("Admin consent URL: " + $url)
                    if ($parsed2.Map.ContainsKey("open")) {
                        try { Start-Process $url | Out-Null } catch {}
                    }
                }
                default {
                    Write-Warn "Usage: cli app add|reconsent"
                }
            }
        }
        "completion" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: cli completion pwsh|sh setup|update"
                return
            }
            $shell = $rest[0].ToLowerInvariant()
            $action = if ($rest.Count -gt 1) { $rest[1].ToLowerInvariant() } else { "" }
            if ($action -ne "setup" -and $action -ne "update") {
                Write-Warn "Usage: cli completion pwsh|sh setup|update"
                return
            }
            $dir = Join-Path $Paths.Data "completion"
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            if ($shell -eq "pwsh") {
                $target = Join-Path $dir "m365cmd-completion.ps1"
                "Register-ArgumentCompleter -Native -CommandName m365cmd -ScriptBlock { param(`$wordToComplete,`$commandAst,`$cursorPos) return @() }" | Set-Content -Path $target -Encoding ASCII
                Write-Info ("PowerShell completion script written: " + $target)
            } elseif ($shell -eq "sh") {
                $target = Join-Path $dir "m365cmd-completion.sh"
                "complete -W \"\" m365cmd" | Set-Content -Path $target -Encoding ASCII
                Write-Info ("Shell completion script written: " + $target)
            } else {
                Write-Warn "Usage: cli completion pwsh|sh setup|update"
            }
        }
        "consent" {
            $tenant = $global:Config.tenant.tenantId
            if (-not $tenant) { $tenant = $global:Config.tenant.defaultDomain }
            $appId = Get-ArgValue $parsed.Map "appId"
            if (-not $appId) { $appId = $global:Config.auth.app.clientId }
            if (-not $appId) {
                Write-Warn "Usage: cli consent --appId <appId>"
                return
            }
            $url = "https://login.microsoftonline.com/" + $tenant + "/adminconsent?client_id=" + $appId
            Write-Host ("Admin consent URL: " + $url)
        }
        "doctor" {
            $checks = @()
            $checks += [pscustomobject]@{ Check = "Graph module"; Status = $(if (Test-ModuleAvailable "Microsoft.Graph") { "ok" } else { "missing" }) }
            $checks += [pscustomobject]@{ Check = "MSAL module"; Status = $(if (Test-ModuleAvailable "MSAL.PS") { "ok" } else { "missing" }) }
            $checks += [pscustomobject]@{ Check = "Config file"; Status = $(if (Test-Path $Paths.Config) { "ok" } else { "missing" }) }
            $ctx = Get-MgContextSafe
            $checks += [pscustomobject]@{ Check = "Graph login"; Status = $(if ($ctx) { "connected" } else { "not connected" }) }
            $checks | Format-Table -AutoSize
        }
        "issue" {
            $url = "https://github.com/hasan-ozdemir/m365cmd/issues"
            Write-Host ("Issues: " + $url)
            if ($parsed.Map.ContainsKey("open")) {
                try { Start-Process $url | Out-Null } catch {}
            }
        }
        default {
            Write-Warn "Usage: cli config|app|completion|consent|doctor|issue <args...>"
        }
    }
}
