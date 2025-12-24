# Handler: Spo
# Purpose: Spo command handlers.
function Require-SpoConnection {
    if (-not (Ensure-ModuleLoaded "Microsoft.Online.SharePoint.PowerShell" -UseWindowsPowerShell)) { return $false }
    if ($global:SpoConnected) { return $true }
    try {
        Get-SPOTenant -ErrorAction Stop | Out-Null
        $global:SpoConnected = $true
        return $true
    } catch {}
    Write-Warn "Not connected to SharePoint Online. Use: spo connect"
    return $false
}


function Handle-SpoCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: spo connect|disconnect|status|site|tenant|onedrive|cmd|cmdlets|rename"
        return
    }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "connect" {
            if (-not (Ensure-ModuleLoaded "Microsoft.Online.SharePoint.PowerShell" -UseWindowsPowerShell)) { return }
            $url = Get-ArgValue $parsed.Map "url"
            if (-not $url) {
                $prefix = Get-ArgValue $parsed.Map "prefix"
                if (-not $prefix) { $prefix = $global:Config.tenant.defaultPrefix }
                if ($prefix) {
                    $url = "https://" + $prefix + "-admin.sharepoint.com"
                }
            }
            if (-not $url) {
                Write-Warn "Usage: spo connect [--url https://tenant-admin.sharepoint.com] OR --prefix <tenantPrefix>"
                return
            }
            $authUrl = Get-ArgValue $parsed.Map "authUrl"
            try {
                if ($authUrl) {
                    Connect-SPOService -Url $url -AuthenticationUrl $authUrl | Out-Null
                } else {
                    Connect-SPOService -Url $url | Out-Null
                }
                $global:SpoConnected = $true
                $global:SpoAdminUrl = $url
                Write-Info ("Connected to SharePoint Online: " + $url)
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "disconnect" {
            if (-not (Ensure-ModuleLoaded "Microsoft.Online.SharePoint.PowerShell" -UseWindowsPowerShell)) { return }
            try {
                if (Get-Command Disconnect-SPOService -ErrorAction SilentlyContinue) {
                    Disconnect-SPOService | Out-Null
                }
                $global:SpoConnected = $false
                Write-Info "Disconnected from SharePoint Online."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "status" {
            $status = if ($global:SpoConnected) { "connected" } else { "not connected" }
            $url = if ($global:SpoAdminUrl) { $global:SpoAdminUrl } else { "<not set>" }
            Write-Host ("SharePoint Online: " + $status)
            Write-Host ("Admin URL      : " + $url)
        }
        "cmdlets" {
            if (-not (Ensure-ModuleLoaded "Microsoft.Online.SharePoint.PowerShell" -UseWindowsPowerShell)) { return }
            $filter = Get-ArgValue $parsed.Map "filter"
            $cmds = Get-Command -Module Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue
            if ($filter) {
                $cmds = $cmds | Where-Object { $_.Name -like ("*" + $filter + "*") }
            }
            $cmds | Sort-Object Name | Select-Object -ExpandProperty Name | Format-Wide -Column 3
        }
        "cmd" {
            if (-not (Ensure-ModuleLoaded "Microsoft.Online.SharePoint.PowerShell" -UseWindowsPowerShell)) { return }
            if (-not $global:SpoConnected) {
                Write-Warn "Not connected to SharePoint Online. Use: spo connect"
                return
            }
            $cmdlet = $parsed.Positionals | Select-Object -First 1
            if (-not $cmdlet) {
                Write-Warn "Usage: spo cmd <cmdlet> [--params key=value[,key=value]] [--json <payload>] [--bodyFile <file>] [--set key=value]"
                return
            }
            $paramObj = Resolve-CmdletParams $parsed
            try {
                if ($paramObj.Keys.Count -gt 0) {
                    & $cmdlet @paramObj
                } else {
                    & $cmdlet
                }
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "site" {
            if (-not (Require-SpoConnection)) { return }
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action) {
                Write-Warn "Usage: spo site list|get|create|update|delete"
                return
            }
            switch ($action) {
                "list" {
                    $includePersonal = Parse-Bool (Get-ArgValue $parsed.Map "includePersonal") $false
                    try {
                        $params = @{ Limit = "All" }
                        if ($includePersonal) { $params.IncludePersonalSite = $true }
                        Get-SPOSite @params | Select-Object Url, Title, Template, StorageQuota, StorageUsageCurrent | Format-Table -AutoSize
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "get" {
                    if (-not $identity) {
                        Write-Warn "Usage: spo site get <url>"
                        return
                    }
                    try {
                        Get-SPOSite -Identity $identity | Format-List *
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "create" {
                    $url = Get-ArgValue $parsed.Map "url"
                    $owner = Get-ArgValue $parsed.Map "owner"
                    $title = Get-ArgValue $parsed.Map "title"
                    $quota = Get-ArgValue $parsed.Map "storage"
                    $template = Get-ArgValue $parsed.Map "template"
                    if (-not $url -or -not $owner) {
                        Write-Warn "Usage: spo site create --url <siteUrl> --owner <upn> [--title <title>] [--storage 1000] [--template STS#0]"
                        return
                    }
                    if (-not $quota) { $quota = 1000 }
                    try {
                        $params = @{
                            Url          = $url
                            Owner        = $owner
                            StorageQuota = [int64]$quota
                        }
                        if ($title) { $params.Title = $title }
                        if ($template) { $params.Template = $template }
                        New-SPOSite @params | Out-Null
                        Write-Info "Site collection creation started."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "update" {
                    if (-not $identity) {
                        Write-Warn "Usage: spo site update <url> --set key=value[,key=value]"
                        return
                    }
                    $setRaw = Get-ArgValue $parsed.Map "set"
                    if (-not $setRaw) {
                        Write-Warn "Usage: spo site update <url> --set key=value[,key=value]"
                        return
                    }
                    $body = Parse-Value $setRaw
                    if ($body -is [string]) { $body = Parse-KvPairs $setRaw }
                    if ($body.Keys.Count -eq 0) {
                        Write-Warn "No properties to update."
                        return
                    }
                    try {
                        $params = @{ Identity = $identity }
                        foreach ($k in $body.Keys) { $params[$k] = $body[$k] }
                        Set-SPOSite @params | Out-Null
                        Write-Info "Site updated."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "delete" {
                    if (-not $identity) {
                        Write-Warn "Usage: spo site delete <url> [--force]"
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
                    try {
                        Remove-SPOSite -Identity $identity -Confirm:$false | Out-Null
                        Write-Info "Site collection removed (sent to recycle bin)."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                default {
                    Write-Warn "Usage: spo site list|get|create|update|delete"
                }
            }
        }
        "tenant" {
            if (-not (Require-SpoConnection)) { return }
            $action = $parsed.Positionals | Select-Object -First 1
            if (-not $action) {
                Write-Warn "Usage: spo tenant get|update"
                return
            }
            switch ($action) {
                "get" {
                    try {
                        Get-SPOTenant | Format-List *
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "update" {
                    $setRaw = Get-ArgValue $parsed.Map "set"
                    if (-not $setRaw) {
                        Write-Warn "Usage: spo tenant update --set key=value[,key=value]"
                        return
                    }
                    $body = Parse-Value $setRaw
                    if ($body -is [string]) { $body = Parse-KvPairs $setRaw }
                    if ($body.Keys.Count -eq 0) {
                        Write-Warn "No properties to update."
                        return
                    }
                    try {
                        $params = @{}
                        foreach ($k in $body.Keys) { $params[$k] = $body[$k] }
                        Set-SPOTenant @params | Out-Null
                        Write-Info "Tenant settings updated."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                default {
                    Write-Warn "Usage: spo tenant get|update"
                }
            }
        }
        "onedrive" {
            if (-not (Require-SpoConnection)) { return }
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action) {
                Write-Warn "Usage: spo onedrive list|get|update"
                return
            }
            switch ($action) {
                "list" {
                    try {
                        $params = @{ IncludePersonalSite = $true; Limit = "All" }
                        $filter = Get-ArgValue $parsed.Map "filter"
                        if ($filter) { $params.Filter = $filter }
                        Get-SPOSite @params | Select-Object Url, Owner, StorageQuota, StorageUsageCurrent, LastContentModifiedDate | Format-Table -AutoSize
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "get" {
                    if (-not $identity) {
                        Write-Warn "Usage: spo onedrive get <url>"
                        return
                    }
                    try {
                        Get-SPOSite -Identity $identity | Format-List *
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "update" {
                    if (-not $identity) {
                        Write-Warn "Usage: spo onedrive update <url> --set key=value[,key=value]"
                        return
                    }
                    $setRaw = Get-ArgValue $parsed.Map "set"
                    if (-not $setRaw) {
                        Write-Warn "Usage: spo onedrive update <url> --set key=value[,key=value]"
                        return
                    }
                    $body = Parse-Value $setRaw
                    if ($body -is [string]) { $body = Parse-KvPairs $setRaw }
                    if ($body.Keys.Count -eq 0) {
                        Write-Warn "No properties to update."
                        return
                    }
                    try {
                        $params = @{ Identity = $identity }
                        foreach ($k in $body.Keys) { $params[$k] = $body[$k] }
                        Set-SPOSite @params | Out-Null
                        Write-Info "OneDrive site updated."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                default {
                    Write-Warn "Usage: spo onedrive list|get|update"
                }
            }
        }
        "rename" {
            if (-not (Ensure-ModuleLoaded "Microsoft.Online.SharePoint.PowerShell" -UseWindowsPowerShell)) { return }
            $action = $parsed.Positionals | Select-Object -First 1
            $current = Get-ArgValue $parsed.Map "current"
            if (-not $current) { $current = $global:Config.tenant.defaultPrefix }
            $newPrefix = Get-ArgValue $parsed.Map "new"
            $adminUpn = Get-ArgValue $parsed.Map "adminUpn"
            if (-not $adminUpn) { $adminUpn = $global:Config.admin.defaultUpn }
            $scheduleRaw = Get-ArgValue $parsed.Map "schedule"
            $authUrl = Get-ArgValue $parsed.Map "authUrl"
            if (-not $authUrl) { $authUrl = "https://login.microsoftonline.com/organizations" }
            $credMsg = Get-ArgValue $parsed.Map "credentialMessage"
            if (-not $credMsg) { $credMsg = "SharePoint domain rename - enter password for Global Admin" }
            $useInteractive = Parse-Bool (Get-ArgValue $parsed.Map "interactive") $true
            $forceCredential = Parse-Bool (Get-ArgValue $parsed.Map "forceCredential") $false
            $url = Get-ArgValue $parsed.Map "url"
            if (-not $url -and $current) { $url = "https://" + $current + "-admin.sharepoint.com" }

            if (-not $action) {
                Write-Warn "Usage: spo rename status|start ..."
                return
            }
            if (-not $url) {
                Write-Warn "Usage: spo rename <action> --current <tenantPrefix> OR --url <adminUrl>"
                return
            }

            try {
                Connect-SpoTenantAdaptive -Url $url -AdminUpn $adminUpn -CredentialMessage $credMsg `
                    -UseInteractive $useInteractive -ForceCredential $forceCredential -AuthenticationUrl $authUrl | Out-Null
                $global:SpoConnected = $true
                $global:SpoAdminUrl = $url
                Get-SPOTenant | Out-Null
            } catch {
                $detail = $_.Exception.Message
                if ($_.Exception.InnerException) { $detail = $_.Exception.InnerException.Message }
                if ($detail -match "AADSTS50076") {
                    Write-Err ("Could not connect due to MFA requirement (AADSTS50076). Use MFA-friendly login or update the SPO module. Details: " + $detail)
                } else {
                    Write-Err ("Could not connect to SharePoint admin service. Details: " + $detail)
                }
                return
            }

            if ($action -eq "status" -or $action -eq "check") {
                Show-SpoRenameStatusDetails | Out-Null
                return
            }

            if ($action -eq "start" -or $action -eq "submit") {
                if (-not $newPrefix) {
                    Write-Warn "Usage: spo rename start --new <prefix> [--schedule <datetime>]"
                    return
                }
                $schedule = if ($scheduleRaw) { [datetime]::Parse($scheduleRaw) } else { (Get-Date).AddHours(25) }
                try {
                    Validate-SpoRenameSchedule $schedule
                } catch {
                    Write-Err $_.Exception.Message
                    return
                }
                if (-not (Get-Command Start-SPOTenantRename -ErrorAction SilentlyContinue)) {
                    Write-Err "Start-SPOTenantRename not found. Update the SPO module."
                    return
                }
                try {
                    Start-SPOTenantRename -DomainName $newPrefix -ScheduledDateTime $schedule -Confirm:$false -ErrorAction Stop
                    Write-Host "Submitted. Monitor with: spo rename status" -ForegroundColor Green
                } catch {
                    $msg = $_.Exception.Message
                    if ($_.Exception.InnerException) { $msg = $_.Exception.InnerException.Message }
                    if ($msg -match "Error Code:\\s*-?783" -or $msg -match "\\b783\\b") {
                        Write-Err ("Tenant rename limit reached (error 783). A SharePoint tenant can be renamed only once. Details: " + $msg)
                    } else {
                        Write-Err ("Start-SPOTenantRename failed. Details: " + $msg)
                    }
                }
                return
            }

            Write-Warn "Usage: spo rename status|start ..."
        }
        default {
            Write-Warn "Usage: spo connect|disconnect|status|site|tenant|onedrive|cmd|cmdlets|rename"
        }
    }
}

function Validate-SpoRenameSchedule {
    param([datetime]$Schedule)
    $now = Get-Date
    if ($Schedule -lt $now.AddHours(24)) {
        throw "Schedule must be at least 24 hours in the future. Current: $now"
    }
    if ($Schedule -gt $now.AddDays(30)) {
        throw "Schedule must be within 30 days. Current: $now"
    }
}

function Connect-SpoTenantAdaptive {
    param(
        [string]$Url,
        [string]$AdminUpn,
        [string]$CredentialMessage,
        [bool]$UseInteractive,
        [bool]$ForceCredential,
        [string]$AuthenticationUrl
    )
    $cmd = Get-Command Connect-SPOService -ErrorAction Stop
    $supportsInteractive = $cmd.Parameters.ContainsKey("Interactive")
    $supportsUseWebLogin = $cmd.Parameters.ContainsKey("UseWebLogin")
    $supportsModernAuth  = $cmd.Parameters.ContainsKey("ModernAuth")
    $supportsAuthUrl     = $cmd.Parameters.ContainsKey("AuthenticationUrl")
    $lastError = $null

    if ($UseInteractive -and $supportsInteractive) {
        try {
            Write-Host "Connecting (interactive)..."
            Connect-SPOService -Url $Url -Interactive -ErrorAction Stop
            return $true
        } catch {
            $lastError = $_
            Write-Warn ("Interactive login failed: " + $_.Exception.Message)
        }
    }

    try {
        Write-Host "Connecting (MFA-friendly login)..."
        Connect-SPOService -Url $Url -ErrorAction Stop
        return $true
    } catch {
        $lastError = $_
        Write-Warn ("MFA-friendly login failed: " + $_.Exception.Message)
    }

    if ($supportsUseWebLogin) {
        try {
            Write-Host "Connecting (web login)..."
            Connect-SPOService -Url $Url -UseWebLogin -ErrorAction Stop
            return $true
        } catch {
            $lastError = $_
            Write-Warn ("Web login failed: " + $_.Exception.Message)
        }
    }

    if ($ForceCredential) {
        $creds = Get-Credential -UserName $AdminUpn -Message $CredentialMessage
        if ($supportsModernAuth -and $supportsAuthUrl) {
            try {
                Write-Host "Connecting (credential + ModernAuth)..."
                Connect-SPOService -Url $Url -Credential $creds -ModernAuth $true -AuthenticationUrl $AuthenticationUrl -ErrorAction Stop
                return $true
            } catch {
                $lastError = $_
                Write-Warn ("Credential + ModernAuth failed: " + $_.Exception.Message)
            }
        }
        Write-Host "Connecting (credential prompt)..."
        Connect-SPOService -Url $Url -Credential $creds -ErrorAction Stop
        return $true
    }

    if ($lastError) { throw $lastError }
    throw "All connection attempts failed."
}

function Get-SpoRenameStatusDetails {
    if (-not (Get-Command Get-SPOTenantRenameStatus -ErrorAction SilentlyContinue)) {
        Write-Warn "Get-SPOTenantRenameStatus is not available. Update the SPO module to query rename status."
        return $null
    }
    try {
        return Get-SPOTenantRenameStatus -ErrorAction Stop
    } catch {
        Write-Warn ("Get-SPOTenantRenameStatus failed: " + $_.Exception.Message)
        return $null
    }
}

function Show-SpoRenameStatusDetails {
    $status = Get-SpoRenameStatusDetails
    if (-not $status) {
        Write-Host "No rename status returned (none scheduled or not supported by module)."
        return $false
    }
    Write-Host "Rename status details:" -ForegroundColor Cyan
    $status | Format-List * | Out-String | Write-Host
    $state = $status.State
    if (-not $state) { $state = $status.Status }
    if (-not $state) { $state = $status.RenameState }
    if (-not $state) { $state = $status.Phase }
    if ($state) {
        $s = $state.ToString()
        switch -Regex ($s) {
            "Scheduled|InProgress|Processing|Running" { Write-Host "Interpretation: pending or in progress." -ForegroundColor Yellow; break }
            "Completed|Succeeded|Success" { Write-Host "Interpretation: rename already completed." -ForegroundColor Green; break }
            "Failed|Error" { Write-Host "Interpretation: last rename failed." -ForegroundColor Red; break }
            "NotStarted|None|Unknown|NotScheduled" { Write-Host "Interpretation: no rename scheduled." -ForegroundColor Yellow; break }
            default { Write-Host ("Interpretation: state '" + $s + "' (unrecognized). Review details above.") -ForegroundColor Yellow }
        }
    } else {
        Write-Host "Interpretation: state field not available; review details above." -ForegroundColor Yellow
    }
    return $true
}

