# Handler: Addins
# Purpose: Addins command handlers.
function Ensure-OrgAddinModule {
    if (-not $IsWindows) {
        Write-Warn "Office add-in centralized deployment module is supported only on Windows."
        return $false
    }
    return (Ensure-ModuleLoaded "O365CentralizedAddInDeployment" -UseWindowsPowerShell)
}


function Require-OrgAddinConnection {
    if (-not (Ensure-OrgAddinModule)) { return $false }
    if ($global:OrgAddinConnected) { return $true }
    Write-Warn "Not connected to the Office add-ins service. Use: addin org connect"
    return $false
}


function Get-ManifestBytes {
    param([string]$Path)
    if (-not $Path) { return $null }
    if (-not (Test-Path $Path)) {
        Write-Warn "Manifest file not found."
        return $null
    }
    return [System.IO.File]::ReadAllBytes($Path)
}


function Handle-AddinCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: addin exo|org|onsend <...>"
        return
    }
    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    switch ($sub) {
        "exo" { Handle-AddinExoCommand $rest }
        "outlook" { Handle-AddinExoCommand $rest }
        "mail" { Handle-AddinExoCommand $rest }
        "org" { Handle-AddinOrgCommand $rest }
        "office" { Handle-AddinOrgCommand $rest }
        "central" { Handle-AddinOrgCommand $rest }
        "onsend" { Handle-AddinOnSendCommand $rest }
        default {
            Write-Warn "Usage: addin exo|org|onsend <...>"
        }
    }
}


function Handle-AddinExoCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: addin exo list|get|install|update|remove|enable|disable|refresh [--org] [--mailbox <upn>]"
        return
    }
    if (-not (Require-ExoConnection)) { return }

    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $identity = $parsed.Positionals | Select-Object -First 1

    $org = $parsed.Map.ContainsKey("org") -or $parsed.Map.ContainsKey("organization")
    $mailbox = Get-ArgValue $parsed.Map "mailbox"
    $manifestPath = Get-ArgValue $parsed.Map "manifestPath"
    if (-not $manifestPath) { $manifestPath = Get-ArgValue $parsed.Map "path" }
    if (-not $manifestPath) { $manifestPath = Get-ArgValue $parsed.Map "file" }
    $manifestUrl = Get-ArgValue $parsed.Map "manifestUrl"
    if (-not $manifestUrl) { $manifestUrl = Get-ArgValue $parsed.Map "url" }
    $providedTo = Get-ArgValue $parsed.Map "providedTo"
    $userListRaw = Get-ArgValue $parsed.Map "userList"
    if (-not $userListRaw) { $userListRaw = Get-ArgValue $parsed.Map "users" }
    $defaultState = Get-ArgValue $parsed.Map "defaultState"
    if (-not $defaultState) { $defaultState = Get-ArgValue $parsed.Map "defaultStateForUser" }
    $enabledRaw = Get-ArgValue $parsed.Map "enabled"

    switch ($action) {
        "list" {
            try {
                $items = $null
                if ($parsed.Map.ContainsKey("all")) {
                    $items = Get-App
                } elseif ($mailbox) {
                    $items = Get-App -Mailbox $mailbox
                } else {
                    $items = Get-App -OrganizationApp
                }
                if ($items) {
                    $items | Select-Object DisplayName, AppId, Enabled, DefaultStateForUser, ProvidedTo | Format-Table -AutoSize
                }
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "get" {
            if (-not $identity) {
                Write-Warn "Usage: addin exo get <appId> [--org] [--mailbox <upn>]"
                return
            }
            try {
                if ($mailbox) {
                    Get-App -Mailbox $mailbox -Identity $identity | Format-List *
                } elseif ($org) {
                    Get-App -OrganizationApp -Identity $identity | Format-List *
                } else {
                    Get-App -Identity $identity | Format-List *
                }
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "create" { $action = "install" }
        "install" {
            if (-not $manifestPath -and -not $manifestUrl) {
                Write-Warn "Usage: addin exo install --manifestPath <file.xml> OR --manifestUrl <url> [--org] [--mailbox <upn>] [--providedTo AllUsers|SpecificUsers] [--userList a,b] [--defaultState Enabled|Disabled|AlwaysEnabled]"
                return
            }
            $params = @{}
            if ($org) { $params.OrganizationApp = $true }
            if ($mailbox) { $params.Mailbox = $mailbox }
            if ($manifestPath) {
                $bytes = Get-ManifestBytes $manifestPath
                if (-not $bytes) { return }
                $params.FileData = $bytes
            } elseif ($manifestUrl) {
                $params.ManifestURL = $manifestUrl
            }
            if ($providedTo) { $params.ProvidedTo = $providedTo }
            if ($userListRaw) { $params.UserList = Parse-CommaList $userListRaw }
            if ($defaultState) { $params.DefaultStateForUser = $defaultState }
            try {
                $app = New-App @params
                if ($app) {
                    Write-Info ("Installed add-in: " + $app.DisplayName)
                }
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "update" { $action = "refresh" }
        "refresh" {
            if (-not $identity) {
                Write-Warn "Usage: addin exo refresh <appId> [--org] [--mailbox <upn>] [--manifestPath <file.xml>|--manifestUrl <url>] [--providedTo ...] [--userList a,b] [--defaultState Enabled|Disabled|AlwaysEnabled] [--enabled true|false]"
                return
            }
            $params = @{ Identity = $identity }
            if ($org) { $params.OrganizationApp = $true }
            if ($mailbox) { $params.Mailbox = $mailbox }
            if ($manifestPath) {
                $bytes = Get-ManifestBytes $manifestPath
                if (-not $bytes) { return }
                $params.FileData = $bytes
            } elseif ($manifestUrl) {
                $params.ManifestURL = $manifestUrl
            }
            if ($providedTo) { $params.ProvidedTo = $providedTo }
            if ($userListRaw) { $params.UserList = Parse-CommaList $userListRaw }
            if ($defaultState) { $params.DefaultStateForUser = $defaultState }
            if ($enabledRaw) { $params.Enabled = (Parse-Bool $enabledRaw $true) }
            if ($params.Keys.Count -le 1) {
                Write-Warn "No update parameters provided."
                return
            }
            try {
                Set-App @params | Out-Null
                Write-Info "Add-in updated."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "enable" {
            if (-not $identity) {
                Write-Warn "Usage: addin exo enable <appId> [--org] [--mailbox <upn>]"
                return
            }
            try {
                if ($org) {
                    Set-App -OrganizationApp -Identity $identity -Enabled $true | Out-Null
                } elseif ($mailbox) {
                    Enable-App -Mailbox $mailbox -Identity $identity | Out-Null
                } else {
                    Write-Warn "Use --org for organization add-ins or --mailbox for user add-ins."
                    return
                }
                Write-Info "Add-in enabled."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "disable" {
            if (-not $identity) {
                Write-Warn "Usage: addin exo disable <appId> [--org] [--mailbox <upn>]"
                return
            }
            try {
                if ($org) {
                    Set-App -OrganizationApp -Identity $identity -Enabled $false | Out-Null
                } elseif ($mailbox) {
                    Disable-App -Mailbox $mailbox -Identity $identity | Out-Null
                } else {
                    Write-Warn "Use --org for organization add-ins or --mailbox for user add-ins."
                    return
                }
                Write-Info "Add-in disabled."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "remove" {
            if (-not $identity) {
                Write-Warn "Usage: addin exo remove <appId> [--org] [--mailbox <upn>] [--force]"
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
                if ($mailbox) {
                    Remove-App -Mailbox $mailbox -Identity $identity -Confirm:$false | Out-Null
                } elseif ($org) {
                    Remove-App -OrganizationApp -Identity $identity -Confirm:$false | Out-Null
                } else {
                    Remove-App -Identity $identity -Confirm:$false | Out-Null
                }
                Write-Info "Add-in removed."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        default {
            Write-Warn "Usage: addin exo list|get|install|update|remove|enable|disable|refresh [--org] [--mailbox <upn>]"
        }
    }
}


function Handle-AddinOnSendCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: addin onsend status|enable|disable [--policy <name>] [--user <upn>] [--all] [--filter <filter>]"
        return
    }
    if (-not (Require-ExoConnection)) { return }

    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $policy = Get-ArgValue $parsed.Map "policy"
    $user = Get-ArgValue $parsed.Map "user"
    $all = $parsed.Map.ContainsKey("all")
    $filter = Get-ArgValue $parsed.Map "filter"

    switch ($action) {
        "status" {
            try {
                if ($user) {
                    $mb = Get-CASMailbox -Identity $user
                    Write-Host ("User policy: " + $mb.OwaMailboxPolicy)
                    if ($mb.OwaMailboxPolicy) {
                        Get-OwaMailboxPolicy -Identity $mb.OwaMailboxPolicy | Select-Object Name, OnSendAddinsEnabled | Format-Table -AutoSize
                    }
                    return
                }
                if ($policy) {
                    Get-OwaMailboxPolicy -Identity $policy | Select-Object Name, OnSendAddinsEnabled | Format-Table -AutoSize
                    return
                }
                Get-OwaMailboxPolicy | Select-Object Name, OnSendAddinsEnabled | Format-Table -AutoSize
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "enable" {
            if (-not $policy) {
                Write-Warn "Usage: addin onsend enable --policy <name> [--user <upn>] [--all] [--filter <filter>]"
                return
            }
            try {
                $pol = Get-OwaMailboxPolicy -Identity $policy -ErrorAction SilentlyContinue
                if (-not $pol) {
                    New-OwaMailboxPolicy -Name $policy | Out-Null
                }
                Set-OwaMailboxPolicy -Identity $policy -OnSendAddinsEnabled $true | Out-Null
                if ($user) {
                    Set-CASMailbox -Identity $user -OwaMailboxPolicy $policy | Out-Null
                    Write-Info "OnSend enabled for user."
                    return
                }
                if ($all -or $filter) {
                    $f = if ($filter) { $filter } else { "RecipientTypeDetails -eq 'UserMailbox'" }
                    Get-User -Filter $f | Set-CASMailbox -OwaMailboxPolicy $policy | Out-Null
                    Write-Info "OnSend enabled for mailbox scope."
                }
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "disable" {
            if (-not $policy) {
                Write-Warn "Usage: addin onsend disable --policy <name> [--user <upn>] [--all] [--filter <filter>]"
                return
            }
            try {
                $pol = Get-OwaMailboxPolicy -Identity $policy -ErrorAction SilentlyContinue
                if (-not $pol) {
                    Write-Warn "Policy not found."
                    return
                }
                Set-OwaMailboxPolicy -Identity $policy -OnSendAddinsEnabled $false | Out-Null
                if ($user) {
                    Set-CASMailbox -Identity $user -OwaMailboxPolicy $policy | Out-Null
                    Write-Info "OnSend disabled for user."
                    return
                }
                if ($all -or $filter) {
                    $f = if ($filter) { $filter } else { "RecipientTypeDetails -eq 'UserMailbox'" }
                    Get-User -Filter $f | Set-CASMailbox -OwaMailboxPolicy $policy | Out-Null
                    Write-Info "OnSend disabled for mailbox scope."
                }
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        default {
            Write-Warn "Usage: addin onsend status|enable|disable [--policy <name>] [--user <upn>] [--all] [--filter <filter>]"
        }
    }
}


function Handle-AddinOrgCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: addin org connect|disconnect|status|list|get|create|update|remove|assign|enable|disable|refresh"
        return
    }

    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $id = $parsed.Positionals | Select-Object -First 1
    if (-not $id) { $id = Get-ArgValue $parsed.Map "productId" }
    if (-not $id) { $id = Get-ArgValue $parsed.Map "id" }

    switch ($action) {
        "connect" {
            if (-not (Ensure-OrgAddinModule)) { return }
            $upn = Get-ArgValue $parsed.Map "upn"
            try {
                if ($upn) {
                    Connect-OrganizationAddInService -UserPrincipalName $upn | Out-Null
                } else {
                    Connect-OrganizationAddInService | Out-Null
                }
                $global:OrgAddinConnected = $true
                Write-Info "Connected to Office add-ins service."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "disconnect" {
            if (-not (Ensure-OrgAddinModule)) { return }
            try {
                if (Get-Command -Name Disconnect-OrganizationAddInService -ErrorAction SilentlyContinue) {
                    Disconnect-OrganizationAddInService | Out-Null
                }
                $global:OrgAddinConnected = $false
                Write-Info "Disconnected."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "status" {
            $status = if ($global:OrgAddinConnected) { "connected" } else { "not connected" }
            Write-Host ("Office add-ins: " + $status)
        }
        "list" {
            if (-not (Require-OrgAddinConnection)) { return }
            try {
                $items = Get-OrganizationAddIn
                if ($items) {
                    $items | Select-Object ProductId, Title, Enabled, Version, ProviderName | Format-Table -AutoSize
                }
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "get" {
            if (-not (Require-OrgAddinConnection)) { return }
            if (-not $id) {
                Write-Warn "Usage: addin org get <productId>"
                return
            }
            try {
                Get-OrganizationAddIn -ProductId $id | Format-List *
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "create" { $action = "install" }
        "install" {
            if (-not (Require-OrgAddinConnection)) { return }
            $manifestPath = Get-ArgValue $parsed.Map "manifestPath"
            if (-not $manifestPath) { $manifestPath = Get-ArgValue $parsed.Map "path" }
            $assetId = Get-ArgValue $parsed.Map "assetId"
            $locale = Get-ArgValue $parsed.Map "locale"
            $contentMarket = Get-ArgValue $parsed.Map "contentMarket"
            $membersRaw = Get-ArgValue $parsed.Map "members"
            if (-not $membersRaw) { $membersRaw = Get-ArgValue $parsed.Map "users" }
            $params = @{}
            if ($manifestPath) { $params.ManifestPath = $manifestPath }
            if ($assetId) { $params.AssetId = $assetId }
            if (-not $manifestPath -and -not $assetId) {
                Write-Warn "Usage: addin org install --manifestPath <file.xml|url> OR --assetId <storeId> [--locale en-US] [--contentMarket en-US] [--members a,b]"
                return
            }
            if ($locale) { $params.Locale = $locale }
            if ($contentMarket) { $params.ContentMarket = $contentMarket }
            if ($membersRaw) { $params.Members = Parse-CommaList $membersRaw }
            try {
                $resp = New-OrganizationAddIn @params
                if ($resp) { Write-Info "Add-in created." }
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "update" { $action = "refresh" }
        "refresh" {
            if (-not (Require-OrgAddinConnection)) { return }
            if (-not $id) {
                Write-Warn "Usage: addin org refresh <productId> [--manifestPath <file.xml|url>] [--enabled true|false] [--locale en-US] [--contentMarket en-US]"
                return
            }
            $manifestPath = Get-ArgValue $parsed.Map "manifestPath"
            if (-not $manifestPath) { $manifestPath = Get-ArgValue $parsed.Map "path" }
            $assetId = Get-ArgValue $parsed.Map "assetId"
            $locale = Get-ArgValue $parsed.Map "locale"
            $contentMarket = Get-ArgValue $parsed.Map "contentMarket"
            $enabledRaw = Get-ArgValue $parsed.Map "enabled"
            $params = @{ ProductId = $id }
            if ($manifestPath) { $params.ManifestPath = $manifestPath }
            if ($assetId) { $params.AssetId = $assetId }
            if ($locale) { $params.Locale = $locale }
            if ($contentMarket) { $params.ContentMarket = $contentMarket }
            if ($enabledRaw) { $params.Enabled = (Parse-Bool $enabledRaw $true) }
            if ($params.Keys.Count -le 1) {
                Write-Warn "No update parameters provided."
                return
            }
            try {
                Set-OrganizationAddIn @params | Out-Null
                Write-Info "Add-in updated."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "enable" {
            if (-not (Require-OrgAddinConnection)) { return }
            if (-not $id) {
                Write-Warn "Usage: addin org enable <productId>"
                return
            }
            try {
                Set-OrganizationAddIn -ProductId $id -Enabled $true | Out-Null
                Write-Info "Add-in enabled."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "disable" {
            if (-not (Require-OrgAddinConnection)) { return }
            if (-not $id) {
                Write-Warn "Usage: addin org disable <productId>"
                return
            }
            try {
                Set-OrganizationAddIn -ProductId $id -Enabled $false | Out-Null
                Write-Info "Add-in disabled."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "assign" {
            if (-not (Require-OrgAddinConnection)) { return }
            if (-not $id) {
                Write-Warn "Usage: addin org assign <productId> [--members a,b] [--add|--remove|--everyone]"
                return
            }
            $membersRaw = Get-ArgValue $parsed.Map "members"
            if (-not $membersRaw) { $membersRaw = Get-ArgValue $parsed.Map "users" }
            $members = if ($membersRaw) { Parse-CommaList $membersRaw } else { @() }
            $add = $parsed.Map.ContainsKey("add")
            $remove = $parsed.Map.ContainsKey("remove")
            $everyone = $parsed.Map.ContainsKey("everyone") -or $parsed.Map.ContainsKey("assignToEveryone")
            if ($everyone) {
                Set-OrganizationAddInAssignments -ProductId $id -AssignToEveryone | Out-Null
                Write-Info "Assigned to everyone."
                return
            }
            if (-not $members -or $members.Count -eq 0) {
                Write-Warn "Members required. Use --members a,b or --everyone."
                return
            }
            try {
                if ($remove) {
                    Set-OrganizationAddInAssignments -ProductId $id -Remove -Members $members | Out-Null
                } elseif ($add) {
                    Set-OrganizationAddInAssignments -ProductId $id -Add -Members $members | Out-Null
                } else {
                    Set-OrganizationAddInAssignments -ProductId $id -Members $members | Out-Null
                }
                Write-Info "Assignments updated."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "remove" {
            if (-not (Require-OrgAddinConnection)) { return }
            if (-not $id) {
                Write-Warn "Usage: addin org remove <productId> [--force]"
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
                Remove-OrganizationAddIn -ProductId $id | Out-Null
                Write-Info "Add-in removed."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "cmd" {
            if (-not (Require-OrgAddinConnection)) { return }
            $cmdlet = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $cmdlet) { $cmdlet = $parsed.Positionals | Select-Object -First 1 }
            if (-not $cmdlet) {
                Write-Warn "Usage: addin org cmd <cmdlet> [--params key=value[,key=value]] [--json <payload>] [--bodyFile <file>] [--set key=value]"
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
        default {
            Write-Warn "Usage: addin org connect|disconnect|status|list|get|create|update|remove|assign|enable|disable|refresh"
        }
    }
}

