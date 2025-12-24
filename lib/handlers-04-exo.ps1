# Handler: Exo
# Purpose: Exo command handlers.
function Require-ExoConnection {
    if (-not (Ensure-ModuleLoaded "ExchangeOnlineManagement")) { return $false }
    if ($global:ExoConnected) { return $true }
    if (Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue) {
        try {
            $info = Get-ConnectionInformation | Select-Object -First 1
            if ($info) {
                $global:ExoConnected = $true
                return $true
            }
        } catch {}
    }
    Write-Warn "Not connected to Exchange Online. Use: exo connect"
    return $false
}

function Handle-ExoCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: exo connect|disconnect|status|mailbox|addin|onsend"
        return
    }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "connect" {
            if (-not (Ensure-ModuleLoaded "ExchangeOnlineManagement")) { return }
            $upn = Get-ArgValue $parsed.Map "upn"
            if (-not $upn) { $upn = $global:Config.admin.defaultUpn }
            $delegated = Get-ArgValue $parsed.Map "delegatedOrg"
            $env = Get-ArgValue $parsed.Map "env"
            $disableWam = Parse-Bool (Get-ArgValue $parsed.Map "disableWam") $false
            $params = @{}
            if ($upn) { $params.UserPrincipalName = $upn }
            if ($delegated) { $params.DelegatedOrganization = $delegated }
            if ($env) { $params.ExchangeEnvironmentName = $env }
            if ($disableWam) { $params.DisableWAM = $true }
            try {
                Connect-ExchangeOnline @params | Out-Null
                $global:ExoConnected = $true
                Write-Info "Connected to Exchange Online."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "disconnect" {
            if (-not (Ensure-ModuleLoaded "ExchangeOnlineManagement")) { return }
            try {
                if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) {
                    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
                }
                $global:ExoConnected = $false
                Write-Info "Disconnected from Exchange Online."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "status" {
            $status = if ($global:ExoConnected) { "connected" } else { "not connected" }
            Write-Host ("Exchange Online: " + $status)
        }
        "cmdlets" {
            if (-not (Ensure-ModuleLoaded "ExchangeOnlineManagement")) { return }
            $filter = Get-ArgValue $parsed.Map "filter"
            $cmds = Get-Command -Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
            if ($filter) {
                $cmds = $cmds | Where-Object { $_.Name -like ("*" + $filter + "*") }
            }
            $cmds | Sort-Object Name | Select-Object -ExpandProperty Name | Format-Wide -Column 3
        }
        "cmd" {
            if (-not (Ensure-ModuleLoaded "ExchangeOnlineManagement")) { return }
            if (-not $global:ExoConnected) {
                Write-Warn "Not connected to Exchange Online. Use: exo connect"
                return
            }
            $cmdlet = $parsed.Positionals | Select-Object -First 1
            if (-not $cmdlet) {
                Write-Warn "Usage: exo cmd <cmdlet> [--params key=value[,key=value]] [--json <payload>] [--bodyFile <file>] [--set key=value]"
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
        "mailbox" {
            if (-not (Require-ExoConnection)) { return }
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action) {
                Write-Warn "Usage: exo mailbox list|get|create|update|delete|perm"
                return
            }
            $useExo = $null -ne (Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue)
            switch ($action) {
                "list" {
                    $filter = Get-ArgValue $parsed.Map "filter"
                    $select = Get-ArgValue $parsed.Map "select"
                    $props = if ($select) { Parse-CommaList $select } else { @("DisplayName","PrimarySmtpAddress","RecipientTypeDetails") }
                    try {
                        if ($useExo) {
                            $params = @{ ResultSize = "Unlimited" }
                            if ($filter) { $params.Filter = $filter }
                            $items = Get-EXOMailbox @params
                        } else {
                            $items = if ($filter) { Get-Mailbox -Filter $filter -ResultSize Unlimited } else { Get-Mailbox -ResultSize Unlimited }
                        }
                        $items | Select-Object -Property $props | Format-Table -AutoSize
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "get" {
                    if (-not $identity) {
                        Write-Warn "Usage: exo mailbox get <identity>"
                        return
                    }
                    try {
                        if ($useExo) {
                            Get-EXOMailbox -Identity $identity | Format-List *
                        } else {
                            Get-Mailbox -Identity $identity | Format-List *
                        }
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "update" {
                    if (-not $identity) {
                        Write-Warn "Usage: exo mailbox update <identity> --set key=value[,key=value]"
                        return
                    }
                    $setRaw = Get-ArgValue $parsed.Map "set"
                    if (-not $setRaw) {
                        Write-Warn "Usage: exo mailbox update <identity> --set key=value[,key=value]"
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
                        Set-Mailbox -Identity $identity @params | Out-Null
                        Write-Info "Mailbox updated."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "create" {
                    $type = Get-ArgValue $parsed.Map "type"
                    $name = Get-ArgValue $parsed.Map "name"
                    $alias = Get-ArgValue $parsed.Map "alias"
                    $upn = Get-ArgValue $parsed.Map "upn"
                    $primary = Get-ArgValue $parsed.Map "primarySmtp"
                    $password = Get-ArgValue $parsed.Map "password"
                    if (-not $type -or -not $name) {
                        Write-Warn "Usage: exo mailbox create --type shared|room|equipment|user --name <displayName> [--alias <alias>] [--upn <user@domain>] [--primarySmtp <addr>] [--password <pwd>]"
                        return
                    }
                    $t = $type.ToLowerInvariant()
                    $params = @{ Name = $name }
                    if ($alias) { $params.Alias = $alias }
                    if ($primary) { $params.PrimarySmtpAddress = $primary }
                    if ($upn) { $params.UserPrincipalName = $upn }
                    if ($t -eq "shared") {
                        $params.Shared = $true
                    } elseif ($t -eq "room") {
                        $params.Room = $true
                    } elseif ($t -eq "equipment") {
                        $params.Equipment = $true
                    } elseif ($t -eq "user") {
                        if (-not $upn -or -not $password) {
                            Write-Warn "User mailbox requires --upn and --password."
                            return
                        }
                        $secure = ConvertTo-SecureString -String $password -AsPlainText -Force
                        $params.Password = $secure
                    } else {
                        Write-Warn "Unknown mailbox type. Use shared|room|equipment|user."
                        return
                    }
                    try {
                        New-Mailbox @params | Out-Null
                        Write-Info "Mailbox created."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "delete" {
                    if (-not $identity) {
                        Write-Warn "Usage: exo mailbox delete <identity> [--force]"
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
                        Remove-Mailbox -Identity $identity -Confirm:$false | Out-Null
                        Write-Info "Mailbox deleted."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "perm" {
                    $permAction = $parsed.Positionals | Select-Object -Skip 2 -First 1
                    if ($identity -in @("list","add","remove") -and $permAction) {
                        $tmp = $identity
                        $identity = $permAction
                        $permAction = $tmp
                    }
                    if (-not $identity -or -not $permAction) {
                        Write-Warn "Usage: exo mailbox perm <mailbox> list|add|remove [--user <upn>] [--rights FullAccess]"
                        return
                    }
                    $user = Get-ArgValue $parsed.Map "user"
                    $rights = Get-ArgValue $parsed.Map "rights"
                    if (-not $rights) { $rights = "FullAccess" }
                    switch ($permAction) {
                        "list" {
                            try {
                                Get-MailboxPermission -Identity $identity | Select-Object User, AccessRights, IsInherited, Deny | Format-Table -AutoSize
                            } catch {
                                Write-Err $_.Exception.Message
                            }
                        }
                        "add" {
                            if (-not $user) {
                                Write-Warn "Usage: exo mailbox perm <mailbox> add --user <upn> [--rights FullAccess]"
                                return
                            }
                            try {
                                Add-MailboxPermission -Identity $identity -User $user -AccessRights $rights -InheritanceType All | Out-Null
                                Write-Info "Permission added."
                            } catch {
                                Write-Err $_.Exception.Message
                            }
                        }
                        "remove" {
                            if (-not $user) {
                                Write-Warn "Usage: exo mailbox perm <mailbox> remove --user <upn> [--rights FullAccess]"
                                return
                            }
                            try {
                                Remove-MailboxPermission -Identity $identity -User $user -AccessRights $rights -Confirm:$false | Out-Null
                                Write-Info "Permission removed."
                            } catch {
                                Write-Err $_.Exception.Message
                            }
                        }
                        default {
                            Write-Warn "Usage: exo mailbox perm <mailbox> list|add|remove [--user <upn>] [--rights FullAccess]"
                        }
                    }
                }
                default {
                    Write-Warn "Usage: exo mailbox list|get|create|update|delete|perm"
                }
            }
        }
        "addin" {
            Handle-AddinExoCommand $rest
        }
        "onsend" {
            Handle-AddinOnSendCommand $rest
        }
        default {
            Write-Warn "Usage: exo connect|disconnect|status|mailbox|cmd|cmdlets|addin|onsend"
        }
    }
}

