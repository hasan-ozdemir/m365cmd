# Handler: Teams
# Purpose: Teams command handlers.
function Require-TeamsConnection {
    if (-not (Ensure-ModuleLoaded "MicrosoftTeams")) { return $false }
    if ($global:TeamsConnected) { return $true }
    Write-Warn "Not connected to Microsoft Teams. Use: teams connect"
    return $false
}


function Handle-TeamsCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: teams connect|disconnect|status|list|get|create|delete|user|channel|policy|config|cmd|cmdlets"
        return
    }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "connect" {
            if (-not (Ensure-ModuleLoaded "MicrosoftTeams")) { return }
            $account = Get-ArgValue $parsed.Map "upn"
            if (-not $account) { $account = $global:Config.admin.defaultUpn }
            $tenantId = Get-ArgValue $parsed.Map "tenantId"
            $params = @{}
            if ($account) { $params.AccountId = $account }
            if ($tenantId) { $params.TenantId = $tenantId }
            try {
                Connect-MicrosoftTeams @params | Out-Null
                $global:TeamsConnected = $true
                Write-Info "Connected to Microsoft Teams."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "disconnect" {
            if (-not (Ensure-ModuleLoaded "MicrosoftTeams")) { return }
            try {
                if (Get-Command Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue) {
                    Disconnect-MicrosoftTeams | Out-Null
                }
                $global:TeamsConnected = $false
                Write-Info "Disconnected from Microsoft Teams."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "status" {
            $status = if ($global:TeamsConnected) { "connected" } else { "not connected" }
            Write-Host ("Microsoft Teams: " + $status)
        }
        "cmdlets" {
            if (-not (Ensure-ModuleLoaded "MicrosoftTeams")) { return }
            $filter = Get-ArgValue $parsed.Map "filter"
            $cmds = Get-Command -Module MicrosoftTeams -ErrorAction SilentlyContinue
            if ($filter) {
                $cmds = $cmds | Where-Object { $_.Name -like ("*" + $filter + "*") }
            }
            $cmds | Sort-Object Name | Select-Object -ExpandProperty Name | Format-Wide -Column 3
        }
        "cmd" {
            if (-not (Ensure-ModuleLoaded "MicrosoftTeams")) { return }
            if (-not $global:TeamsConnected) {
                Write-Warn "Not connected to Microsoft Teams. Use: teams connect"
                return
            }
            $cmdlet = $parsed.Positionals | Select-Object -First 1
            if (-not $cmdlet) {
                Write-Warn "Usage: teams cmd <cmdlet> [--params key=value[,key=value]] [--json <payload>] [--bodyFile <file>] [--set key=value]"
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
        "config" {
            if (-not (Require-TeamsConnection)) { return }
            $action = $parsed.Positionals | Select-Object -First 1
            if (-not $action) {
                Write-Warn "Usage: teams config get|update"
                return
            }
            switch ($action) {
                "get" {
                    try {
                        Get-CsTeamsClientConfiguration | Format-List *
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "update" {
                    $setRaw = Get-ArgValue $parsed.Map "set"
                    if (-not $setRaw) {
                        Write-Warn "Usage: teams config update --set key=value[,key=value]"
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
                        Set-CsTeamsClientConfiguration @params | Out-Null
                        Write-Info "Teams client configuration updated."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                default {
                    Write-Warn "Usage: teams config get|update"
                }
            }
        }
        "policy" {
            if (-not (Require-TeamsConnection)) { return }
            $action = $parsed.Positionals | Select-Object -First 1
            $name = $parsed.Positionals | Select-Object -Skip 1 -First 1
            $type = Get-ArgValue $parsed.Map "type"
            if (-not $action -or -not $type) {
                Write-Warn "Usage: teams policy list|get|create|update|delete --type messaging|meeting [name] [--set key=value]"
                return
            }
            $t = $type.ToLowerInvariant()
            $map = @{}
            if ($t -eq "messaging") {
                $map = @{ Get = "Get-CsTeamsMessagingPolicy"; New = "New-CsTeamsMessagingPolicy"; Set = "Set-CsTeamsMessagingPolicy"; Remove = "Remove-CsTeamsMessagingPolicy" }
            } elseif ($t -eq "meeting") {
                $map = @{ Get = "Get-CsTeamsMeetingPolicy"; New = "New-CsTeamsMeetingPolicy"; Set = "Set-CsTeamsMeetingPolicy"; Remove = "Remove-CsTeamsMeetingPolicy" }
            } else {
                Write-Warn "Unknown policy type. Use messaging|meeting."
                return
            }

            switch ($action) {
                "list" {
                    try {
                        & $map.Get | Format-Table -AutoSize
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "get" {
                    if (-not $name) {
                        Write-Warn "Usage: teams policy get --type messaging|meeting <name>"
                        return
                    }
                    try {
                        & $map.Get -Identity $name | Format-List *
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "create" {
                    if (-not $name) {
                        Write-Warn "Usage: teams policy create --type messaging|meeting <name> [--set key=value]"
                        return
                    }
                    $setRaw = Get-ArgValue $parsed.Map "set"
                    $body = @{}
                    if ($setRaw) {
                        $body = Parse-Value $setRaw
                        if ($body -is [string]) { $body = Parse-KvPairs $setRaw }
                    }
                    try {
                        $params = @{ Identity = $name }
                        foreach ($k in $body.Keys) { $params[$k] = $body[$k] }
                        & $map.New @params | Out-Null
                        Write-Info "Policy created."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "update" {
                    if (-not $name) {
                        Write-Warn "Usage: teams policy update --type messaging|meeting <name> --set key=value"
                        return
                    }
                    $setRaw = Get-ArgValue $parsed.Map "set"
                    if (-not $setRaw) {
                        Write-Warn "Usage: teams policy update --type messaging|meeting <name> --set key=value"
                        return
                    }
                    $body = Parse-Value $setRaw
                    if ($body -is [string]) { $body = Parse-KvPairs $setRaw }
                    if ($body.Keys.Count -eq 0) {
                        Write-Warn "No properties to update."
                        return
                    }
                    try {
                        $params = @{ Identity = $name }
                        foreach ($k in $body.Keys) { $params[$k] = $body[$k] }
                        & $map.Set @params | Out-Null
                        Write-Info "Policy updated."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "delete" {
                    if (-not $name) {
                        Write-Warn "Usage: teams policy delete --type messaging|meeting <name> [--force]"
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
                        & $map.Remove -Identity $name -Confirm:$false | Out-Null
                        Write-Info "Policy deleted."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                default {
                    Write-Warn "Usage: teams policy list|get|create|update|delete --type messaging|meeting [name]"
                }
            }
        }
        "list" {
            if (-not (Require-TeamsConnection)) { return }
            try {
                Get-Team | Select-Object GroupId, DisplayName, Visibility | Format-Table -AutoSize
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "get" {
            if (-not (Require-TeamsConnection)) { return }
            $teamId = $parsed.Positionals | Select-Object -First 1
            if (-not $teamId) {
                Write-Warn "Usage: teams get <groupId>"
                return
            }
            try {
                Get-Team -GroupId $teamId | Format-List *
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "create" {
            if (-not (Require-TeamsConnection)) { return }
            $name = Get-ArgValue $parsed.Map "name"
            if (-not $name) { $name = Get-ArgValue $parsed.Map "displayName" }
            $desc = Get-ArgValue $parsed.Map "description"
            $visibility = Get-ArgValue $parsed.Map "visibility"
            if (-not $visibility) { $visibility = "Private" }
            if (-not $name) {
                Write-Warn "Usage: teams create --name <name> [--description text] [--visibility Private|Public]"
                return
            }
            try {
                New-Team -DisplayName $name -Description $desc -Visibility $visibility | Out-Null
                Write-Info "Team created."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "delete" {
            if (-not (Require-TeamsConnection)) { return }
            $teamId = $parsed.Positionals | Select-Object -First 1
            if (-not $teamId) {
                Write-Warn "Usage: teams delete <groupId> [--force]"
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
                Remove-Team -GroupId $teamId -Confirm:$false | Out-Null
                Write-Info "Team removed."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "user" {
            if (-not (Require-TeamsConnection)) { return }
            $action = $parsed.Positionals | Select-Object -First 1
            $teamId = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action -or -not $teamId) {
                Write-Warn "Usage: teams user list|add|remove <groupId> [--user <upn>] [--role Owner|Member]"
                return
            }
            switch ($action) {
                "list" {
                    try {
                        Get-TeamUser -GroupId $teamId | Select-Object User, Role | Format-Table -AutoSize
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "add" {
                    $user = Get-ArgValue $parsed.Map "user"
                    $role = Get-ArgValue $parsed.Map "role"
                    if (-not $role) { $role = "Member" }
                    if (-not $user) {
                        Write-Warn "Usage: teams user add <groupId> --user <upn> [--role Owner|Member]"
                        return
                    }
                    try {
                        Add-TeamUser -GroupId $teamId -User $user -Role $role | Out-Null
                        Write-Info "Team user added."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "remove" {
                    $user = Get-ArgValue $parsed.Map "user"
                    if (-not $user) {
                        Write-Warn "Usage: teams user remove <groupId> --user <upn>"
                        return
                    }
                    try {
                        Remove-TeamUser -GroupId $teamId -User $user | Out-Null
                        Write-Info "Team user removed."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                default {
                    Write-Warn "Usage: teams user list|add|remove <groupId> [--user <upn>] [--role Owner|Member]"
                }
            }
        }
        "channel" {
            if (-not (Require-TeamsConnection)) { return }
            $action = $parsed.Positionals | Select-Object -First 1
            $teamId = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action -or -not $teamId) {
                Write-Warn "Usage: teams channel list|create|remove <groupId> [--name <name>] [--id <id>] [--description text] [--type Standard|Private|Shared]"
                return
            }
            switch ($action) {
                "list" {
                    try {
                        Get-TeamChannel -GroupId $teamId | Select-Object Id, DisplayName, MembershipType | Format-Table -AutoSize
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "create" {
                    $name = Get-ArgValue $parsed.Map "name"
                    $desc = Get-ArgValue $parsed.Map "description"
                    $type = Get-ArgValue $parsed.Map "type"
                    if (-not $type) { $type = "Standard" }
                    if (-not $name) {
                        Write-Warn "Usage: teams channel create <groupId> --name <name> [--description text] [--type Standard|Private|Shared]"
                        return
                    }
                    try {
                        New-TeamChannel -GroupId $teamId -DisplayName $name -Description $desc -MembershipType $type | Out-Null
                        Write-Info "Channel created."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "remove" {
                    $id = Get-ArgValue $parsed.Map "id"
                    $name = Get-ArgValue $parsed.Map "name"
                    if (-not $id -and -not $name) {
                        Write-Warn "Usage: teams channel remove <groupId> --id <id> OR --name <displayName>"
                        return
                    }
                    try {
                        if ($id) {
                            Remove-TeamChannel -GroupId $teamId -Id $id | Out-Null
                        } else {
                            Remove-TeamChannel -GroupId $teamId -DisplayName $name | Out-Null
                        }
                        Write-Info "Channel removed."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                default {
                    Write-Warn "Usage: teams channel list|create|remove <groupId>"
                }
            }
        }
        default {
            Write-Warn "Usage: teams connect|disconnect|status|list|get|create|delete|user|channel|policy|config|cmd|cmdlets"
        }
    }
}

