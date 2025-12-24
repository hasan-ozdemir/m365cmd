# Handler: Admin
# Purpose: Admin command handlers.
function Handle-ModuleCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0 -or $InputArgs[0] -eq "list") {
        $required = @($global:Config.modules.required)
        $optional = @($global:Config.modules.optional)
        Write-Host "Required modules:"
        foreach ($name in $required) {
            $ok = Test-ModuleAvailable $name
            Write-Host ("  " + $name + " : " + ($(if ($ok) { "available" } else { "missing" })))
        }
        Write-Host "Optional modules:"
        foreach ($name in $optional) {
            $ok = Test-ModuleAvailable $name
            Write-Host ("  " + $name + " : " + ($(if ($ok) { "available" } else { "missing" })))
        }
        return
    }

    $sub = $InputArgs[0].ToLowerInvariant()
    $name = if ($InputArgs.Count -ge 2) { $InputArgs[1] } else { "" }
    if (-not $name) {
        Write-Warn "Module name required."
        return
    }

    switch ($sub) {
        "install" {
            try {
                Set-LocalModulePath
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            } catch {}
            if (-not (Get-Command -Name Save-Module -ErrorAction SilentlyContinue)) {
                Write-Err "Save-Module is not available. Install PowerShellGet."
                return
            }
            try {
                if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                    Install-PackageProvider -Name NuGet -Force -Scope CurrentUser | Out-Null
                }
                Save-Module -Name $name -Path $Paths.Modules -Force -ErrorAction Stop
                Write-Info ("Saved module: " + $name)
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "update" {
            try {
                Set-LocalModulePath
                if (-not (Get-Command -Name Save-Module -ErrorAction SilentlyContinue)) {
                    Write-Err "Save-Module is not available. Install PowerShellGet."
                    return
                }
                Save-Module -Name $name -Path $Paths.Modules -Force -ErrorAction Stop
                Write-Info ("Updated module: " + $name)
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "remove" {
            $target = Join-Path $Paths.Modules $name
            if (Test-Path $target) {
                Remove-Item -Path $target -Recurse -Force
                Write-Info ("Removed module: " + $name)
            } else {
                Write-Warn "Module not found in local modules path."
            }
        }
        default {
            Write-Warn "Unknown module subcommand. Use: module list|install|update|remove <name>"
        }
    }
}



function ConvertTo-Hashtable {
    param([object]$InputObject)
    if (-not $InputObject) { return @{} }
    if ($InputObject -is [hashtable]) { return $InputObject }
    $tmp = @{}
    foreach ($p in $InputObject.PSObject.Properties) {
        $tmp[$p.Name] = $p.Value
    }
    return $tmp
}


function Resolve-PrincipalIdFromMap {
    param(
        [hashtable]$Map,
        [string]$FallbackIdentity
    )
    if (-not $Map) { $Map = @{} }
    $principalId = Get-ArgValue $Map "principalId"
    if ($principalId) { return $principalId }
    $user = Get-ArgValue $Map "user"
    $group = Get-ArgValue $Map "group"
    $sp = Get-ArgValue $Map "sp"
    $principal = Get-ArgValue $Map "principal"

    if ($user) {
        $u = Resolve-UserObject $user
        if ($u) { return $u.Id }
    }
    if ($group) {
        $g = Resolve-GroupObject $group
        if ($g) { return $g.Id }
    }
    if ($sp) {
        $s = Resolve-ServicePrincipalObject $sp
        if ($s) { return $s.Id }
    }
    if ($principal) {
        $u = Resolve-UserObject $principal
        if ($u) { return $u.Id }
        $g = Resolve-GroupObject $principal
        if ($g) { return $g.Id }
        $s = Resolve-ServicePrincipalObject $principal
        if ($s) { return $s.Id }
    }
    if ($FallbackIdentity) {
        $u = Resolve-UserObject $FallbackIdentity
        if ($u) { return $u.Id }
        $g = Resolve-GroupObject $FallbackIdentity
        if ($g) { return $g.Id }
        $s = Resolve-ServicePrincipalObject $FallbackIdentity
        if ($s) { return $s.Id }
    }
    return $null
}


function Resolve-RoleDefinitionObject {
    param(
        [string]$Identity,
        [string]$Api,
        [bool]$AllowFallback = $true
    )
    if (-not $Identity) { return $null }
    if ($Identity -match "^[0-9a-fA-F-]{36}$") {
        return Invoke-GraphRequestAuto -Method "GET" -Uri ("/roleManagement/directory/roleDefinitions/" + $Identity) -Api $Api -AllowFallback:$AllowFallback
    }
    $esc = Escape-ODataString $Identity
    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ("/roleManagement/directory/roleDefinitions?$filter=displayName eq '$esc'") -Api $Api -AllowFallback:$AllowFallback
    if ($resp -and $resp.value) { return ($resp.value | Select-Object -First 1) }
    return $null
}


function Resolve-RoleDefinitionId {
    param(
        [string]$Identity,
        [string]$Api,
        [bool]$AllowFallback = $true
    )
    $obj = Resolve-RoleDefinitionObject -Identity $Identity -Api $Api -AllowFallback:$AllowFallback
    if ($obj -and $obj.Id) { return $obj.Id }
    return $null
}


function Invoke-UserDeleteOperation {
    param(
        [string]$Identity,
        [hashtable]$Map
    )
    $filter = Get-ArgValue $Map "filter"
    $idsRaw = Get-ArgValue $Map "ids"
    $whatIf = Parse-Bool (Get-ArgValue $Map "whatif") $false
    $force = Parse-Bool (Get-ArgValue $Map "force") $false

    if ($filter) {
        try {
            $users = Get-MgUser -Filter $filter -All -Property Id,DisplayName,UserPrincipalName
        } catch {
            Write-Err $_.Exception.Message
            return
        }
        if (-not $users) {
            Write-Info "No users matched."
            return
        }
        Write-Host ("Matched users: " + $users.Count)
        $users | Select-Object Id,DisplayName,UserPrincipalName | Format-Table -AutoSize
        if ($whatIf) { return }
        if (-not $force) {
            $confirm = Read-Host "Type DELETE to confirm bulk delete"
            if ($confirm -ne "DELETE") {
                Write-Info "Canceled."
                return
            }
        }
        foreach ($u in @($users)) {
            try {
                Remove-MgUser -UserId $u.Id -ErrorAction Stop
                Write-Info ("Deleted: " + $u.UserPrincipalName)
            } catch {
                Write-Err ("Failed: " + $u.UserPrincipalName + " : " + $_.Exception.Message)
            }
        }
        return
    }

    if ($idsRaw) {
        $ids = Parse-CommaList $idsRaw
        if (-not $ids -or $ids.Count -eq 0) {
            Write-Warn "No ids provided."
            return
        }
        $targets = @()
        foreach ($item in $ids) {
            $u = Resolve-UserObject $item
            if ($u) { $targets += $u }
        }
        if (-not $targets) {
            Write-Warn "No users resolved."
            return
        }
        Write-Host ("Matched users: " + $targets.Count)
        $targets | Select-Object Id,DisplayName,UserPrincipalName | Format-Table -AutoSize
        if ($whatIf) { return }
        if (-not $force) {
            $confirm = Read-Host "Type DELETE to confirm bulk delete"
            if ($confirm -ne "DELETE") {
                Write-Info "Canceled."
                return
            }
        }
        foreach ($u in @($targets)) {
            try {
                Remove-MgUser -UserId $u.Id -ErrorAction Stop
                Write-Info ("Deleted: " + $u.UserPrincipalName)
            } catch {
                Write-Err ("Failed: " + $u.UserPrincipalName + " : " + $_.Exception.Message)
            }
        }
        return
    }

    if (-not $Identity) {
        Write-Warn "Usage: user delete <upn|id> [--force] OR user delete --filter <odata> [--force] [--whatif] OR user delete --ids id1,id2 [--force]"
        return
    }
    if (-not $force) {
        $confirm = Read-Host "Type DELETE to confirm"
        if ($confirm -ne "DELETE") {
            Write-Info "Canceled."
            return
        }
    }
    try {
        Remove-MgUser -UserId $Identity -ErrorAction Stop
        Write-Info "User deleted."
    } catch {
        Write-Err $_.Exception.Message
    }
}


function Invoke-GroupDeleteOperation {
    param(
        [string]$Identity,
        [hashtable]$Map
    )
    $filter = Get-ArgValue $Map "filter"
    $idsRaw = Get-ArgValue $Map "ids"
    $whatIf = Parse-Bool (Get-ArgValue $Map "whatif") $false
    $force = Parse-Bool (Get-ArgValue $Map "force") $false

    if ($filter) {
        try {
            $groups = Get-MgGroup -Filter $filter -All -Property Id,DisplayName,Mail
        } catch {
            Write-Err $_.Exception.Message
            return
        }
        if (-not $groups) {
            Write-Info "No groups matched."
            return
        }
        Write-Host ("Matched groups: " + $groups.Count)
        $groups | Select-Object Id,DisplayName,Mail | Format-Table -AutoSize
        if ($whatIf) { return }
        if (-not $force) {
            $confirm = Read-Host "Type DELETE to confirm bulk delete"
            if ($confirm -ne "DELETE") {
                Write-Info "Canceled."
                return
            }
        }
        foreach ($g in @($groups)) {
            try {
                Remove-MgGroup -GroupId $g.Id -ErrorAction Stop
                Write-Info ("Deleted: " + $g.DisplayName)
            } catch {
                Write-Err ("Failed: " + $g.DisplayName + " : " + $_.Exception.Message)
            }
        }
        return
    }

    if ($idsRaw) {
        $ids = Parse-CommaList $idsRaw
        if (-not $ids -or $ids.Count -eq 0) {
            Write-Warn "No ids provided."
            return
        }
        $targets = @()
        foreach ($item in $ids) {
            $g = Resolve-GroupObject $item
            if ($g) { $targets += $g }
        }
        if (-not $targets) {
            Write-Warn "No groups resolved."
            return
        }
        Write-Host ("Matched groups: " + $targets.Count)
        $targets | Select-Object Id,DisplayName,Mail | Format-Table -AutoSize
        if ($whatIf) { return }
        if (-not $force) {
            $confirm = Read-Host "Type DELETE to confirm bulk delete"
            if ($confirm -ne "DELETE") {
                Write-Info "Canceled."
                return
            }
        }
        foreach ($g in @($targets)) {
            try {
                Remove-MgGroup -GroupId $g.Id -ErrorAction Stop
                Write-Info ("Deleted: " + $g.DisplayName)
            } catch {
                Write-Err ("Failed: " + $g.DisplayName + " : " + $_.Exception.Message)
            }
        }
        return
    }

    if (-not $Identity) {
        Write-Warn "Usage: group delete <id|displayName> [--force] OR group delete --filter <odata> [--force] [--whatif] OR group delete --ids id1,id2 [--force]"
        return
    }
    if (-not $force) {
        $confirm = Read-Host "Type DELETE to confirm"
        if ($confirm -ne "DELETE") {
            Write-Info "Canceled."
            return
        }
    }
    $group = Resolve-GroupObject $Identity
    if (-not $group) {
        Write-Warn "Group not found."
        return
    }
    try {
        Remove-MgGroup -GroupId $group.Id -ErrorAction Stop
        Write-Info "Group deleted."
    } catch {
        Write-Err $_.Exception.Message
    }
}


function Handle-UserCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: user list|get|create|update|delete|bulkdelete|props|apps|roles|enable|disable|password|upn|email|alias|session|mfa|license"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    if ($sub -eq "username" -or $sub -eq "name") { $sub = "upn" }
    if ($sub -eq "passwd") { $sub = "password" }

    switch ($sub) {
        "list" {
            $filter = Get-ArgValue $parsed.Map "filter"
            $select = Get-ArgValue $parsed.Map "select"
            $props = if ($select) { Parse-CommaList $select } else { @("Id", "DisplayName", "UserPrincipalName", "AccountEnabled") }
            try {
                $users = if ($filter) { Get-MgUser -Filter $filter -All -Property $props } else { Get-MgUser -All -Property $props }
                $users | Select-Object -Property $props | Format-Table -AutoSize
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "get" {
            $identity = $parsed.Positionals | Select-Object -First 1
            if (-not $identity) {
                Write-Warn "Usage: user get <upn|id>"
                return
            }
            try {
                $u = Get-MgUser -UserId $identity -Property *
                $u | ConvertTo-Json -Depth 6
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "props" {
            $identity = $parsed.Positionals | Select-Object -First 1
            if (-not $identity) {
                Write-Warn "Usage: user props <upn|id>"
                return
            }
            try {
                Get-MgUser -UserId $identity -Property * | Format-List *
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "create" {
            $upn = Get-ArgValue $parsed.Map "upn"
            $alias = Get-ArgValue $parsed.Map "alias"
            $domain = Get-ArgValue $parsed.Map "domain"
            $displayName = Get-ArgValue $parsed.Map "displayName"
            $password = Get-ArgValue $parsed.Map "password"
            if (-not $password) { $password = Get-ArgValue $parsed.Map "pwd" }
            $usageLocation = Get-ArgValue $parsed.Map "usageLocation"
            $forceChange = Parse-Bool (Get-ArgValue $parsed.Map "forceChange") $true
            $accountEnabled = Parse-Bool (Get-ArgValue $parsed.Map "enabled") $true
            $mailNickname = Get-ArgValue $parsed.Map "mailNickname"
            if (-not $upn) {
                if (-not $alias) { $alias = $mailNickname }
                if (-not $alias -and $displayName) { $alias = ($displayName -replace "\\s", "") }
                if (-not $domain) { $domain = $global:Config.tenant.defaultDomain }
                if ($alias -and $domain) { $upn = ($alias + "@" + $domain) }
            }
            if (-not $mailNickname -and $upn) { $mailNickname = ($upn -split "@")[0] }

            $params = @{ AccountEnabled = $accountEnabled }
            if ($displayName) { $params.DisplayName = $displayName }
            if ($upn) { $params.UserPrincipalName = $upn }
            if ($mailNickname) { $params.MailNickname = $mailNickname }
            if ($password) {
                $params.PasswordProfile = @{
                    Password = $password
                    ForceChangePasswordNextSignIn = $forceChange
                }
            }
            if ($usageLocation) { $params.UsageLocation = $usageLocation }

            $extra = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
            if ($extra) {
                $extraMap = ConvertTo-Hashtable $extra
                foreach ($k in $extraMap.Keys) { $params[$k] = $extraMap[$k] }
            }

            if (-not $params.ContainsKey("DisplayName") -or -not $params.ContainsKey("UserPrincipalName") -or -not $params.ContainsKey("MailNickname") -or -not $params.ContainsKey("PasswordProfile")) {
                Write-Warn "Usage: user create --upn <upn> OR --alias <name> [--domain domain.com] --displayName <name> --password <pwd> [--usageLocation TR] [--forceChange true|false] [--enabled true|false] [--json <payload>]"
                return
            }
            try {
                $user = New-MgUser @params
                Write-Host ("Created user: " + $user.UserPrincipalName)
                Write-Host ("Id: " + $user.Id)
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "update" {
            $identity = $parsed.Positionals | Select-Object -First 1
            $setRaw = Get-ArgValue $parsed.Map "set"
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            $bodyFile = Get-ArgValue $parsed.Map "bodyFile"
            if (-not $identity -or (-not $setRaw -and -not $jsonRaw -and -not $bodyFile)) {
                Write-Warn 'Usage: user update <upn|id> --set key=value[,key=value] OR --json <payload> OR --bodyFile <path>'
                return
            }
            $body = Read-JsonPayload $jsonRaw $bodyFile $setRaw
            $body = ConvertTo-Hashtable $body
            if (-not $body -or $body.Keys.Count -eq 0) {
                Write-Warn "No properties to update."
                return
            }
            try {
                Update-MgUser -UserId $identity -BodyParameter $body | Out-Null
                Write-Info "User updated."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "delete" {
            $identity = $parsed.Positionals | Select-Object -First 1
            Invoke-UserDeleteOperation -Identity $identity -Map $parsed.Map
        }
        "bulkdelete" {
            Invoke-UserDeleteOperation -Identity $null -Map $parsed.Map
        }
        "apps" {
            $identity = $parsed.Positionals | Select-Object -First 1
            $explain = $parsed.Map.ContainsKey("explain")
            if (-not $identity) {
                Write-Warn "Usage: user apps <upn|id> [--explain]"
                return
            }
            try {
                $details = Get-MgUserLicenseDetail -UserId $identity -All
                if (-not $details) {
                    Write-Info "No license details found."
                    return
                }
                foreach ($d in $details) {
                    Write-Host ("SkuId: " + $d.SkuId)
                    if ($d.SkuPartNumber) { Write-Host ("SkuPartNumber: " + $d.SkuPartNumber) }
                    foreach ($p in $d.ServicePlans) {
                        Write-Host ("  " + $p.ServicePlanName + " : " + $p.ProvisioningStatus)
                        if ($explain) {
                            $info = Get-ServicePlanAutomationInfo $p.ServicePlanName
                            if ($info) {
                                Write-Host ("    Category: " + $info.Category)
                                Write-Host ("    Modules : " + ($info.Modules -join ", "))
                                Write-Host ("    Automation: " + ($info.Capabilities -join "; "))
                            }
                        }
                    }
                    Write-Host ""
                }
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "roles" {
            $identity = $parsed.Positionals | Select-Object -First 1
            if (-not $identity) {
                Write-Warn "Usage: user roles <upn|id>"
                return
            }
            try {
                $memberOf = Get-MgUserMemberOf -UserId $identity -All
                $roles = @()
                foreach ($item in $memberOf) {
                    $type = $null
                    if ($item.AdditionalProperties -and $item.AdditionalProperties['@odata.type']) {
                        $type = $item.AdditionalProperties['@odata.type']
                    } elseif ($item.OdataType) {
                        $type = $item.OdataType
                    }
                    if ($type -eq "#microsoft.graph.directoryRole") {
                        $roles += $item
                    }
                }
                if (-not $roles) {
                    Write-Info "No directory roles found."
                    return
                }
                foreach ($r in $roles) {
                    $role = Get-MgDirectoryRole -DirectoryRoleId $r.Id -ErrorAction SilentlyContinue
                    if ($role) {
                        Write-Host ("  " + $role.DisplayName + " (" + $role.Id + ")")
                    } else {
                        Write-Host ("  " + $r.Id)
                    }
                }
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "enable" {
            $identity = $parsed.Positionals | Select-Object -First 1
            if (-not $identity) {
                Write-Warn "Usage: user enable <upn|id>"
                return
            }
            try {
                Update-MgUser -UserId $identity -BodyParameter @{ accountEnabled = $true } | Out-Null
                Write-Info "User enabled."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "disable" {
            $identity = $parsed.Positionals | Select-Object -First 1
            if (-not $identity) {
                Write-Warn "Usage: user disable <upn|id>"
                return
            }
            try {
                Update-MgUser -UserId $identity -BodyParameter @{ accountEnabled = $false } | Out-Null
                Write-Info "User disabled."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "password" {
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action -or -not $identity) {
                Write-Warn "Usage: user password set|reset|change <upn|id> --password <pwd> [--forceChange true|false]"
                return
            }
            if ($action -notin @("set","reset","change")) {
                Write-Warn "Usage: user password set|reset|change <upn|id> --password <pwd> [--forceChange true|false]"
                return
            }
            $pwd = Get-ArgValue $parsed.Map "password"
            if (-not $pwd) { $pwd = Get-ArgValue $parsed.Map "pwd" }
            $forceChange = Parse-Bool (Get-ArgValue $parsed.Map "forceChange") $true
            if (-not $pwd) {
                Write-Warn "Password is required."
                return
            }
            $body = @{
                PasswordProfile = @{
                    Password = $pwd
                    ForceChangePasswordNextSignIn = $forceChange
                }
            }
            try {
                Update-MgUser -UserId $identity -BodyParameter $body | Out-Null
                Write-Info "Password updated."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "upn" {
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action -or -not $identity) {
                Write-Warn "Usage: user upn set <upn|id> --upn <newUpn> OR --alias <name> [--domain domain.com]"
                return
            }
            if ($action -notin @("set","change")) {
                Write-Warn "Usage: user upn set <upn|id> --upn <newUpn> OR --alias <name> [--domain domain.com]"
                return
            }
            $newUpn = Get-ArgValue $parsed.Map "upn"
            if (-not $newUpn) {
                $alias = Get-ArgValue $parsed.Map "alias"
                $domain = Get-ArgValue $parsed.Map "domain"
                if (-not $domain) { $domain = $global:Config.tenant.defaultDomain }
                if ($alias -and $domain) { $newUpn = ($alias + "@" + $domain) }
            }
            if (-not $newUpn) {
                Write-Warn "New UPN required."
                return
            }
            try {
                Update-MgUser -UserId $identity -BodyParameter @{ userPrincipalName = $newUpn } | Out-Null
                Write-Info ("UPN updated: " + $newUpn)
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "email" {
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action -or -not $identity) {
                Write-Warn "Usage: user email list|add|remove <upn|id> [--address <email>]"
                return
            }
            $user = $null
            try {
                $user = Get-MgUser -UserId $identity -Property Id,DisplayName,UserPrincipalName,Mail,OtherMails -ErrorAction Stop
            } catch {
                Write-Err $_.Exception.Message
                return
            }
            $address = Get-ArgValue $parsed.Map "address"
            if (-not $address) { $address = Get-ArgValue $parsed.Map "email" }
            if ($action -eq "list") {
                Write-Host ("Primary: " + $user.Mail)
                Write-Host "OtherMails:"
                foreach ($m in @($user.OtherMails)) { Write-Host ("  " + $m) }
                return
            }
            if (-not $address) {
                Write-Warn "Address required."
                return
            }
            $others = @($user.OtherMails)
            if ($action -eq "add") {
                if ($others -notcontains $address) { $others += $address }
            } elseif ($action -eq "remove") {
                $others = @($others | Where-Object { $_ -ne $address })
            } else {
                Write-Warn "Usage: user email list|add|remove <upn|id> [--address <email>]"
                return
            }
            try {
                Update-MgUser -UserId $user.Id -BodyParameter @{ otherMails = $others } | Out-Null
                Write-Info "Email addresses updated."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "alias" {
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action -or -not $identity) {
                Write-Warn "Usage: user alias list|add|remove <upn|id> [--address <alias@domain>] [--primary]"
                return
            }
            $user = $null
            try {
                $user = Get-MgUser -UserId $identity -Property Id,DisplayName,UserPrincipalName,ProxyAddresses -ErrorAction Stop
            } catch {
                Write-Err $_.Exception.Message
                return
            }
            if ($action -eq "list") {
                foreach ($p in @($user.ProxyAddresses)) { Write-Host ("  " + $p) }
                return
            }
            $address = Get-ArgValue $parsed.Map "address"
            if (-not $address) { $address = Get-ArgValue $parsed.Map "alias" }
            if (-not $address) {
                Write-Warn "Alias address required."
                return
            }
            $primary = Parse-Bool (Get-ArgValue $parsed.Map "primary") $false
            $proxyList = @($user.ProxyAddresses)
            if ($action -eq "add") {
                $normalized = $address
                if ($normalized -notmatch "^[sS][mM][tT][pP]:") {
                    $normalized = ($(if ($primary) { "SMTP:" } else { "smtp:" })) + $normalized
                } elseif ($primary -and $normalized -notmatch "^SMTP:") {
                    $normalized = "SMTP:" + ($normalized -replace "^[sS][mM][tT][pP]:", "")
                }
                if ($primary) {
                    $tmp = @()
                    foreach ($p in $proxyList) {
                        if ($p -match "^SMTP:") {
                            $tmp += ("smtp:" + $p.Substring(5))
                        } else {
                            $tmp += $p
                        }
                    }
                    $proxyList = $tmp
                }
                $exists = $false
                foreach ($p in $proxyList) {
                    if ($p.ToLowerInvariant() -eq $normalized.ToLowerInvariant()) { $exists = $true }
                }
                if (-not $exists) { $proxyList += $normalized }
            } elseif ($action -eq "remove") {
                $needle = $address
                if ($needle -match "^[sS][mM][tT][pP]:") {
                    $needle = $needle.Substring(5)
                }
                $tmp = @()
                foreach ($p in $proxyList) {
                    $cmp = $p
                    if ($cmp -match "^[sS][mM][tT][pP]:") { $cmp = $cmp.Substring(5) }
                    if ($cmp.ToLowerInvariant() -ne $needle.ToLowerInvariant()) {
                        $tmp += $p
                    }
                }
                $proxyList = $tmp
            } else {
                Write-Warn "Usage: user alias list|add|remove <upn|id> [--address <alias@domain>] [--primary]"
                return
            }
            try {
                Update-MgUser -UserId $user.Id -BodyParameter @{ proxyAddresses = $proxyList } | Out-Null
                Write-Info "Aliases updated."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "session" {
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action -or -not $identity) {
                Write-Warn "Usage: user session revoke <upn|id>"
                return
            }
            if ($action -ne "revoke") {
                Write-Warn "Usage: user session revoke <upn|id>"
                return
            }
            $user = Resolve-UserObject $identity
            if (-not $user) {
                Write-Warn "User not found."
                return
            }
            $resp = Invoke-GraphRequest -Method "POST" -Uri ("/users/" + $user.Id + "/revokeSignInSessions")
            if ($resp -ne $null) { Write-Info "Sessions revoked." }
        }
        "mfa" {
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action -or -not $identity) {
                Write-Warn "Usage: user mfa reset <upn|id> [--force]"
                return
            }
            if ($action -ne "reset") {
                Write-Warn "Usage: user mfa reset <upn|id> [--force]"
                return
            }
            $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
            if (-not $force) {
                $confirm = Read-Host "Type RESET to confirm MFA reset"
                if ($confirm -ne "RESET") {
                    Write-Info "Canceled."
                    return
                }
            }
            $user = Resolve-UserObject $identity
            if (-not $user) {
                Write-Warn "User not found."
                return
            }
            $resp = Invoke-GraphRequest -Method "GET" -Uri ("/users/" + $user.Id + "/authentication/methods")
            if (-not $resp -or -not $resp.value) {
                Write-Info "No auth methods found."
                return
            }
            foreach ($m in @($resp.value)) {
                $otype = $m.'@odata.type'
                if ($otype -and $otype.ToLowerInvariant().Contains("passwordauthenticationmethod")) { continue }
                try {
                    Invoke-GraphRequest -Method "DELETE" -Uri ("/users/" + $user.Id + "/authentication/methods/" + $m.id) | Out-Null
                    Write-Info ("Deleted method: " + $m.id)
                } catch {
                    Write-Err ("Failed to delete method: " + $m.id + " : " + $_.Exception.Message)
                }
            }
        }
        "license" {
            Handle-LicenseCommand $rest
        }
        default {
            Write-Warn "Unknown user subcommand. Use: user list|get|create|update|delete|bulkdelete|props|apps|roles|enable|disable|password|upn|email|alias|session|mfa|license"
        }
    }
}



function Handle-LicenseCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: license list|assign|remove|update"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "list" {
            try {
                $skus = Get-MgSubscribedSku -All
                $skus | Select-Object SkuPartNumber, SkuId, ConsumedUnits, CapabilityStatus | Format-Table -AutoSize
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "assign" {
            $identity = $parsed.Positionals | Select-Object -First 1
            $sku = Get-ArgValue $parsed.Map "sku"
            $disabled = Parse-GuidList (Get-ArgValue $parsed.Map "disablePlans")
            if (-not $identity -or -not $sku) {
                Write-Warn "Usage: license assign <upn|id> --sku <skuId> [--disablePlans planId,planId]"
                return
            }
            $user = Resolve-UserObject $identity
            if (-not $user) {
                Write-Warn "User not found."
                return
            }
            try {
                $add = @(@{ SkuId = [Guid]$sku; DisabledPlans = $disabled })
                Set-MgUserLicense -UserId $user.Id -AddLicenses $add -RemoveLicenses @() | Out-Null
                Write-Info "License assigned."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "remove" {
            $identity = $parsed.Positionals | Select-Object -First 1
            $sku = Get-ArgValue $parsed.Map "sku"
            if (-not $identity -or -not $sku) {
                Write-Warn "Usage: license remove <upn|id> --sku <skuId>"
                return
            }
            $user = Resolve-UserObject $identity
            if (-not $user) {
                Write-Warn "User not found."
                return
            }
            try {
                Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses @([Guid]$sku) | Out-Null
                Write-Info "License removed."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "update" {
            $identity = $parsed.Positionals | Select-Object -First 1
            $addRaw = Get-ArgValue $parsed.Map "add"
            $removeRaw = Get-ArgValue $parsed.Map "remove"
            $disabled = Parse-GuidList (Get-ArgValue $parsed.Map "disablePlans")
            if (-not $identity -or (-not $addRaw -and -not $removeRaw)) {
                Write-Warn "Usage: license update <upn|id> [--add skuId,skuId] [--remove skuId,skuId] [--disablePlans planId,planId]"
                return
            }
            $user = Resolve-UserObject $identity
            if (-not $user) {
                Write-Warn "User not found."
                return
            }
            $addList = @()
            foreach ($sku in (Parse-CommaList $addRaw)) {
                $addList += @{ SkuId = [Guid]$sku; DisabledPlans = $disabled }
            }
            $removeList = @()
            foreach ($sku in (Parse-CommaList $removeRaw)) {
                $removeList += [Guid]$sku
            }
            try {
                Set-MgUserLicense -UserId $user.Id -AddLicenses $addList -RemoveLicenses $removeList | Out-Null
                Write-Info "Licenses updated."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        default {
            Write-Warn "Unknown license subcommand. Use: license list|assign|remove|update"
        }
    }
}



function Handle-RoleCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: role list|assign|remove|templates|definitions|assignments"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $useBeta = $parsed.Map.ContainsKey("beta")
    $useV1 = $parsed.Map.ContainsKey("v1")
    $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
    if ($parsed.Map.ContainsKey("auto")) { $allowFallback = $true }
    if ($useBeta -or $useV1) { $allowFallback = $false }
    $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
    if ($sub -in @("template","templates")) { $sub = "templates" }
    if ($sub -in @("definition","definitions","def","defs")) { $sub = "definitions" }
    if ($sub -in @("assignment","assignments")) { $sub = "assignments" }

    switch ($sub) {
        "list" {
            try {
                Get-MgDirectoryRole -All | Select-Object Id, DisplayName, Description | Format-Table -AutoSize
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "templates" {
            try {
                Get-MgDirectoryRoleTemplate -All | Select-Object Id, DisplayName, Description | Format-Table -AutoSize
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "definitions" {
            $action = $parsed.Positionals | Select-Object -First 1
            if (-not $action) { $action = "list" }
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            switch ($action.ToLowerInvariant()) {
                "list" {
                    $qh = Build-QueryAndHeaders $parsed.Map @("id","displayName","description","isBuiltIn")
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ("/roleManagement/directory/roleDefinitions" + $qh.Query) -Headers $qh.Headers -Api $api -AllowFallback:$allowFallback
                    if ($resp -and $resp.value) {
                        Write-GraphTable $resp.value @("Id","DisplayName","Description","IsBuiltIn")
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 6
                    }
                }
                "get" {
                    if (-not $identity) {
                        Write-Warn "Usage: role definitions get <roleDefinitionId|name>"
                        return
                    }
                    $obj = Resolve-RoleDefinitionObject -Identity $identity -Api $api -AllowFallback:$allowFallback
                    if ($obj) { $obj | ConvertTo-Json -Depth 6 }
                }
                default {
                    Write-Warn "Usage: role definitions list|get"
                }
            }
        }
        "assignments" {
            $action = $parsed.Positionals | Select-Object -First 1
            if (-not $action) { $action = "list" }
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if ($action.ToLowerInvariant() -ne "list") {
                Write-Warn "Usage: role assignments list [--principal <id|upn>] [--definition <roleDefinitionId|name>] [--filter <odata>]"
                return
            }
            $principalId = Resolve-PrincipalIdFromMap -Map $parsed.Map -FallbackIdentity $identity
            $defRaw = Get-ArgValue $parsed.Map "definition"
            if (-not $defRaw) { $defRaw = Get-ArgValue $parsed.Map "roleDefinition" }
            $defId = $null
            if ($defRaw) { $defId = Resolve-RoleDefinitionId -Identity $defRaw -Api $api -AllowFallback:$allowFallback }
            $filter = Get-ArgValue $parsed.Map "filter"
            $parts = @()
            if ($principalId) { $parts += ("principalId eq '" + $principalId + "'") }
            if ($defId) { $parts += ("roleDefinitionId eq '" + $defId + "'") }
            if (-not $filter -and $parts.Count -gt 0) {
                $parsed.Map["filter"] = ($parts -join " and ")
            } elseif ($filter -and $parts.Count -gt 0) {
                $parsed.Map["filter"] = "(" + $filter + ") and " + ($parts -join " and ")
            }
            $qh = Build-QueryAndHeaders $parsed.Map @("id","principalId","roleDefinitionId","directoryScopeId")
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ("/roleManagement/directory/roleAssignments" + $qh.Query) -Headers $qh.Headers -Api $api -AllowFallback:$allowFallback
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Id","PrincipalId","RoleDefinitionId","DirectoryScopeId")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 6
            }
        }
        "assign" {
            $identity = $parsed.Positionals | Select-Object -First 1
            $roleName = Get-ArgValue $parsed.Map "role"
            $useRoleAssignment = $parsed.Map.ContainsKey("definition") -or $parsed.Map.ContainsKey("roleDefinition") -or $parsed.Map.ContainsKey("scope") -or $parsed.Map.ContainsKey("principalId")
            if ($useRoleAssignment) {
                $principalId = Resolve-PrincipalIdFromMap -Map $parsed.Map -FallbackIdentity $identity
                $defRaw = Get-ArgValue $parsed.Map "definition"
                if (-not $defRaw) { $defRaw = Get-ArgValue $parsed.Map "roleDefinition" }
                if (-not $defRaw) { $defRaw = Get-ArgValue $parsed.Map "role" }
                if (-not $principalId -or -not $defRaw) {
                    Write-Warn "Usage: role assign <principal> --definition <roleDefinitionId|name> [--scope /]"
                    return
                }
                $defId = Resolve-RoleDefinitionId -Identity $defRaw -Api $api -AllowFallback:$allowFallback
                if (-not $defId) {
                    Write-Warn "Role definition not found."
                    return
                }
                $scope = Get-ArgValue $parsed.Map "scope"
                if (-not $scope) { $scope = "/" }
                $body = @{ principalId = $principalId; roleDefinitionId = $defId; directoryScopeId = $scope }
                $resp = Invoke-GraphRequestAuto -Method "POST" -Uri "/roleManagement/directory/roleAssignments" -Body $body -Api $api -AllowFallback:$allowFallback
                if ($resp) { Write-Info ("Role assignment created: " + $resp.id) }
                return
            }

            if (-not $identity -or -not $roleName) {
                Write-Warn "Usage: role assign <upn|id> --role <roleName|roleId>"
                return
            }
            $principalId = Resolve-PrincipalIdFromMap -Map $parsed.Map -FallbackIdentity $identity
            if (-not $principalId) {
                Write-Warn "Principal not found."
                return
            }
            $role = Ensure-DirectoryRole $roleName
            if (-not $role) {
                Write-Warn "Role not found or could not be activated."
                return
            }
            try {
                $body = @{ "@odata.id" = ("https://graph.microsoft.com/v1.0/directoryObjects/" + $principalId) }
                New-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -BodyParameter $body | Out-Null
                Write-Info "Role assigned."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "remove" {
            $identity = $parsed.Positionals | Select-Object -First 1
            $roleName = Get-ArgValue $parsed.Map "role"
            $useRoleAssignment = $parsed.Map.ContainsKey("definition") -or $parsed.Map.ContainsKey("roleDefinition") -or $parsed.Map.ContainsKey("scope") -or $parsed.Map.ContainsKey("principalId") -or $parsed.Map.ContainsKey("assignment")
            if ($useRoleAssignment) {
                $assignmentId = Get-ArgValue $parsed.Map "assignment"
                if (-not $assignmentId) { $assignmentId = Get-ArgValue $parsed.Map "assignmentId" }
                if ($assignmentId) {
                    $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri ("/roleManagement/directory/roleAssignments/" + $assignmentId) -Api $api -AllowFallback:$allowFallback -AllowNullResponse
                    if ($resp -ne $null) { Write-Info "Role assignment removed." }
                    return
                }
                $principalId = Resolve-PrincipalIdFromMap -Map $parsed.Map -FallbackIdentity $identity
                $defRaw = Get-ArgValue $parsed.Map "definition"
                if (-not $defRaw) { $defRaw = Get-ArgValue $parsed.Map "roleDefinition" }
                if (-not $defRaw) { $defRaw = Get-ArgValue $parsed.Map "role" }
                if (-not $principalId -or -not $defRaw) {
                    Write-Warn "Usage: role remove <principal> --definition <roleDefinitionId|name> OR --assignment <assignmentId>"
                    return
                }
                $defId = Resolve-RoleDefinitionId -Identity $defRaw -Api $api -AllowFallback:$allowFallback
                if (-not $defId) {
                    Write-Warn "Role definition not found."
                    return
                }
                $filter = "principalId eq '$principalId' and roleDefinitionId eq '$defId'"
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ("/roleManagement/directory/roleAssignments?$filter=" + (Encode-QueryValue $filter)) -Api $api -AllowFallback:$allowFallback
                if ($resp -and $resp.value) {
                    foreach ($ra in @($resp.value)) {
                        Invoke-GraphRequestAuto -Method "DELETE" -Uri ("/roleManagement/directory/roleAssignments/" + $ra.Id) -Api $api -AllowFallback:$allowFallback -AllowNullResponse | Out-Null
                        Write-Info ("Role assignment removed: " + $ra.Id)
                    }
                } else {
                    Write-Info "No role assignments found."
                }
                return
            }

            if (-not $identity -or -not $roleName) {
                Write-Warn "Usage: role remove <upn|id> --role <roleName|roleId>"
                return
            }
            $principalId = Resolve-PrincipalIdFromMap -Map $parsed.Map -FallbackIdentity $identity
            if (-not $principalId) {
                Write-Warn "Principal not found."
                return
            }
            $role = Resolve-DirectoryRole $roleName
            if (-not $role) {
                Write-Warn "Role not found."
                return
            }
            try {
                Remove-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -DirectoryObjectId $principalId | Out-Null
                Write-Info "Role removed."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        default {
            Write-Warn "Unknown role subcommand. Use: role list|assign|remove|templates|definitions|assignments"
        }
    }
}



function Handle-GroupCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: group list|get|create|update|delete|bulkdelete|member|owner"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "list" {
            $filter = Get-ArgValue $parsed.Map "filter"
            $select = Get-ArgValue $parsed.Map "select"
            $props = if ($select) { Parse-CommaList $select } else { @("Id","DisplayName","Mail","MailEnabled","SecurityEnabled","GroupTypes","Visibility") }
            try {
                $groups = if ($filter) { Get-MgGroup -Filter $filter -All -Property $props } else { Get-MgGroup -All -Property $props }
                $groups | Select-Object -Property $props | Format-Table -AutoSize
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "get" {
            $identity = $parsed.Positionals | Select-Object -First 1
            if (-not $identity) {
                Write-Warn "Usage: group get <id|displayName>"
                return
            }
            $group = Resolve-GroupObject $identity
            if (-not $group) {
                Write-Warn "Group not found."
                return
            }
            $group | Format-List *
        }
        "create" {
            $name = Get-ArgValue $parsed.Map "name"
            if (-not $name) { $name = Get-ArgValue $parsed.Map "displayName" }
            $mailNickname = Get-ArgValue $parsed.Map "mailNickname"
            $description = Get-ArgValue $parsed.Map "description"
            $type = Get-ArgValue $parsed.Map "type"
            $visibility = Get-ArgValue $parsed.Map "visibility"

            $body = ConvertTo-Hashtable (Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set"))
            if ($name -and -not $body.ContainsKey("DisplayName")) { $body.DisplayName = $name }
            if (-not $mailNickname -and $body.ContainsKey("MailNickname")) { $mailNickname = $body.MailNickname }
            if (-not $mailNickname -and $name) { $mailNickname = ($name -replace '\\s','') }
            if ($mailNickname -and -not $body.ContainsKey("MailNickname")) { $body.MailNickname = $mailNickname }
            if ($description -and -not $body.ContainsKey("Description")) { $body.Description = $description }
            if ($visibility -and -not $body.ContainsKey("Visibility")) { $body.Visibility = $visibility }

            $hasType = $body.ContainsKey("GroupTypes") -or $body.ContainsKey("MailEnabled") -or $body.ContainsKey("SecurityEnabled")
            if (-not $hasType) {
                if (-not $type) { $type = "unified" }
                switch ($type.ToLowerInvariant()) {
                    "unified" { $body.GroupTypes = @("Unified"); $body.MailEnabled = $true;  $body.SecurityEnabled = $false }
                    "m365"    { $body.GroupTypes = @("Unified"); $body.MailEnabled = $true;  $body.SecurityEnabled = $false }
                    "security" { $body.GroupTypes = @(); $body.MailEnabled = $false; $body.SecurityEnabled = $true }
                    "mailsecurity" { $body.GroupTypes = @(); $body.MailEnabled = $true; $body.SecurityEnabled = $true }
                    "distribution" { $body.GroupTypes = @(); $body.MailEnabled = $true; $body.SecurityEnabled = $false }
                    default   { $body.GroupTypes = @("Unified"); $body.MailEnabled = $true;  $body.SecurityEnabled = $false }
                }
            }

            if (-not $body.ContainsKey("DisplayName")) {
                Write-Warn "Usage: group create --name <name> [--mailNickname nick] [--type unified|security|mailsecurity|distribution] [--description text] [--visibility Public|Private] [--json <payload>]"
                return
            }

            try {
                $group = New-MgGroup -BodyParameter $body
                Write-Host ("Created group: " + $group.DisplayName)
                Write-Host ("Id: " + $group.Id)
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "update" {
            $identity = $parsed.Positionals | Select-Object -First 1
            $setRaw = Get-ArgValue $parsed.Map "set"
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            $bodyFile = Get-ArgValue $parsed.Map "bodyFile"
            if (-not $identity -or (-not $setRaw -and -not $jsonRaw -and -not $bodyFile)) {
                Write-Warn 'Usage: group update <id> --set key=value[,key=value] OR --json <payload> OR --bodyFile <path>'
                return
            }
            $group = Resolve-GroupObject $identity
            if (-not $group) {
                Write-Warn "Group not found."
                return
            }
            $body = ConvertTo-Hashtable (Read-JsonPayload $jsonRaw $bodyFile $setRaw)
            if (-not $body -or $body.Keys.Count -eq 0) {
                Write-Warn "No properties to update."
                return
            }
            try {
                Update-MgGroup -GroupId $group.Id -BodyParameter $body | Out-Null
                Write-Info "Group updated."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "delete" {
            $identity = $parsed.Positionals | Select-Object -First 1
            Invoke-GroupDeleteOperation -Identity $identity -Map $parsed.Map
        }
        "bulkdelete" {
            Invoke-GroupDeleteOperation -Identity $null -Map $parsed.Map
        }
        "member" {
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action -or -not $identity) {
                Write-Warn "Usage: group member list|add|remove <groupId|name> [--member <upn|id>] [--objectId <id>]"
                return
            }
            $group = Resolve-GroupObject $identity
            if (-not $group) {
                Write-Warn "Group not found."
                return
            }
            switch ($action) {
                "list" {
                    try {
                        $members = Get-MgGroupMember -GroupId $group.Id -All
                        $rows = @()
                        foreach ($m in @($members)) {
                            $props = $m.AdditionalProperties
                            $rows += [pscustomobject]@{
                                Id = $m.Id
                                Type = $props['@odata.type']
                                DisplayName = $props['displayName']
                                UserPrincipalName = $props['userPrincipalName']
                            }
                        }
                        $rows | Format-Table -AutoSize
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "add" {
                    $objectId = Get-ArgValue $parsed.Map "objectId"
                    $objectIdsRaw = Get-ArgValue $parsed.Map "objectIds"
                    $member = Get-ArgValue $parsed.Map "member"
                    $membersRaw = Get-ArgValue $parsed.Map "members"
                    $targets = @()
                    if ($objectId) { $targets += $objectId }
                    foreach ($id in (Parse-CommaList $objectIdsRaw)) { $targets += $id }
                    if ($member) {
                        $u = Resolve-UserObject $member
                        if ($u) { $targets += $u.Id }
                    }
                    foreach ($m in (Parse-CommaList $membersRaw)) {
                        $u = Resolve-UserObject $m
                        if ($u) { $targets += $u.Id }
                    }
                    $targets = @($targets | Select-Object -Unique)
                    if (-not $targets -or $targets.Count -eq 0) {
                        Write-Warn "Usage: group member add <groupId|name> --member <upn|id> [--members a,b] OR --objectId <id> [--objectIds id1,id2]"
                        return
                    }
                    foreach ($tid in $targets) {
                        try {
                            $body = @{ "@odata.id" = ("https://graph.microsoft.com/v1.0/directoryObjects/" + $tid) }
                            New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter $body | Out-Null
                            Write-Info ("Member added: " + $tid)
                        } catch {
                            Write-Err ("Failed to add member " + $tid + " : " + $_.Exception.Message)
                        }
                    }
                }
                "remove" {
                    $objectId = Get-ArgValue $parsed.Map "objectId"
                    $objectIdsRaw = Get-ArgValue $parsed.Map "objectIds"
                    $member = Get-ArgValue $parsed.Map "member"
                    $membersRaw = Get-ArgValue $parsed.Map "members"
                    $targets = @()
                    if ($objectId) { $targets += $objectId }
                    foreach ($id in (Parse-CommaList $objectIdsRaw)) { $targets += $id }
                    if ($member) {
                        $u = Resolve-UserObject $member
                        if ($u) { $targets += $u.Id }
                    }
                    foreach ($m in (Parse-CommaList $membersRaw)) {
                        $u = Resolve-UserObject $m
                        if ($u) { $targets += $u.Id }
                    }
                    $targets = @($targets | Select-Object -Unique)
                    if (-not $targets -or $targets.Count -eq 0) {
                        Write-Warn "Usage: group member remove <groupId|name> --member <upn|id> [--members a,b] OR --objectId <id> [--objectIds id1,id2]"
                        return
                    }
                    foreach ($tid in $targets) {
                        try {
                            Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $tid | Out-Null
                            Write-Info ("Member removed: " + $tid)
                        } catch {
                            Write-Err ("Failed to remove member " + $tid + " : " + $_.Exception.Message)
                        }
                    }
                }
                default {
                    Write-Warn "Usage: group member list|add|remove <groupId|name> [--member <upn|id>] [--objectId <id>]"
                }
            }
        }
        "owner" {
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action -or -not $identity) {
                Write-Warn "Usage: group owner list|add|remove <groupId|name> [--owner <upn|id>] [--objectId <id>]"
                return
            }
            $group = Resolve-GroupObject $identity
            if (-not $group) {
                Write-Warn "Group not found."
                return
            }
            switch ($action) {
                "list" {
                    try {
                        $owners = Get-MgGroupOwner -GroupId $group.Id -All
                        $rows = @()
                        foreach ($o in @($owners)) {
                            $props = $o.AdditionalProperties
                            $rows += [pscustomobject]@{
                                Id = $o.Id
                                Type = $props['@odata.type']
                                DisplayName = $props['displayName']
                                UserPrincipalName = $props['userPrincipalName']
                            }
                        }
                        $rows | Format-Table -AutoSize
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "add" {
                    $objectId = Get-ArgValue $parsed.Map "objectId"
                    $objectIdsRaw = Get-ArgValue $parsed.Map "objectIds"
                    $owner = Get-ArgValue $parsed.Map "owner"
                    $ownersRaw = Get-ArgValue $parsed.Map "owners"
                    $targets = @()
                    if ($objectId) { $targets += $objectId }
                    foreach ($id in (Parse-CommaList $objectIdsRaw)) { $targets += $id }
                    if ($owner) {
                        $u = Resolve-UserObject $owner
                        if ($u) { $targets += $u.Id }
                    }
                    foreach ($m in (Parse-CommaList $ownersRaw)) {
                        $u = Resolve-UserObject $m
                        if ($u) { $targets += $u.Id }
                    }
                    $targets = @($targets | Select-Object -Unique)
                    if (-not $targets -or $targets.Count -eq 0) {
                        Write-Warn "Usage: group owner add <groupId|name> --owner <upn|id> [--owners a,b] OR --objectId <id> [--objectIds id1,id2]"
                        return
                    }
                    foreach ($tid in $targets) {
                        try {
                            $body = @{ "@odata.id" = ("https://graph.microsoft.com/v1.0/directoryObjects/" + $tid) }
                            New-MgGroupOwnerByRef -GroupId $group.Id -BodyParameter $body | Out-Null
                            Write-Info ("Owner added: " + $tid)
                        } catch {
                            Write-Err ("Failed to add owner " + $tid + " : " + $_.Exception.Message)
                        }
                    }
                }
                "remove" {
                    $objectId = Get-ArgValue $parsed.Map "objectId"
                    $objectIdsRaw = Get-ArgValue $parsed.Map "objectIds"
                    $owner = Get-ArgValue $parsed.Map "owner"
                    $ownersRaw = Get-ArgValue $parsed.Map "owners"
                    $targets = @()
                    if ($objectId) { $targets += $objectId }
                    foreach ($id in (Parse-CommaList $objectIdsRaw)) { $targets += $id }
                    if ($owner) {
                        $u = Resolve-UserObject $owner
                        if ($u) { $targets += $u.Id }
                    }
                    foreach ($m in (Parse-CommaList $ownersRaw)) {
                        $u = Resolve-UserObject $m
                        if ($u) { $targets += $u.Id }
                    }
                    $targets = @($targets | Select-Object -Unique)
                    if (-not $targets -or $targets.Count -eq 0) {
                        Write-Warn "Usage: group owner remove <groupId|name> --owner <upn|id> [--owners a,b] OR --objectId <id> [--objectIds id1,id2]"
                        return
                    }
                    foreach ($tid in $targets) {
                        try {
                            Remove-MgGroupOwnerByRef -GroupId $group.Id -DirectoryObjectId $tid | Out-Null
                            Write-Info ("Owner removed: " + $tid)
                        } catch {
                            Write-Err ("Failed to remove owner " + $tid + " : " + $_.Exception.Message)
                        }
                    }
                }
                default {
                    Write-Warn "Usage: group owner list|add|remove <groupId|name> [--owner <upn|id>] [--objectId <id>]"
                }
            }
        }
        default {
            Write-Warn "Unknown group subcommand. Use: group list|get|create|update|delete|bulkdelete|member|owner"
        }
    }
}



function Handle-DomainCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: domain list|get|add|verify|default|dns"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "list" {
            try {
                Get-MgDomain -All | Select-Object Id, IsDefault, IsVerified, AuthenticationType | Format-Table -AutoSize
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "get" {
            $domain = $parsed.Positionals | Select-Object -First 1
            if (-not $domain) {
                Write-Warn "Usage: domain get <domain>"
                return
            }
            try {
                Get-MgDomain -DomainId $domain | Format-List *
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "add" {
            $domain = $parsed.Positionals | Select-Object -First 1
            if (-not $domain) {
                Write-Warn "Usage: domain add <domain>"
                return
            }
            try {
                $body = @{ Id = $domain }
                New-MgDomain -BodyParameter $body | Out-Null
                Write-Info "Domain added. Use: domain dns <domain> then domain verify <domain>"
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "verify" {
            $domain = $parsed.Positionals | Select-Object -First 1
            if (-not $domain) {
                Write-Warn "Usage: domain verify <domain>"
                return
            }
            try {
                Confirm-MgDomain -DomainId $domain | Out-Null
                Write-Info "Domain verification requested."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "default" {
            $domain = $parsed.Positionals | Select-Object -First 1
            if (-not $domain) {
                Write-Warn "Usage: domain default <domain>"
                return
            }
            try {
                Update-MgDomain -DomainId $domain -BodyParameter @{ IsDefault = $true } | Out-Null
                Write-Info "Domain set as default."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "dns" {
            $domain = $parsed.Positionals | Select-Object -First 1
            if (-not $domain) {
                Write-Warn "Usage: domain dns <domain>"
                return
            }
            try {
                Write-Host "Verification records:"
                $v = Get-MgDomainVerificationDnsRecord -DomainId $domain -All
                if ($v) { $v | Select-Object RecordType, Label, SupportedService, Ttl, AdditionalProperties | Format-Table -AutoSize }
                Write-Host "Service configuration records:"
                $s = Get-MgDomainServiceConfigurationRecord -DomainId $domain -All
                if ($s) { $s | Select-Object RecordType, Label, Ttl, AdditionalProperties | Format-Table -AutoSize }
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        default {
            Write-Warn "Unknown domain subcommand. Use: domain list|get|add|verify|default|dns"
        }
    }
}



function Handle-AppCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: app list|get|create|update|delete|redirect|secret|cert|perm|consent|guide"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "list" {
            try {
                Get-MgApplication -All -Property Id, AppId, DisplayName, SignInAudience |
                    Select-Object Id, AppId, DisplayName, SignInAudience | Format-Table -AutoSize
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "get" {
            $identity = $parsed.Positionals | Select-Object -First 1
            if (-not $identity) {
                Write-Warn "Usage: app get <appId|objectId>"
                return
            }
            $app = Resolve-ApplicationObject $identity
            if (-not $app) {
                Write-Warn "Application not found."
                return
            }
            $app | ConvertTo-Json -Depth 6
        }
        "create" {
            $name = Get-ArgValue $parsed.Map "name"
            $type = (Get-ArgValue $parsed.Map "type")
            $redirect = Parse-CommaList (Get-ArgValue $parsed.Map "redirect")
            $audience = (Get-ArgValue $parsed.Map "audience")
            if (-not $name) {
                Write-Warn "Usage: app create --name <name> [--type spa|web|native|daemon|mobile|console] [--redirect uri,uri] [--audience AzureADMyOrg]"
                return
            }
            if (-not $type) { $type = "web" }
            if (-not $audience) { $audience = "AzureADMyOrg" }

            $body = @{
                DisplayName    = $name
                SignInAudience = $audience
            }
            switch ($type.ToLowerInvariant()) {
                "spa"     { $body.Spa = @{ RedirectUris = $redirect } }
                "web"     { $body.Web = @{ RedirectUris = $redirect } }
                "native"  { $body.PublicClient = @{ RedirectUris = $redirect } }
                "mobile"  { $body.PublicClient = @{ RedirectUris = $redirect } }
                "console" { $body.PublicClient = @{ RedirectUris = $redirect } }
                "daemon"  { $body.Web = @{ RedirectUris = $redirect } }
                default   { $body.Web = @{ RedirectUris = $redirect } }
            }
            try {
                $app = New-MgApplication -BodyParameter $body
                Write-Host ("Created app: " + $app.DisplayName)
                Write-Host ("AppId: " + $app.AppId)
                Write-Host ("ObjectId: " + $app.Id)
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "update" {
            $identity = $parsed.Positionals | Select-Object -First 1
            $setRaw = Get-ArgValue $parsed.Map "set"
            if (-not $identity -or -not $setRaw) {
                Write-Warn 'Usage: app update <appId|objectId> --set key=value[,key=value] OR --set ''{"web":{"redirectUris":["https://..."]}}'''
                return
            }
            $app = Resolve-ApplicationObject $identity
            if (-not $app) {
                Write-Warn "Application not found."
                return
            }
            $body = Parse-Value $setRaw
            if ($body -is [string]) { $body = Parse-KvPairs $setRaw }
            if ($body.Keys.Count -eq 0) {
                Write-Warn "No properties to update."
                return
            }
            try {
                Update-MgApplication -ApplicationId $app.Id -BodyParameter $body | Out-Null
                Write-Info "Application updated."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "delete" {
            $identity = $parsed.Positionals | Select-Object -First 1
            if (-not $identity) {
                Write-Warn "Usage: app delete <appId|objectId> [--force]"
                return
            }
            $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
            $app = Resolve-ApplicationObject $identity
            if (-not $app) {
                Write-Warn "Application not found."
                return
            }
            if (-not $force) {
                $confirm = Read-Host "Type DELETE to confirm"
                if ($confirm -ne "DELETE") {
                    Write-Info "Canceled."
                    return
                }
            }
            try {
                Remove-MgApplication -ApplicationId $app.Id -ErrorAction Stop
                Write-Info "Application deleted."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "redirect" {
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            $uri = Get-ArgValue $parsed.Map "uri"
            $type = Get-ArgValue $parsed.Map "type"
            if (-not $action -or -not $identity -or -not $uri) {
                Write-Warn "Usage: app redirect add|remove <appId|objectId> --uri <url> [--type spa|web|public]"
                return
            }
            $app = Resolve-ApplicationObject $identity
            if (-not $app) {
                Write-Warn "Application not found."
                return
            }
            if (-not $type) {
                if ($app.Spa) { $type = "spa" }
                elseif ($app.Web) { $type = "web" }
                elseif ($app.PublicClient) { $type = "public" }
                else { $type = "web" }
            }
            $propName = switch ($type.ToLowerInvariant()) {
                "spa"    { "Spa" }
                "public" { "PublicClient" }
                default  { "Web" }
            }
            $current = @()
            if ($app.$propName -and $app.$propName.RedirectUris) {
                $current = @($app.$propName.RedirectUris)
            }
            if ($action -eq "add") {
                if ($current -notcontains $uri) { $current += $uri }
            } elseif ($action -eq "remove") {
                $current = @($current | Where-Object { $_ -ne $uri })
            } else {
                Write-Warn "Usage: app redirect add|remove <appId|objectId> --uri <url> [--type spa|web|public]"
                return
            }
            try {
                $body = @{}
                $body[$propName] = @{ RedirectUris = $current }
                Update-MgApplication -ApplicationId $app.Id -BodyParameter $body | Out-Null
                Write-Info "Redirect URIs updated."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "secret" {
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action -or -not $identity) {
                Write-Warn "Usage: app secret list|add|remove <appId|objectId> [--months 12] [--displayName name] [--keyId <guid>]"
                return
            }
            $app = Resolve-ApplicationObject $identity
            if (-not $app) {
                Write-Warn "Application not found."
                return
            }
            switch ($action) {
                "list" {
                    $app.PasswordCredentials | Select-Object KeyId, DisplayName, StartDateTime, EndDateTime | Format-Table -AutoSize
                }
                "add" {
                    $months = Get-ArgValue $parsed.Map "months"
                    if (-not $months) { $months = 12 }
                    $displayName = Get-ArgValue $parsed.Map "displayName"
                    if (-not $displayName) { $displayName = "m365cmd" }
                    $end = (Get-Date).AddMonths([int]$months)
                    $cred = @{
                        DisplayName = $displayName
                        EndDateTime = $end
                    }
                    try {
                        $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential $cred
                        Write-Host ("SecretText: " + $secret.SecretText)
                        Write-Host ("KeyId: " + $secret.KeyId)
                        Write-Info "Save the secret now. It is only shown once."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "remove" {
                    $keyId = Get-ArgValue $parsed.Map "keyId"
                    if (-not $keyId) {
                        Write-Warn "Usage: app secret remove <appId|objectId> --keyId <guid>"
                        return
                    }
                    try {
                        Remove-MgApplicationPassword -ApplicationId $app.Id -KeyId $keyId | Out-Null
                        Write-Info "Secret removed."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                default {
                    Write-Warn "Usage: app secret list|add|remove <appId|objectId>"
                }
            }
        }
        "cert" {
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action -or -not $identity) {
                Write-Warn "Usage: app cert list|add|remove <appId|objectId> [--path cert.cer] [--months 12] [--displayName name] [--keyId <guid>]"
                return
            }
            $app = Resolve-ApplicationObject $identity
            if (-not $app) {
                Write-Warn "Application not found."
                return
            }
            switch ($action) {
                "list" {
                    $app.KeyCredentials | Select-Object KeyId, DisplayName, StartDateTime, EndDateTime, Type, Usage | Format-Table -AutoSize
                }
                "add" {
                    $path = Get-ArgValue $parsed.Map "path"
                    if (-not $path -or -not (Test-Path $path)) {
                        Write-Warn "Usage: app cert add <appId|objectId> --path <cert.cer> [--months 12] [--displayName name]"
                        return
                    }
                    $months = Get-ArgValue $parsed.Map "months"
                    if (-not $months) { $months = 12 }
                    $displayName = Get-ArgValue $parsed.Map "displayName"
                    if (-not $displayName) { $displayName = "m365cmd-cert" }
                    try {
                        $rawBytes = [System.IO.File]::ReadAllBytes($path)
                        $base64 = [System.Convert]::ToBase64String($rawBytes)
                        $keyBytes = [System.Text.Encoding]::ASCII.GetBytes($base64)
                        $newKey = @{
                            Type          = "AsymmetricX509Cert"
                            Usage         = "Verify"
                            Key           = $keyBytes
                            DisplayName   = $displayName
                            StartDateTime = Get-Date
                            EndDateTime   = (Get-Date).AddMonths([int]$months)
                        }
                        $existing = @()
                        foreach ($k in @($app.KeyCredentials)) {
                            $existing += @{
                                Type          = $k.Type
                                Usage         = $k.Usage
                                Key           = $k.Key
                                KeyId         = $k.KeyId
                                DisplayName   = $k.DisplayName
                                StartDateTime = $k.StartDateTime
                                EndDateTime   = $k.EndDateTime
                                CustomKeyIdentifier = $k.CustomKeyIdentifier
                            }
                        }
                        $existing += $newKey
                        Update-MgApplication -ApplicationId $app.Id -BodyParameter @{ KeyCredentials = $existing } | Out-Null
                        Write-Info "Certificate added."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                "remove" {
                    $keyId = Get-ArgValue $parsed.Map "keyId"
                    if (-not $keyId) {
                        Write-Warn "Usage: app cert remove <appId|objectId> --keyId <guid>"
                        return
                    }
                    try {
                        $remaining = @()
                        foreach ($k in @($app.KeyCredentials)) {
                            if ($k.KeyId -ne $keyId) {
                                $remaining += @{
                                    Type          = $k.Type
                                    Usage         = $k.Usage
                                    Key           = $k.Key
                                    KeyId         = $k.KeyId
                                    DisplayName   = $k.DisplayName
                                    StartDateTime = $k.StartDateTime
                                    EndDateTime   = $k.EndDateTime
                                    CustomKeyIdentifier = $k.CustomKeyIdentifier
                                }
                            }
                        }
                        Update-MgApplication -ApplicationId $app.Id -BodyParameter @{ KeyCredentials = $remaining } | Out-Null
                        Write-Info "Certificate removed."
                    } catch {
                        Write-Err $_.Exception.Message
                    }
                }
                default {
                    Write-Warn "Usage: app cert list|add|remove <appId|objectId>"
                }
            }
        }
        "perm" {
            $action = $parsed.Positionals | Select-Object -First 1
            $identity = $parsed.Positionals | Select-Object -Skip 1 -First 1
            if (-not $action -or -not $identity) {
                Write-Warn "Usage: app perm list|add|remove <appId|objectId> [--scope <name>] [--type delegated|application] [--all]"
                return
            }
            $app = Resolve-ApplicationObject $identity
            if (-not $app) {
                Write-Warn "Application not found."
                return
            }
            $graphSp = Get-GraphServicePrincipal
            if (-not $graphSp) {
                Write-Warn "Microsoft Graph service principal not found."
                return
            }
            if ($action -eq "list") {
                if ($parsed.Map.ContainsKey("all")) {
                    Write-Host "Delegated permissions:"
                    $graphSp.Oauth2PermissionScopes | ForEach-Object { Write-Host ("  Scope : " + $_.Value + " (" + $_.Id + ")") }
                    Write-Host "Application permissions:"
                    $graphSp.AppRoles | ForEach-Object { Write-Host ("  Role  : " + $_.Value + " (" + $_.Id + ")") }
                    return
                }
                $req = @($app.RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq $graphSp.AppId } | Select-Object -First 1)
                if (-not $req) {
                    Write-Info "No Microsoft Graph permissions assigned."
                    return
                }
                foreach ($ra in @($req.ResourceAccess)) {
                    if ($ra.Type -eq "Scope") {
                        $perm = $graphSp.Oauth2PermissionScopes | Where-Object { $_.Id -eq $ra.Id } | Select-Object -First 1
                    } else {
                        $perm = $graphSp.AppRoles | Where-Object { $_.Id -eq $ra.Id } | Select-Object -First 1
                    }
                    $name = if ($perm) { $perm.Value } else { "<unknown>" }
                    Write-Host ("  " + $ra.Type + " : " + $name + " (" + $ra.Id + ")")
                }
                return
            }

            $permName = Get-ArgValue $parsed.Map "scope"
            $type = Get-ArgValue $parsed.Map "type"
            if (-not $permName) {
                Write-Warn "Usage: app perm add|remove <appId|objectId> --scope <name> [--type delegated|application]"
                return
            }
            $perm = Resolve-GraphPermission $graphSp $permName $type
            if (-not $perm) {
                Write-Warn "Permission not found."
                return
            }

            $reqList = Convert-RequiredResourceAccessList $app.RequiredResourceAccess
            $entry = $reqList | Where-Object { $_.ResourceAppId -eq $graphSp.AppId } | Select-Object -First 1
            if (-not $entry) {
                $entry = @{ ResourceAppId = $graphSp.AppId; ResourceAccess = @() }
                $reqList += $entry
            }
            if ($action -eq "add") {
                $exists = $entry.ResourceAccess | Where-Object { $_.Id -eq $perm.Id -and $_.Type -eq $perm.Type }
                if (-not $exists) {
                    $entry.ResourceAccess += @{ Id = $perm.Id; Type = $perm.Type }
                }
            } elseif ($action -eq "remove") {
                $entry.ResourceAccess = @($entry.ResourceAccess | Where-Object { $_.Id -ne $perm.Id -or $_.Type -ne $perm.Type })
                if ($entry.ResourceAccess.Count -eq 0) {
                    $reqList = @($reqList | Where-Object { $_.ResourceAppId -ne $graphSp.AppId })
                }
            } else {
                Write-Warn "Usage: app perm add|remove <appId|objectId> --scope <name> [--type delegated|application]"
                return
            }
            try {
                Update-MgApplication -ApplicationId $app.Id -BodyParameter @{ RequiredResourceAccess = $reqList } | Out-Null
                Write-Info "Permissions updated."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "consent" {
            Write-Warn "Admin consent automation is not implemented yet. Use the Azure portal or Graph API."
        }
        "guide" {
            $identity = $parsed.Positionals | Select-Object -First 1
            $target = Get-ArgValue $parsed.Map "target"
            if (-not $identity) {
                Write-Warn "Usage: app guide <appId|objectId> [--target react|node|dotnet|spa|daemon]"
                return
            }
            $app = Resolve-ApplicationObject $identity
            if (-not $app) {
                Write-Warn "Application not found."
                return
            }
            if (-not $target) { $target = "react" }

            $tenantId = $global:Config.tenant.tenantId
            if (-not $tenantId) {
                $ctx = Get-MgContextSafe
                if ($ctx -and $ctx.TenantId) { $tenantId = $ctx.TenantId }
            }
            $tenantHint = if ($tenantId) { $tenantId } else { $global:Config.tenant.defaultDomain }
            $authority = "https://login.microsoftonline.com/" + $tenantHint

            Write-Host ("App: " + $app.DisplayName)
            Write-Host ("ClientId: " + $app.AppId)
            Write-Host ("Tenant: " + $tenantHint)
            Write-Host ("Authority: " + $authority)
            Write-Host ""
            switch ($target.ToLowerInvariant()) {
                "node" {
                    Write-Host "Node (MSAL) sample:"
                    Write-Host ("  clientId  : " + $app.AppId)
                    Write-Host ("  authority : " + $authority)
                    Write-Host ("  redirect  : " + (($app.Web.RedirectUris | Select-Object -First 1) -as [string]))
                }
                "dotnet" {
                    Write-Host "Dotnet sample:"
                    Write-Host ("  Instance : https://login.microsoftonline.com/")
                    Write-Host ("  TenantId : " + $tenantHint)
                    Write-Host ("  ClientId : " + $app.AppId)
                }
                "spa" {
                    Write-Host "SPA sample:"
                    Write-Host ("  clientId  : " + $app.AppId)
                    Write-Host ("  authority : " + $authority)
                    Write-Host ("  redirect  : " + (($app.Spa.RedirectUris | Select-Object -First 1) -as [string]))
                }
                "daemon" {
                    Write-Host "Daemon (client credentials) sample:"
                    Write-Host ("  clientId  : " + $app.AppId)
                    Write-Host ("  tenant    : " + $tenantHint)
                    Write-Host ("  authority : " + $authority)
                    Write-Host ("  scope     : https://graph.microsoft.com/.default")
                }
                default {
                    Write-Host "React (MSAL) sample:"
                    Write-Host ("  clientId  : " + $app.AppId)
                    Write-Host ("  authority : " + $authority)
                    Write-Host ("  redirect  : " + (($app.Spa.RedirectUris | Select-Object -First 1) -as [string]))
                }
            }
        }
        default {
            Write-Warn "Unknown app subcommand. Use: app list|get|create|update|delete|redirect|secret|cert|perm|consent|guide"
        }
    }
}



function Handle-OrgCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: org list|get|update"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    $useBeta = $parsed.Map.ContainsKey("beta")
    $useV1 = $parsed.Map.ContainsKey("v1")
    $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
    if ($parsed.Map.ContainsKey("auto")) { $allowFallback = $true }
    if ($useBeta -or $useV1) { $allowFallback = $false }
    $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }

    switch ($sub) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @("id", "displayName")
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ("/organization" + $qh.Query) -Headers $qh.Headers -Api $api -AllowFallback:$allowFallback
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Id", "DisplayName")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 6
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if ($id) {
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ("/organization/" + $id) -Api $api -AllowFallback:$allowFallback
            } else {
                $resp = Get-OrganizationDefault
            }
            if ($resp) { $resp | ConvertTo-Json -Depth 8 }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $setRaw = Get-ArgValue $parsed.Map "set"
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            if (-not $setRaw -and -not $jsonRaw) {
                Write-Warn "Usage: org update [orgId] --set key=value[,key=value] OR --json <payload>"
                return
            }
            $body = $null
            if ($jsonRaw) {
                $body = Parse-Value $jsonRaw
            } else {
                $body = Parse-KvPairs $setRaw
            }
            if (-not $body -or $body.Keys.Count -eq 0) {
                Write-Warn "No properties to update."
                return
            }
            if (-not $id) {
                $org = Get-OrganizationDefault
                if (-not $org -or -not $org.Id) {
                    Write-Warn "Organization not found."
                    return
                }
                $id = $org.Id
            }
            $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ("/organization/" + $id) -Body $body -Api $api -AllowFallback:$allowFallback
            if ($resp) { Write-Info "Organization updated." }
        }
        default {
            Write-Warn "Usage: org list|get|update"
        }
    }
}



function Handle-DirSettingCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: dirsetting list|get|create|update|delete|template"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }

    if ($sub -in @("template", "templates")) {
        $action = if ($rest.Count -gt 0) { $rest[0].ToLowerInvariant() } else { "list" }
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed = Parse-NamedArgs $rest2
        $useBeta = $parsed.Map.ContainsKey("beta")
        $useV1 = $parsed.Map.ContainsKey("v1")
        $allowFallback = $false
        $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
        $base = if ($useBeta) { "/directorySettingTemplates" } else { "/groupSettingTemplates" }

        switch ($action) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed.Map @("id", "displayName")
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $api -AllowFallback:$allowFallback
                if ($resp -and $resp.value) {
                    Write-GraphTable $resp.value @("Id", "DisplayName")
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 6
                }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: dirsetting template get <templateId> [--beta]"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 6 }
            }
            default {
                Write-Warn "Usage: dirsetting template list|get [--beta]"
            }
        }
        return
    }

    $parsed = Parse-NamedArgs $rest
    $useBeta = $parsed.Map.ContainsKey("beta")
    $useV1 = $parsed.Map.ContainsKey("v1")
    $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
    if ($parsed.Map.ContainsKey("auto")) { $allowFallback = $true }
    if ($useBeta -or $useV1) { $allowFallback = $false }
    $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
    $groupId = Get-ArgValue $parsed.Map "group"
    $base = if ($groupId) { "/groups/" + $groupId + "/settings" } else { "/settings" }

    switch ($sub) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @("id", "displayName")
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $api -AllowFallback:$allowFallback
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Id", "DisplayName")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 6
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: dirsetting get <settingId> [--group <groupId>]"
                return
            }
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
            if ($resp) { $resp | ConvertTo-Json -Depth 6 }
        }
        "create" {
            $templateId = Get-ArgValue $parsed.Map "templateId"
            $valuesRaw = Get-ArgValue $parsed.Map "values"
            $setRaw = Get-ArgValue $parsed.Map "set"
            if (-not $templateId) {
                Write-Warn "Usage: dirsetting create --templateId <id> [--set key=value,...] [--values <json>] [--group <groupId>]"
                return
            }
            $values = @()
            if ($valuesRaw) {
                $vals = Parse-Value $valuesRaw
                if ($vals -is [hashtable] -and $vals.values) {
                    $values = $vals.values
                } elseif ($vals) {
                    $values = $vals
                }
            } elseif ($setRaw) {
                $pairs = Parse-KvPairs $setRaw
                $values = Convert-SettingValuesFromPairs $pairs
            }
            $body = @{ templateId = $templateId }
            if ($values -and $values.Count -gt 0) { $body.values = $values }
            $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $api -AllowFallback:$allowFallback
            if ($resp) { Write-Info "Directory setting created." }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $valuesRaw = Get-ArgValue $parsed.Map "values"
            $setRaw = Get-ArgValue $parsed.Map "set"
            if (-not $id -or (-not $valuesRaw -and -not $setRaw)) {
                Write-Warn "Usage: dirsetting update <settingId> --set key=value,... OR --values <json> [--group <groupId>]"
                return
            }
            $values = @()
            if ($valuesRaw) {
                $vals = Parse-Value $valuesRaw
                if ($vals -is [hashtable] -and $vals.values) {
                    $values = $vals.values
                } elseif ($vals) {
                    $values = $vals
                }
            } else {
                $pairs = Parse-KvPairs $setRaw
                $values = Convert-SettingValuesFromPairs $pairs
            }
            if (-not $values -or $values.Count -eq 0) {
                Write-Warn "No values to update."
                return
            }
            $body = @{ values = $values }
            $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $api -AllowFallback:$allowFallback
            if ($resp) { Write-Info "Directory setting updated." }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
            if (-not $id) {
                Write-Warn "Usage: dirsetting delete <settingId> [--force] [--group <groupId>]"
                return
            }
            if (-not $force) {
                $confirm = Read-Host "Type DELETE to confirm"
                if ($confirm -ne "DELETE") {
                    Write-Info "Canceled."
                    return
                }
            }
            $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback -AllowNullResponse
            if ($resp -ne $null) { Write-Info "Directory setting deleted." } else { Write-Info "Delete requested." }
        }
        default {
            Write-Warn "Usage: dirsetting list|get|create|update|delete|template"
        }
    }
}



