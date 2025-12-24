# Core: Config
# Purpose: Config shared utilities.
function Get-DefaultConfig {
    [ordered]@{
        tenant = [ordered]@{
            defaultPrefix = "bizyum"
            defaultDomain = "bizyum.onmicrosoft.com"
            tenantId      = ""
        }
        admin = [ordered]@{
            defaultUpn = "info@prodyum.com"
        }
        auth = [ordered]@{
            scopes       = @(
                "User.Read.All",
                "User.ReadWrite.All",
                "Directory.ReadWrite.All",
            "RoleManagement.ReadWrite.Directory",
            "Group.ReadWrite.All",
            "GroupMember.ReadWrite.All",
            "ExternalItem.ReadWrite.All",
            "ExternalConnection.ReadWrite.All",
            "Domain.ReadWrite.All",
            "Application.ReadWrite.All",
            "AppRoleAssignment.ReadWrite.All",
            "LicenseAssignment.Read.All",
            "Organization.ReadWrite.All",
            "GroupSettings.ReadWrite.All",
            "Files.ReadWrite.All",
            "Sites.ReadWrite.All",
            "Sites.Manage.All",
            "Mail.ReadWrite",
            "Mail.Send",
            "Calendars.ReadWrite",
            "Contacts.ReadWrite",
            "Chat.ReadWrite",
            "ChatMessage.ReadWrite",
            "ChannelMessage.ReadWrite.All",
            "TeamsTab.ReadWrite.All",
            "TeamsAppInstallation.ReadWriteForTeam",
            "TeamsAppInstallation.ReadWriteForUser",
            "TeamsAppInstallation.ReadWriteForChat",
            "People.Read.All",
            "UserAuthenticationMethod.ReadWrite.All",
            "Tasks.ReadWrite",
            "Notes.ReadWrite.All",
            "Bookings.ReadWrite.All",
            "Organization.Read.All",
            "ServiceHealth.Read.All",
            "ServiceMessage.Read.All",
            "IdentityRiskEvent.Read.All",
            "IdentityRiskyUser.ReadWrite.All",
            "LearningProvider.ReadWrite",
            "Community.ReadWrite.All",
            "AuditLog.Read.All",
            "Reports.Read.All",
            "Subscriptions.ReadWrite.All",
            "AppCatalog.ReadWrite.All",
            "SecurityEvents.Read.All",
            "SecurityEvents.ReadWrite.All",
            "SecurityAlert.Read.All",
            "SecurityAlert.ReadWrite.All",
            "SecurityIncident.Read.All",
            "SecurityIncident.ReadWrite.All",
            "ThreatIndicators.ReadWrite.All",
            "ThreatHunting.Read.All",
            "Device.Read.All",
            "DeviceManagementManagedDevices.Read.All",
            "DeviceManagementManagedDevices.ReadWrite.All",
            "DeviceManagementApps.Read.All",
            "DeviceManagementApps.ReadWrite.All",
            "DeviceManagementConfiguration.Read.All",
            "DeviceManagementConfiguration.ReadWrite.All",
            "DeviceManagementServiceConfig.Read.All",
            "Policy.Read.All",
            "InformationProtectionPolicy.Read.All",
            "RecordsManagement.Read.All",
            "RecordsManagement.ReadWrite.All",
            "eDiscovery.ReadWrite.All",
            "SubjectRightsRequest.ReadWrite.All",
            "Policy.ReadWrite.ConditionalAccess",
            "AccessReview.ReadWrite.All",
            "OrgSettings-Forms.Read.All",
            "OrgSettings-Forms.ReadWrite.All",
            "OnlineMeetingTranscript.Read.All",
            "OnlineMeetingTranscript.Read.Chat",
            "OnlineMeetingRecording.Read.All",
            "OnlineMeetingRecording.Read.Chat",
            "OnlineMeetings.Read.All",
            "OnlineMeetings.ReadWrite"
        )
            loginMode    = "interactive"
            contextScope = "CurrentUser"
            publicClientId = "04f0c124-f2bc-4f4b-8c20-74bf17c5a1f9"
            app = [ordered]@{
                clientId     = ""
                clientSecret = ""
                tenantId     = ""
            }
        }
        modules = [ordered]@{
            required    = @("Microsoft.Graph")
            optional    = @(
                "Microsoft.Graph.Beta",
                "ExchangeOnlineManagement",
                "MicrosoftTeams",
                "Microsoft.Online.SharePoint.PowerShell",
                "MSAL.PS",
                "O365CentralizedAddInDeployment"
            )
            installPath = ".\\modules"
            autoInstall = $true
        }
        graph = [ordered]@{
            defaultApi           = "v1"
            fallbackToBeta       = $true
            autoSyncMetadata     = $true
            metadataRefreshHours = 24
        }
        pp = [ordered]@{
            baseUrl     = "https://api.powerplatform.com"
            apiVersion  = "2022-03-01-preview"
            clientId    = "04f0c124-f2bc-4f4b-8c20-74bf17c5a1f9"
            loginMode   = "interactive"
        }
        forms = [ordered]@{
            baseUrl = "https://forms.office.com"
            apiBase = "/formapi/api"
        }
        o365 = [ordered]@{
            manageApiBase = "https://manage.office.com"
            publisherId   = ""
        }
        defender = [ordered]@{
            baseUrl       = "https://api.security.microsoft.com"
            centerBaseUrl = "https://api.securitycenter.microsoft.com"
        }
        aliases = [ordered]@{
            global = [ordered]@{}
            local  = [ordered]@{
                "u" = "user {args}"
                "g" = "group {args}"
                "r" = "role {args}"
            }
        }
        presets = [ordered]@{
            "user-hard-reset" = "user disable {args}; user session revoke {args}; user mfa reset {args}"
        }
        output = [ordered]@{
            logPath  = ".\\logs\\m365cmd.log"
            dataPath = ".\\data"
        }
    }
}



function Test-ConfigKey {
    param(
        [object]$Target,
        [string]$Key
    )
    if ($null -eq $Target) { return $false }
    if ($Target -is [System.Collections.IDictionary]) {
        return $Target.Contains($Key)
    }
    return ($Target.PSObject.Properties.Name -contains $Key)
}

function Set-ConfigKey {
    param(
        [object]$Target,
        [string]$Key,
        [object]$Value
    )
    if ($Target -is [System.Collections.IDictionary]) {
        $Target[$Key] = $Value
        return
    }
    $Target | Add-Member -NotePropertyName $Key -NotePropertyValue $Value -Force
}


function Save-Config {
    param([object]$Config)
    $json = $Config | ConvertTo-Json -Depth 8
    Set-Content -Path $Paths.Config -Value $json -Encoding ASCII
}



function Load-Config {
    if (-not (Test-Path $Paths.Config)) {
        $cfg = Get-DefaultConfig
        Save-Config $cfg
        return $cfg
    }
    try {
        return (Get-Content -Raw $Paths.Config | ConvertFrom-Json)
    } catch {
        Write-Warn "Config JSON is invalid. Recreating defaults."
        Copy-Item -Path $Paths.Config -Destination ($Paths.Config + ".bad") -Force
        $cfg = Get-DefaultConfig
        Save-Config $cfg
        return $cfg
    }
}



function Normalize-Config {
    param([object]$Config)
    if (-not $Config) { return (Get-DefaultConfig) }
    $defaults = Get-DefaultConfig
    $changed = $false

    foreach ($section in @("tenant", "admin", "auth", "modules", "graph", "pp", "forms", "o365", "defender", "aliases", "presets", "output")) {
        if (-not $Config.$section) {
            $Config | Add-Member -NotePropertyName $section -NotePropertyValue $defaults.$section -Force
            $changed = $true
        }
    }

    if (-not $Config.auth.scopes) {
        $Config.auth.scopes = @($defaults.auth.scopes)
        $changed = $true
    } else {
        $scopes = @($Config.auth.scopes)
        foreach ($s in @($defaults.auth.scopes)) {
            if ($scopes -notcontains $s) {
                $scopes += $s
                $changed = $true
            }
        }
        $Config.auth.scopes = $scopes
    }

    foreach ($key in @("required", "optional")) {
        if (-not $Config.modules.$key) {
            $Config.modules.$key = @($defaults.modules.$key)
            $changed = $true
        } else {
            $list = @($Config.modules.$key)
            foreach ($m in @($defaults.modules.$key)) {
                if ($list -notcontains $m) {
                    $list += $m
                    $changed = $true
                }
            }
            $Config.modules.$key = $list
        }
    }
    if (-not $Config.modules.installPath) {
        $Config.modules.installPath = $defaults.modules.installPath
        $changed = $true
    }
    if ($null -eq $Config.modules.autoInstall) {
        $Config.modules.autoInstall = $defaults.modules.autoInstall
        $changed = $true
    }

    foreach ($key in @("logPath", "dataPath")) {
        if (-not $Config.output.$key) {
            $Config.output.$key = $defaults.output.$key
            $changed = $true
        }
    }

    foreach ($key in @("defaultApi", "fallbackToBeta", "autoSyncMetadata", "metadataRefreshHours")) {
        if ($null -eq $Config.graph.$key) {
            $Config.graph.$key = $defaults.graph.$key
            $changed = $true
        }
    }

    foreach ($key in @("baseUrl", "apiVersion", "clientId", "loginMode")) {
        if (-not $Config.pp.$key) {
            $Config.pp.$key = $defaults.pp.$key
            $changed = $true
        }
    }

    foreach ($key in @("baseUrl", "apiBase")) {
        if (-not $Config.forms.$key) {
            $Config.forms.$key = $defaults.forms.$key
            $changed = $true
        }
    }

    if (-not $Config.auth.app) {
        $Config.auth | Add-Member -NotePropertyName app -NotePropertyValue $defaults.auth.app -Force
        $changed = $true
    } else {
        foreach ($key in @("clientId", "clientSecret", "tenantId")) {
            if ($null -eq $Config.auth.app.$key) {
                $Config.auth.app.$key = $defaults.auth.app.$key
                $changed = $true
            }
        }
    }
    if (-not $Config.auth.publicClientId) {
        $Config.auth.publicClientId = $defaults.auth.publicClientId
        $changed = $true
    }

    foreach ($key in @("manageApiBase", "publisherId")) {
        if (-not $Config.o365.$key) {
            $Config.o365.$key = $defaults.o365.$key
            $changed = $true
        }
    }

    foreach ($key in @("baseUrl", "centerBaseUrl")) {
        if (-not $Config.defender.$key) {
            $Config.defender.$key = $defaults.defender.$key
            $changed = $true
        }
    }

    if (-not $Config.aliases.global) {
        $Config.aliases.global = $defaults.aliases.global
        $changed = $true
    }
    if (-not $Config.aliases.local) {
        $Config.aliases.local = $defaults.aliases.local
        $changed = $true
    } else {
        foreach ($k in $defaults.aliases.local.Keys) {
            if (-not (Test-ConfigKey $Config.aliases.local $k)) {
                Set-ConfigKey $Config.aliases.local $k $defaults.aliases.local[$k]
                $changed = $true
            }
        }
    }

    if (-not $Config.presets) {
        $Config.presets = $defaults.presets
        $changed = $true
    } else {
        foreach ($k in $defaults.presets.Keys) {
            if (-not (Test-ConfigKey $Config.presets $k)) {
                Set-ConfigKey $Config.presets $k $defaults.presets[$k]
                $changed = $true
            }
        }
    }

    if ($changed) {
        Save-Config $Config
    }
    return $Config
}



function Get-ConfigValue {
    param([string]$Path)
    $parts = $Path -split '\.'
    $current = $global:Config
    foreach ($p in $parts) {
        if ($null -eq $current) { return $null }
        $current = $current.$p
    }
    return $current
}



function Set-ConfigValue {
    param(
        [string]$Path,
        [object]$Value
    )
    $parts = $Path -split '\.'
    $current = $global:Config
    for ($i = 0; $i -lt ($parts.Count - 1); $i++) {
        $key = $parts[$i]
        if ($null -eq $current.$key) {
            $current | Add-Member -NotePropertyName $key -NotePropertyValue ([pscustomobject]@{}) -Force
        }
        $current = $current.$key
    }
    $last = $parts[-1]
    $current | Add-Member -NotePropertyName $last -NotePropertyValue $Value -Force
    Save-Config $global:Config
}



function Show-Status {
    $cfg = $global:Config
    Write-Host "Tenant prefix : $($cfg.tenant.defaultPrefix)"
    Write-Host "Tenant domain : $($cfg.tenant.defaultDomain)"
    if ($cfg.tenant.tenantId) {
        Write-Host "Tenant ID     : $($cfg.tenant.tenantId)"
    }
    Write-Host "Admin UPN     : $($cfg.admin.defaultUpn)"
    if ($cfg.graph) {
        Write-Host ("Graph default : " + $cfg.graph.defaultApi)
        Write-Host ("Graph fallback: " + $cfg.graph.fallbackToBeta)
    }
    if ($cfg.pp) {
        Write-Host ("PP base URL  : " + $cfg.pp.baseUrl)
        Write-Host ("PP api ver   : " + $cfg.pp.apiVersion)
    }
    Write-Host "Modules path  : $($Paths.Modules)"
    Write-Host ("Auto-install  : " + $cfg.modules.autoInstall)
    $graphOk = Test-ModuleAvailable "Microsoft.Graph"
    Write-Host ("Graph module  : " + ($(if ($graphOk) { "available" } else { "missing" })))
    $ctx = Get-MgContextSafe
    if ($ctx) {
        Write-Host "Graph login   : connected"
        if ($ctx.Account) { Write-Host "Account       : $($ctx.Account)" }
        if ($ctx.TenantId) { Write-Host "Tenant ID     : $($ctx.TenantId)" }
        if ($ctx.Scopes) { Write-Host ("Scopes        : " + ($ctx.Scopes -join ", ")) }
    } else {
        Write-Host "Graph login   : not connected"
    }

    $appClient = if ($global:Config.auth.app.clientId) { "set" } else { "not set" }
    Write-Host ("App creds     : " + $appClient)

    if ($global:PpToken -and $global:PpTokenExpires -gt (Get-Date)) {
        Write-Host ("PP login      : connected (expires " + $global:PpTokenExpires.ToString("s") + ")")
    } else {
        Write-Host "PP login      : not connected"
    }

    if ($cfg.forms) {
        Write-Host ("Forms base    : " + $cfg.forms.baseUrl)
    }

    if ($cfg.defender) {
        Write-Host ("Defender base : " + $cfg.defender.baseUrl)
        if ($cfg.defender.centerBaseUrl) {
            Write-Host ("Defender alt  : " + $cfg.defender.centerBaseUrl)
        }
    }
}



