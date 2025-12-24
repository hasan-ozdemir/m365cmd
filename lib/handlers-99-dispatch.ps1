# Handler: Dispatch
# Purpose: Dispatch command handlers.
function Handle-GlobalCommand {
    param(
        [string]$Cmd,
        [string[]]$InputArgs
    )
    switch ($Cmd) {
        "help"   { Show-Help ($InputArgs | Select-Object -First 1) }
        "exit"   { return $false }
        "quit"   { return $false }
        "clear"  { Clear-Host }
        "status" { Show-Status }
        "login"  { Invoke-Login ($InputArgs | Select-Object -First 1) }
        "logout" { Invoke-Logout }
        "whoami" {
            $ctx = Get-MgContextSafe
            if ($ctx) {
                Write-Host "Account : $($ctx.Account)"
                Write-Host "Tenant  : $($ctx.TenantId)"
                Write-Host ("Scopes  : " + ($ctx.Scopes -join ", "))
            } else {
                Write-Warn "Not connected."
            }
        }
        "tenant" {
            if (-not $InputArgs -or $InputArgs.Count -eq 0 -or $InputArgs[0] -eq "show") {
                Write-Host "Prefix : $($global:Config.tenant.defaultPrefix)"
                Write-Host "Domain : $($global:Config.tenant.defaultDomain)"
                Write-Host "ID     : $($global:Config.tenant.tenantId)"
                return $true
            }
            if ($InputArgs[0] -eq "set" -and $InputArgs.Count -ge 3) {
                $key = $InputArgs[1].ToLowerInvariant()
                $val = $InputArgs[2]
                switch ($key) {
                    "prefix" { Set-ConfigValue "tenant.defaultPrefix" $val }
                    "domain" { Set-ConfigValue "tenant.defaultDomain" $val }
                    "id"     { Set-ConfigValue "tenant.tenantId" $val }
                    default  { Write-Warn "Unknown tenant key. Use prefix|domain|id" }
                }
                Write-Info "Tenant defaults updated."
                return $true
            }
            Write-Warn "Usage: /tenant show | /tenant set prefix|domain|id <value>"
        }
        "config" {
            if (-not $InputArgs -or $InputArgs.Count -eq 0 -or $InputArgs[0] -eq "show") {
                Write-Host ($global:Config | ConvertTo-Json -Depth 8)
                return $true
            }
            $sub = $InputArgs[0].ToLowerInvariant()
            if ($sub -eq "get" -and $InputArgs.Count -ge 2) {
                $val = Get-ConfigValue $InputArgs[1]
                if ($val -is [string]) {
                    Write-Host $val
                } else {
                    Write-Host ($val | ConvertTo-Json -Depth 8)
                }
                return $true
            }
            if ($sub -eq "set" -and $InputArgs.Count -ge 3) {
                $path = $InputArgs[1]
                $raw = $InputArgs[2]
                $val = Parse-Value $raw
                Set-ConfigValue $path $val
                Write-Info "Config updated."
                return $true
            }
            Write-Warn "Usage: /config show | /config get <path> | /config set <path> <json-or-text>"
        }
        default {
            Write-Warn "Unknown global command. Use /help."
        }
    }
    return $true
}



function Handle-LocalCommand {
    param(
        [string]$Cmd,
        [string[]]$InputArgs
    )
    switch ($Cmd) {
        "module"  { Handle-ModuleCommand $InputArgs }
        "admin"   { Handle-AdminPortalCommand $InputArgs }
        "m365cli" { Handle-M365CliCommand $InputArgs }
        "m365"    { Handle-M365Command $InputArgs }
        "user"    { Handle-UserCommand $InputArgs }
        "license" { Handle-LicenseCommand $InputArgs }
        "role"    { Handle-RoleCommand $InputArgs }
        "group"   { Handle-GroupCommand $InputArgs }
        "domain"  { Handle-DomainCommand $InputArgs }
        "org"     { Handle-OrgCommand $InputArgs }
        "dirsetting" { Handle-DirSettingCommand $InputArgs }
        "site"    { Handle-SiteCommand $InputArgs }
        "splist"  { Handle-SPListCommand $InputArgs }
        "spage"   { Handle-SPPageCommand $InputArgs }
        "spcolumn" { Handle-SPColumnCommand $InputArgs }
        "spctype" { Handle-SPContentTypeCommand $InputArgs }
        "spperm" { Handle-SPPermissionCommand $InputArgs }
        "search"  { Handle-SearchCommand $InputArgs }
        "extconn" { Handle-ExternalConnectionCommand $InputArgs }
        "drive"   { Handle-DriveCommand $InputArgs }
        "file"    { Handle-FileCommand $InputArgs }
        "onedrive" { Handle-OneDriveCommand $InputArgs }
        "meeting" { Handle-MeetingCommand $InputArgs }
        "mail"    { Handle-MailCommand $InputArgs }
        "outlook" { Handle-OutlookCommand $InputArgs }
        "calendar" { Handle-CalendarCommand $InputArgs }
        "contacts" { Handle-ContactsCommand $InputArgs }
        "people" { Handle-PeopleCommand $InputArgs }
        "todo"    { Handle-TodoCommand $InputArgs }
        "planner" { Handle-PlannerCommand $InputArgs }
        "excel"   { Handle-ExcelCommand $InputArgs }
        "onenote" { Handle-OneNoteCommand $InputArgs }
        "word"    { Handle-WordCommand $InputArgs }
        "powerpoint" { Handle-PowerPointCommand $InputArgs }
        "visio"   { Handle-VisioCommand $InputArgs }
        "authmethod" { Handle-AuthMethodCommand $InputArgs }
        "risk"    { Handle-RiskCommand $InputArgs }
        "subscription" { Handle-SubscriptionCommand $InputArgs }
        "teamstab" { Handle-TeamsTabCommand $InputArgs }
        "teamsapp" { Handle-TeamsAppCommand $InputArgs }
        "teamsappinst" { Handle-TeamsAppInstallCommand $InputArgs }
        "chat"    { Handle-ChatCommand $InputArgs }
        "channelmsg" { Handle-ChannelMessageCommand $InputArgs }
        "device"  { Handle-DeviceCommand $InputArgs }
        "audit"   { Handle-AuditCommand $InputArgs }
        "auditfeed" { Handle-AuditFeedCommand $InputArgs }
        "report"  { Handle-ReportCommand $InputArgs }
        "forms"   { Handle-FormsCommand $InputArgs }
        "stream"  { Handle-StreamCommand $InputArgs }
        "clipchamp" { Handle-ClipchampCommand $InputArgs }
        "copilot" { Handle-CopilotCommand $InputArgs }
        "bookings" { Handle-BookingsCommand $InputArgs }
        "orgx"    { Handle-OrgXCommand $InputArgs }
        "orgexplorer" { Handle-OrgXCommand $InputArgs }
        "whiteboard" { Handle-WhiteboardCommand $InputArgs }
        "apps" { Handle-AppsCommand $InputArgs }
        "insights" { Handle-InsightsCommand $InputArgs }
        "connections" { Handle-ConnectionsCommand $InputArgs }
        "engage" { Handle-EngageCommand $InputArgs }
        "yammer" { Handle-EngageCommand $InputArgs }
        "lists" { Handle-ListsCommand $InputArgs }
        "learning" { Handle-LearningCommand $InputArgs }
        "loop" { Handle-LoopCommand $InputArgs }
        "sway" { Handle-SwayCommand $InputArgs }
        "kaizala" { Handle-KaizalaCommand $InputArgs }
        "security" { Handle-SecurityCommand $InputArgs }
        "defender" { Handle-DefenderCommand $InputArgs }
        "ca"      { Handle-CACommand $InputArgs }
        "accessreview" { Handle-AccessReviewCommand $InputArgs }
        "intune"  { Handle-IntuneCommand $InputArgs }
        "label"   { Handle-LabelCommand $InputArgs }
        "billing" { Handle-BillingCommand $InputArgs }
        "pp"      { Handle-PPCommand $InputArgs }
        "powerapps" { Handle-PowerAppsCommand $InputArgs }
        "powerautomate" { Handle-PowerAutomateCommand $InputArgs }
        "powerpages" { Handle-PowerPagesCommand $InputArgs }
        "viva"    { Handle-VivaCommand $InputArgs }
        "purview" { Handle-PurviewCommand $InputArgs }
        "compliance" { Handle-ComplianceCommand $InputArgs }
        "health"  { Handle-HealthCommand $InputArgs }
        "message" { Handle-MessageCommand $InputArgs }
        "exo"     { Handle-ExoCommand $InputArgs }
        "teams"   { Handle-TeamsCommand $InputArgs }
        "spo"     { Handle-SpoCommand $InputArgs }
        "addin"   { Handle-AddinCommand $InputArgs }
        "alias"   { Handle-AliasCommand $InputArgs }
        "preset"  { Handle-PresetCommand $InputArgs }
        "manifest" { Handle-ManifestCommand $InputArgs }
        "app"     { Handle-AppCommand $InputArgs }
        "graph"   { Handle-GraphCommand $InputArgs }
        "webhook" { Handle-WebhookCommand $InputArgs }
        default   { Write-Warn "Unknown command. Use /help." }
    }
}



