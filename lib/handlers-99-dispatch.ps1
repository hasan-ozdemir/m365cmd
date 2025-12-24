# Handler: Dispatch
# Purpose: Dispatch command handlers.
function Handle-GlobalCommand {
    param(
        [string]$Cmd,
        [string[]]$Args
    )
    switch ($Cmd) {
        "help"   { Show-Help ($Args | Select-Object -First 1) }
        "exit"   { return $false }
        "quit"   { return $false }
        "clear"  { Clear-Host }
        "status" { Show-Status }
        "login"  { Invoke-Login ($Args | Select-Object -First 1) }
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
            if (-not $Args -or $Args.Count -eq 0 -or $Args[0] -eq "show") {
                Write-Host "Prefix : $($global:Config.tenant.defaultPrefix)"
                Write-Host "Domain : $($global:Config.tenant.defaultDomain)"
                Write-Host "ID     : $($global:Config.tenant.tenantId)"
                return $true
            }
            if ($Args[0] -eq "set" -and $Args.Count -ge 3) {
                $key = $Args[1].ToLowerInvariant()
                $val = $Args[2]
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
            if (-not $Args -or $Args.Count -eq 0 -or $Args[0] -eq "show") {
                Write-Host ($global:Config | ConvertTo-Json -Depth 8)
                return $true
            }
            $sub = $Args[0].ToLowerInvariant()
            if ($sub -eq "get" -and $Args.Count -ge 2) {
                $val = Get-ConfigValue $Args[1]
                if ($val -is [string]) {
                    Write-Host $val
                } else {
                    Write-Host ($val | ConvertTo-Json -Depth 8)
                }
                return $true
            }
            if ($sub -eq "set" -and $Args.Count -ge 3) {
                $path = $Args[1]
                $raw = $Args[2]
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
        [string[]]$Args
    )
    switch ($Cmd) {
        "module"  { Handle-ModuleCommand $Args }
        "admin"   { Handle-AdminPortalCommand $Args }
        "m365cli" { Handle-M365CliCommand $Args }
        "m365"    { Handle-M365CliCommand $Args }
        "user"    { Handle-UserCommand $Args }
        "license" { Handle-LicenseCommand $Args }
        "role"    { Handle-RoleCommand $Args }
        "group"   { Handle-GroupCommand $Args }
        "domain"  { Handle-DomainCommand $Args }
        "org"     { Handle-OrgCommand $Args }
        "dirsetting" { Handle-DirSettingCommand $Args }
        "site"    { Handle-SiteCommand $Args }
        "splist"  { Handle-SPListCommand $Args }
        "spage"   { Handle-SPPageCommand $Args }
        "spcolumn" { Handle-SPColumnCommand $Args }
        "spctype" { Handle-SPContentTypeCommand $Args }
        "spperm" { Handle-SPPermissionCommand $Args }
        "search"  { Handle-SearchCommand $Args }
        "extconn" { Handle-ExternalConnectionCommand $Args }
        "drive"   { Handle-DriveCommand $Args }
        "file"    { Handle-FileCommand $Args }
        "onedrive" { Handle-OneDriveCommand $Args }
        "meeting" { Handle-MeetingCommand $Args }
        "mail"    { Handle-MailCommand $Args }
        "outlook" { Handle-OutlookCommand $Args }
        "calendar" { Handle-CalendarCommand $Args }
        "contacts" { Handle-ContactsCommand $Args }
        "people" { Handle-PeopleCommand $Args }
        "todo"    { Handle-TodoCommand $Args }
        "planner" { Handle-PlannerCommand $Args }
        "excel"   { Handle-ExcelCommand $Args }
        "onenote" { Handle-OneNoteCommand $Args }
        "word"    { Handle-WordCommand $Args }
        "powerpoint" { Handle-PowerPointCommand $Args }
        "visio"   { Handle-VisioCommand $Args }
        "authmethod" { Handle-AuthMethodCommand $Args }
        "risk"    { Handle-RiskCommand $Args }
        "subscription" { Handle-SubscriptionCommand $Args }
        "teamstab" { Handle-TeamsTabCommand $Args }
        "teamsapp" { Handle-TeamsAppCommand $Args }
        "teamsappinst" { Handle-TeamsAppInstallCommand $Args }
        "chat"    { Handle-ChatCommand $Args }
        "channelmsg" { Handle-ChannelMessageCommand $Args }
        "device"  { Handle-DeviceCommand $Args }
        "audit"   { Handle-AuditCommand $Args }
        "auditfeed" { Handle-AuditFeedCommand $Args }
        "report"  { Handle-ReportCommand $Args }
        "forms"   { Handle-FormsCommand $Args }
        "stream"  { Handle-StreamCommand $Args }
        "clipchamp" { Handle-ClipchampCommand $Args }
        "copilot" { Handle-CopilotCommand $Args }
        "bookings" { Handle-BookingsCommand $Args }
        "orgx"    { Handle-OrgXCommand $Args }
        "orgexplorer" { Handle-OrgXCommand $Args }
        "whiteboard" { Handle-WhiteboardCommand $Args }
        "apps" { Handle-AppsCommand $Args }
        "insights" { Handle-InsightsCommand $Args }
        "connections" { Handle-ConnectionsCommand $Args }
        "engage" { Handle-EngageCommand $Args }
        "yammer" { Handle-EngageCommand $Args }
        "lists" { Handle-ListsCommand $Args }
        "learning" { Handle-LearningCommand $Args }
        "loop" { Handle-LoopCommand $Args }
        "sway" { Handle-SwayCommand $Args }
        "kaizala" { Handle-KaizalaCommand $Args }
        "security" { Handle-SecurityCommand $Args }
        "defender" { Handle-DefenderCommand $Args }
        "ca"      { Handle-CACommand $Args }
        "accessreview" { Handle-AccessReviewCommand $Args }
        "intune"  { Handle-IntuneCommand $Args }
        "label"   { Handle-LabelCommand $Args }
        "billing" { Handle-BillingCommand $Args }
        "pp"      { Handle-PPCommand $Args }
        "powerapps" { Handle-PowerAppsCommand $Args }
        "powerautomate" { Handle-PowerAutomateCommand $Args }
        "powerpages" { Handle-PowerPagesCommand $Args }
        "viva"    { Handle-VivaCommand $Args }
        "purview" { Handle-PurviewCommand $Args }
        "compliance" { Handle-ComplianceCommand $Args }
        "health"  { Handle-HealthCommand $Args }
        "message" { Handle-MessageCommand $Args }
        "exo"     { Handle-ExoCommand $Args }
        "teams"   { Handle-TeamsCommand $Args }
        "spo"     { Handle-SpoCommand $Args }
        "addin"   { Handle-AddinCommand $Args }
        "alias"   { Handle-AliasCommand $Args }
        "preset"  { Handle-PresetCommand $Args }
        "manifest" { Handle-ManifestCommand $Args }
        "app"     { Handle-AppCommand $Args }
        "graph"   { Handle-GraphCommand $Args }
        "webhook" { Handle-WebhookCommand $Args }
        default   { Write-Warn "Unknown command. Use /help." }
    }
}


