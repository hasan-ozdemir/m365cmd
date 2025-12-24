# Handler: Help
# Purpose: Help command handlers.
function Show-Help {
    param([string]$Topic)
    $globalHelp = @(
        "/help [topic]        Show help for commands",
        "/exit                Exit m365cmd",
        "/quit                Exit m365cmd",
        "/clear               Clear the screen",
        "/status              Show connection and config status",
        "/login [interactive|device]  Login to Microsoft Graph",
        "/logout              Disconnect from Microsoft Graph",
        "/whoami              Show current Graph context",
        "/tenant show|set     Show or set tenant defaults",
        "/config show|get|set Show or update config values"
    )
    $localHelp = @(
        "admin    Admin center helpers",
        "user     User CRUD commands",
        "license  License management commands",
        "role     Role assignment commands",
        "group    Group CRUD + members/owners",
        "domain   Domain management",
        "org      Organization profile (Graph)",
        "dirsetting Directory settings (Graph)",
        "site     SharePoint sites (Graph)",
        "splist   SharePoint lists/items (Graph)",
        "spage    SharePoint pages/news (Graph)",
        "spcolumn SharePoint columns (Graph)",
        "spctype  SharePoint content types (Graph)",
        "spperm   SharePoint permissions (Graph)",
        "drive    Drive list/get (OneDrive/SharePoint)",
        "file     Files/folders CRUD + upload/download",
        "onedrive OneDrive sharing alias (file share)",
        "word     Word file helpers (search + file ops)",
        "powerpoint PowerPoint file helpers (search + file ops)",
        "visio    Visio file helpers (search + file ops)",
        "meeting  Online meeting transcripts/recordings",
        "search   Microsoft Graph search",
        "extconn  External connections/items (Search)",
        "mail     Mail folders/messages CRUD + send",
        "outlook  Outlook hub (mail/calendar/contacts/todo)",
        "calendar Calendar + events CRUD + view",
        "contacts People/Contacts CRUD",
        "people   People directory (read-only) + contacts CRUD alias",
        "authmethod User authentication methods (MFA/TAP)",
        "risk     Identity Protection risky users/detections",
        "todo     Microsoft To Do lists/tasks CRUD",
        "planner  Planner plans/buckets/tasks CRUD",
        "excel    Excel workbook/worksheet/table/range CRUD",
        "onenote  OneNote notebooks/sections/pages CRUD",
        "subscription Graph subscriptions/webhooks",
        "teamstab Teams tabs (Graph)",
        "teamsapp Teams app catalog (Graph)",
        "teamsappinst Teams app installations",
        "chat     Teams chats + chat messages CRUD",
        "channelmsg Teams channel messages CRUD",
        "device   Entra devices CRUD",
        "audit    Audit logs (directory/sign-in/provisioning)",
        "auditfeed O365 Management Activity API",
        "report   Microsoft 365 usage reports",
        "forms    Microsoft Forms (open/admin/reports/excel)",
        "stream   Microsoft Stream (on SharePoint) helpers",
        "clipchamp Clipchamp projects/assets (files + search)",
        "copilot  Microsoft 365 Copilot APIs",
        "bookings Microsoft Bookings (Graph)",
        "orgx     Org Explorer (manager/direct reports)",
        "whiteboard Whiteboard files (OneDrive) helper",
        "apps     App catalog + portal/CLI bridging",
        "insights Personal insights (shared/trending/used)",
        "connections Viva Connections helpers (SharePoint-backed)",
        "engage   Viva Engage/Yammer helpers",
        "lists    Microsoft Lists (SharePoint lists)",
        "learning Viva Learning + learning activities",
        "loop     Loop components (.loop files) helpers",
        "sway     Sway shortcuts + URI scheme",
        "kaizala  Kaizala (retired)",
        "security Security + Defender data",
        "defender Microsoft Defender XDR API",
        "ca       Conditional Access policies + named locations",
        "intune   Intune device management",
        "label    Sensitivity labels (beta, read-only)",
        "billing  Billing/commerce (subscriptions + SKUs)",
        "pp       Power Platform admin (environments)",
        "powerapps Power Apps (Power Platform)",
        "powerautomate Power Automate (Power Platform)",
        "powerpages Power Pages (Power Platform)",
        "viva     Viva Learning (providers + content)",
        "purview  Compliance/Purview (eDiscovery + SRR)",
        "compliance Purview/Compliance PowerShell",
        "health   Service health (M365)",
        "message  Message center",
        "exo      Exchange Online commands",
        "teams    Microsoft Teams commands",
        "spo      SharePoint Online commands",
        "addin    Office add-in management (Outlook + centralized)",
        "alias    Command aliases",
        "preset   Command presets",
        "manifest Loader manifest tools",
        "app      App registration commands",
        "accessreview Access reviews (Identity Governance)",
        "graph    Graph cmdlets, metadata, generic CRUD",
        "webhook  Local webhook listener helper",
        "m365     Native CLI-compatible command group",
        "m365cli  CLI for Microsoft 365 bridge",
        "module   Module install/list/update/remove"
    )

    $topics = @{
        "admin" = @(
            "admin open",
            "admin user|license|role|group|domain|org|billing|health|message|security|purview|compliance <...>"
        )
        "user"    = @(
            "user list [--filter <odata>] [--select prop,prop]",
            "user get <upn|id>",
            "user create --upn <upn> OR --alias <name> [--domain domain.com] --displayName <name> --password <pwd> [--usageLocation TR] [--forceChange true|false]",
            "user update <upn|id> --set key=value[,key=value] OR --json <payload> OR --bodyFile <path>",
            "user delete <upn|id> [--force] OR --filter <odata> [--whatif] [--force]",
            "user bulkdelete --filter <odata> [--whatif] [--force]",
            "user props <upn|id>",
            "user apps <upn|id> [--explain]",
            "user roles <upn|id>",
            "user enable|disable <upn|id>",
            "user password set|reset|change <upn|id> --password <pwd> [--forceChange true|false]",
            "user upn set <upn|id> --upn <newUpn> OR --alias <name> [--domain domain.com]",
            "user email list|add|remove <upn|id> [--address <email>]",
            "user alias list|add|remove <upn|id> [--address <alias@domain>] [--primary]",
            "user session revoke <upn|id>",
            "user mfa reset <upn|id> [--force]",
            "user license list|assign|remove|update <upn|id> ..."
        )
        "license" = @(
            "license list",
            "license assign <upn|id> --sku <skuId> [--disablePlans planId,planId]",
            "license remove <upn|id> --sku <skuId>",
            "license update <upn|id> [--add skuId,skuId] [--remove skuId,skuId] [--disablePlans planId,planId]"
        )
        "role"    = @(
            "role list",
            "role templates",
            "role definitions list|get",
            "role assignments list [--principal <id|upn>] [--definition <roleDefinitionId|name>]",
            "role assign <upn|id> --role <roleName|roleId>",
            "role assign <principal> --definition <roleDefinitionId|name> [--scope /]",
            "role remove <upn|id> --role <roleName|roleId>",
            "role remove <principal> --definition <roleDefinitionId|name> OR --assignment <assignmentId>"
        )
        "group"   = @(
            "group list [--filter <odata>] [--select prop,prop]",
            "group get <id|displayName>",
            "group create --name <name> [--mailNickname nick] [--type unified|security|mailsecurity|distribution] [--description text] [--visibility Public|Private]",
            "group update <id> --set key=value[,key=value] OR --json <payload> OR --bodyFile <path>",
            "group delete <id|displayName> [--force] OR --filter <odata> [--whatif] [--force]",
            "group bulkdelete --filter <odata> [--whatif] [--force]",
            "group member list|add|remove <groupId|name> [--member <upn|id>] [--members a,b] [--objectId <id>] [--objectIds id1,id2]",
            "group owner list|add|remove <groupId|name> [--owner <upn|id>] [--owners a,b] [--objectId <id>] [--objectIds id1,id2]"
        )
        "domain"  = @(
            "domain list",
            "domain get <domain>",
            "domain add <domain>",
            "domain dns <domain>",
            "domain verify <domain>",
            "domain default <domain>"
        )
        "org" = @(
            "org list [--select prop,prop] [--beta|--auto]",
            "org get [orgId] [--beta|--auto]",
            "org update [orgId] --set key=value[,key=value] OR --json <payload> [--beta|--auto]"
        )
        "dirsetting" = @(
            "dirsetting template list|get [--beta]",
            "dirsetting list [--group <groupId>] [--beta|--auto]",
            "dirsetting get <settingId> [--group <groupId>] [--beta|--auto]",
            "dirsetting create --templateId <id> [--set key=value,...] [--values <json>] [--group <groupId>] [--beta|--auto]",
            "dirsetting update <settingId> --set key=value,... OR --values <json> [--group <groupId>] [--beta|--auto]",
            "dirsetting delete <settingId> [--force] [--group <groupId>] [--beta|--auto]"
        )
        "site" = @(
            "site list --search <text> [--top <n>]",
            "site get <siteId|hostname:/path:>"
        )
        "splist" = @(
            "splist list --site <siteId|url|hostname:/path:>",
            "splist get <listId> --site <siteId|url|hostname:/path:>",
            "splist create --site <siteId|url|hostname:/path:> [--name <text>] [--template genericList] OR --json <payload>",
            "splist update <listId> --site <siteId|url|hostname:/path:> --json <payload> OR --set key=value",
            "splist delete <listId> --site <siteId|url|hostname:/path:>",
            "splist delta --site <siteId> --list <listId> [--token <deltaLink>]",
            "splist item list --site <siteId> --list <listId>",
            "splist item get <itemId> --site <siteId> --list <listId>",
            "splist item create --site <siteId> --list <listId> --fields key=value[,key=value] OR --json <payload>",
            "splist item update <itemId> --site <siteId> --list <listId> --fields key=value[,key=value] OR --json <payload>",
            "splist item delete <itemId> --site <siteId> --list <listId>"
        )
        "spage" = @(
            "spage list --site <siteId|url|hostname:/path:>",
            "spage get <pageId> --site <siteId|url|hostname:/path:>",
            "spage create --site <siteId|url|hostname:/path:> --name <page.aspx> --title <text> [--news true|false] OR --json <payload>",
            "spage update <pageId> --site <siteId|url|hostname:/path:> --json <payload> OR --set key=value",
            "spage delete <pageId> --site <siteId|url|hostname:/path:>",
            "spage publish <pageId> --site <siteId|url|hostname:/path:>"
        )
        "spcolumn" = @(
            "spcolumn list --site <siteId|url|hostname:/path:> [--list <listId>]",
            "spcolumn get <columnId> --site <siteId|url|hostname:/path:> [--list <listId>]",
            "spcolumn create --site <siteId|url|hostname:/path:> [--list <listId>] --json <payload> OR --set key=value",
            "spcolumn update <columnId> --site <siteId|url|hostname:/path:> [--list <listId>] --json <payload> OR --set key=value",
            "spcolumn delete <columnId> --site <siteId|url|hostname:/path:> [--list <listId>]"
        )
        "spctype" = @(
            "spctype list --site <siteId|url|hostname:/path:> [--list <listId>]",
            "spctype get <contentTypeId> --site <siteId|url|hostname:/path:> [--list <listId>]",
            "spctype create --site <siteId|url|hostname:/path:> [--list <listId>] --json <payload> OR --set key=value",
            "spctype update <contentTypeId> --site <siteId|url|hostname:/path:> [--list <listId>] --json <payload> OR --set key=value",
            "spctype delete <contentTypeId> --site <siteId|url|hostname:/path:> [--list <listId>]"
        )
        "spperm" = @(
            "spperm list --site <siteId|url|hostname:/path:>",
            "spperm get <permissionId> --site <siteId|url|hostname:/path:>",
            "spperm grant --site <siteId|url|hostname:/path:> --json <payload> OR --set key=value",
            "spperm delete <permissionId> --site <siteId|url|hostname:/path:>"
        )
        "drive"   = @(
            "drive list [--user <upn|id>|--site <siteId>|--group <groupId>] [--top <n>] [--select prop,prop]",
            "drive get <driveId>",
            "drive delta [--user <upn|id>|--drive <id>|--site <id>|--group <id>] [--token <deltaLink>] [--beta]"
        )
        "search" = @(
            "search query --entity driveItem|message|event|site|list|listItem|externalItem --text <q> [--from <n>] [--size <n>]",
            "search query --requestsJson <payload>"
        )
        "extconn" = @(
            "extconn list|get|create|update|delete",
            "extconn item list|get|create|update|delete --conn <id>"
        )
        "file"    = @(
            "file list [--user <upn|id>|--drive <id>|--site <id>|--group <id>] [--path <path>|--item <id>]",
            "file get <itemId> [--path <path>] [--user <upn|id>]",
            "file create --name <name> [--path <parentPath>] [--folder] [--content <text>|--local <file>]",
            'file update <itemId> --set key=value[,key=value] OR --set ''{"name":"New"}''',
            "file delete <itemId> [--force]",
            "file download <itemId> --out <file>",
            "file convert <itemId> --out <file> [--format pdf|html|txt]",
            "file preview <itemId> [--path <path>] [--json <payload>]",
            "file upload --local <file> [--dest <path>]",
            "file copy <itemId> [--path <path>] --dest <folderPath> [--name <newName>]",
            "file move <itemId> [--path <path>] --dest <folderPath> [--name <newName>]",
            "file share list|get <itemId> [--path <path>]",
            "file share link create|update|delete <itemId> --perm <permissionId> [--type view|edit] [--scope anonymous|organization]",
            "file share invite <itemId> --to a@b.com,b@b.com [--roles read|write]"
        )
        "onedrive" = @(
            "onedrive share <same as: file share ...>"
        )
        "meeting" = @(
            "meeting list [--user <upn|id>]",
            "meeting get <meetingId> [--user <upn|id>]",
            "meeting create --json <payload> [--user <upn|id>]",
            "meeting update <meetingId> --json <payload> OR --set key=value [--user <upn|id>]",
            "meeting delete <meetingId> [--user <upn|id>]",
            "meeting transcript list --meeting <meetingId> [--user <upn|id>]",
            "meeting transcript get <transcriptId> --meeting <meetingId> [--user <upn|id>]",
            "meeting transcript content <transcriptId> --meeting <meetingId> --out <file> [--user <upn|id>]",
            "meeting recording list --meeting <meetingId> [--user <upn|id>]",
            "meeting recording get <recordingId> --meeting <meetingId> [--user <upn|id>]",
            "meeting recording content <recordingId> --meeting <meetingId> --out <file> [--user <upn|id>]"
        )
        "mail"    = @(
            "mail folder list|get|create|update|delete [--user <upn|id>]",
            "mail message list|get|create|update|delete|send [--user <upn|id>] [--folder <id>]"
        )
        "outlook" = @(
            "outlook mail|calendar|contacts|people|todo|meeting <...>"
        )
        "calendar" = @(
            "calendar list|get|create|update|delete [--user <upn|id>]",
            "calendar view [--user <upn|id>] [--range day|week|month|year] [--date YYYY-MM-DD] [--start <iso>] [--end <iso>] [--tz <tz>]",
            "calendar event list|get|create|update|delete [--user <upn|id>] [--calendar <id>]"
        )
        "contacts" = @(
            "contacts folder list|get|create|update|delete [--user <upn|id>]",
            "contacts item list|get|create|update|delete [--user <upn|id>] [--folder <id>]"
        )
        "people" = @(
            "people list [--user <upn|id>]",
            "people get <id> [--user <upn|id>]",
            "people create|update|delete --json <payload> (aliases to contacts item CRUD)"
        )
        "authmethod" = @(
            "authmethod list|get|delete [--user <upn|id>]",
            "authmethod phone list|get|add|update|delete [--user <upn|id>] [--number <phone>] [--type mobile|alternateMobile|office]",
            "authmethod email list|get|create|update|delete [--user <upn|id>] [--email <address>] [--beta|--v1|--auto]",
            "authmethod tap list|get|create|delete [--user <upn|id>] [--start <iso>] [--lifetime <minutes>] [--once true|false] [--beta|--v1|--auto]"
        )
        "risk" = @(
            "risk detection list|get [--filter <odata>] [--top <n>] [--beta|--v1|--auto]",
            "risk user list|get|history [--filter <odata>] [--top <n>] [--beta|--v1|--auto]",
            "risk user confirm --ids <id1,id2> OR --user <upn,id> [--beta|--v1|--auto]",
            "risk user dismiss --ids <id1,id2> OR --user <upn,id> [--beta|--v1|--auto]"
        )
        "todo" = @(
            "todo list|get|create|update|delete [--user <upn|id>]",
            "todo task list|get|create|update|delete --list <listId> [--user <upn|id>]"
        )
        "planner" = @(
            "planner plan list --group <groupId>",
            "planner plan get|create|update|delete [--id <id>] [--group <groupId>] [--title <text>]",
            "planner bucket list --plan <planId>",
            "planner bucket get|create|update|delete [--id <id>] [--plan <planId>] [--name <text>]",
            "planner task list [--plan <planId>|--bucket <bucketId>]",
            "planner task get|create|update|delete [--id <id>] [--plan <planId>] [--title <text>]"
        )
        "excel" = @(
            "excel workbook list|get [--user <upn|id>] [--drive <id>] [--path <path>|--item <id>]",
            "excel worksheet list|get|create|update|delete [--item <id>|--path <path>] [--user <upn|id>]",
            "excel table list|get|create|update|delete [--item <id>|--path <path>] [--user <upn|id>] [--worksheet <id>]",
            "excel range get|update [--item <id>|--path <path>] [--user <upn|id>] --address <A1>",
            "excel cell get|update [--item <id>|--path <path>] [--user <upn|id>] --row <n> --col <n>"
        )
        "word" = @(
            "word list|search [--query <kql>] [--siteUrl <url>] [--path <url>] [--top <n>] [--from <n>] [--beta|--auto] [--json]",
            "word get|download|upload|create|update|delete|convert|preview|share <...> (uses file args)"
        )
        "powerpoint" = @(
            "powerpoint list|search [--query <kql>] [--siteUrl <url>] [--path <url>] [--top <n>] [--from <n>] [--beta|--auto] [--json]",
            "powerpoint get|download|upload|create|update|delete|convert|preview|share <...> (uses file args)"
        )
        "visio" = @(
            "visio list|search [--query <kql>] [--siteUrl <url>] [--path <url>] [--top <n>] [--from <n>] [--beta|--auto] [--json]",
            "visio get|download|upload|create|update|delete|convert|preview|share <...> (uses file args)"
        )
        "onenote" = @(
            "onenote notebook list|get|create|update|delete [--user <upn|id>]",
            "onenote section list|get|create|update|delete [--user <upn|id>] [--notebook <id>]",
            "onenote page list|get|create|update|delete|content [--user <upn|id>] [--section <id>]"
        )
        "subscription" = @(
            "subscription list",
            "subscription get <id>",
            "subscription create --json <payload>",
            "subscription update <id> --json <payload> OR --set key=value",
            "subscription delete <id>"
        )
        "teamstab" = @(
            "teamstab list --team <teamId> --channel <channelId>",
            "teamstab list --chat <chatId>",
            "teamstab get <tabId> --team <teamId> --channel <channelId>",
            "teamstab create --team <teamId> --channel <channelId> --json <payload>",
            "teamstab update <tabId> --team <teamId> --channel <channelId> --json <payload> OR --set key=value",
            "teamstab delete <tabId> --team <teamId> --channel <channelId>",
            "teamstab get|create|update|delete --chat <chatId>"
        )
        "teamsapp" = @(
            "teamsapp list [--top <n>] [--select prop,prop] [--filter <odata>]",
            "teamsapp get <appId>",
            "teamsapp add --package <zipPath>",
            "teamsapp update <appId> --package <zipPath>",
            "teamsapp delete <appId>"
        )
        "teamsappinst" = @(
            "teamsappinst list --team <teamId>|--chat <chatId>|--user <upn|id> [--expand teamsAppDefinition]",
            "teamsappinst get <installationId> --team <teamId>|--chat <chatId>|--user <upn|id> [--expand teamsAppDefinition]",
            "teamsappinst add --team <teamId>|--chat <chatId>|--user <upn|id> --app <appId> OR --json <payload>",
            "teamsappinst remove <installationId> --team <teamId>|--chat <chatId>|--user <upn|id>"
        )
        "chat" = @(
            "chat list|get|create [--user <upn|id>]",
            "chat message list|get|create|update|delete --chat <chatId>"
        )
        "channelmsg" = @(
            "channelmsg list|get|create|update|delete --team <teamId> --channel <channelId>"
        )
        "device" = @(
            "device list [--filter <odata>] [--select prop,prop] [--top <n>]",
            "device get <id> [--select prop,prop]",
            "device update <id> --set key=value[,key=value] OR --json <payload>",
            "device delete <id> [--force]"
        )
        "audit" = @(
            "audit list [--type directory|signin|provisioning] [--filter <odata>] [--top <n>]",
            "audit get <id> [--type directory|signin|provisioning]"
        )
        "auditfeed" = @(
            "auditfeed list",
            "auditfeed start --type <Audit.AzureActiveDirectory|Audit.Exchange|Audit.SharePoint|Audit.General|DLP.All> [--webhook <url>] [--authId <id>] [--expiration <iso>]",
            "auditfeed stop --type <contentType>",
            "auditfeed content list --type <contentType> --start <iso> --end <iso>",
            "auditfeed content get --uri <contentUri> [--out <file>]",
            "auditfeed notifications list --type <contentType> [--start <iso>] [--end <iso>]"
        )
        "report" = @(
            "report list",
            "report run <name> [--period D7|D30|D90|D180] [--date YYYY-MM-DD] [--format csv|json] [--out <file>] [--beta|--auto]"
        )
        "security" = @(
            "security list [--type alerts|alerts_v2|incidents|secureScores|secureScoreControlProfiles] [--filter <odata>] [--top <n>]",
            "security get <id> [--type alerts|alerts_v2|incidents|secureScores|secureScoreControlProfiles] [--beta|--auto]",
            "security alert list|get|update [--legacy] [--filter <odata>] [--top <n>] [--beta|--auto]",
            "security incident list|get|update [--expand alerts] [--filter <odata>] [--top <n>]",
            "security hunt --query <kql> [--beta|--auto]",
            "security ti list|get|create|update|delete [--filter <odata>] [--top <n>]"
        )
        "defender" = @(
            "defender incident list|get|update [--filter <odata>] [--top <n>]",
            "defender alert list|get|update [--filter <odata>] [--top <n>]",
            "defender hunt --query <kql>",
            "defender machine list|get [--filter <odata>] [--top <n>]",
            "defender machine findbytag --tag <tag>",
            "defender machine isolate <id> --comment <text> [--type Full|Selective]",
            "defender machine unisolate <id> --comment <text>",
            "defender machine collect <id> --comment <text>",
            "defender machine runscan <id> --comment <text> [--type Full|Quick]",
            "defender machine stopfile <id> --comment <text> --sha1 <hash>",
            "defender machineaction list|get [--filter <odata>] [--top <n>]"
        )
        "ca" = @(
            "ca policy list|get|create|update|delete [--filter <odata>] [--top <n>] [--select prop,prop] [--beta|--auto]",
            "ca location list|get|create|update|delete [--filter <odata>] [--top <n>] [--beta|--auto]"
        )
        "intune" = @(
            "intune device list|get|create|update|delete [--filter <odata>] [--top <n>] [--select prop,prop] [--beta|--auto]",
            "intune config list|get|create|update|delete [--filter <odata>] [--top <n>] [--select prop,prop] [--beta|--auto]",
            "intune compliance list|get|create|update|delete [--filter <odata>] [--top <n>] [--select prop,prop] [--beta|--auto]",
            "intune app list|get|create|update|delete [--filter <odata>] [--top <n>] [--select prop,prop] [--beta|--auto]",
            "intune script list|get|create|update|delete [--filter <odata>] [--top <n>] [--select prop,prop] [--beta|--auto]",
            "intune shellscript list|get|create|update|delete [--filter <odata>] [--top <n>] [--select prop,prop] [--beta|--auto]",
            "intune healthscript list|get|create|update|delete [--filter <odata>] [--top <n>] [--select prop,prop] [--beta|--auto]",
            "intune report list|get|export|download|status [--beta|--auto]"
        )
        "label" = @(
            "label list|get [--user <upn|id>|--me|--org] [--beta]"
        )
        "billing" = @(
            "billing sku list|get [--select prop,prop]",
            "billing subscription list|get [--filter <odata>] [--top <n>] [--select prop,prop]"
        )
        "pp" = @(
            "pp login [--device] [--force]",
            "pp logout",
            "pp status",
            "pp env list|get|enable|disable [--filter <odata>] [--top <n>] [--skiptoken <token>]",
            "pp app list|get --env <environmentId> [--top <n>] [--skiptoken <token>]",
            "pp flow list|get|runs|actions --env <environmentId> [--workflowId <id>] [--top <n>] [--skiptoken <token>]",
            "pp req <method> <path> [--json <payload>] [--bodyFile <file>] [--apiVersion <ver>]"
        )
        "powerapps" = @(
            "powerapps list|get --env <environmentId> [--top <n>] [--skiptoken <token>]"
        )
        "powerautomate" = @(
            "powerautomate list|get|runs|actions --env <environmentId> [--workflowId <id>] [--top <n>] [--skiptoken <token>]"
        )
        "powerpages" = @(
            "powerpages site list|get|create|delete|restart --env <environmentId> [--skip <n>]",
            "powerpages op status --url <operationUrl>"
        )
        "viva" = @(
            "viva provider list|get|create|update|delete [--beta|--auto]",
            "viva content list|get|create|update|delete --provider <id> [--beta|--auto]"
        )
        "purview" = @(
            "purview ediscovery case list|get|create|update|delete [--beta|--auto]",
            "purview ediscovery custodian|datasource|hold|search|reviewset list|get|create|update|delete --case <caseId> [--beta|--auto]",
            "purview srr list|get|create|update [--beta|--auto]",
            "purview retention label list|get [--beta|--auto]",
            "purview dlp policy list|get|create|update|delete [--beta|--auto]"
        )
        "compliance" = @(
            "compliance connect [--upn <user>] [--delegatedOrg <domain>] [--disableWam true|false]",
            "compliance disconnect",
            "compliance status",
            "compliance cmdlets [--filter text]",
            "compliance cmd <cmdlet> [--params key=value[,key=value]] [--json <payload>]"
        )
        "health"  = @(
            "health list",
            "health issues <serviceId>"
        )
        "message" = @(
            "message list [--filter <odata>] [--top <n>]",
            "message get <messageId>"
        )
        "exo"     = @(
            "exo connect [--upn <user>] [--delegatedOrg <domain>] [--env <name>] [--disableWam true|false]",
            "exo disconnect",
            "exo status",
            "exo cmdlets [--filter text]",
            "exo cmd <cmdlet> [--params key=value[,key=value]] [--json <payload>]",
            "exo mailbox list|get|create|update|delete|perm",
            "exo mailbox perm <mailbox> list|add|remove [--user <upn>] [--rights FullAccess]",
            "exo addin list|get|install|update|remove|enable|disable|refresh [--org] [--mailbox <upn>]",
            "exo onsend status|enable|disable --policy <name> [--user <upn>] [--all] [--filter <filter>]"
        )
        "addin" = @(
            "addin exo list|get|install|update|remove|enable|disable|refresh [--org] [--mailbox <upn>]",
            "addin onsend status|enable|disable --policy <name> [--user <upn>] [--all] [--filter <filter>]",
            "addin org connect|disconnect|status",
            "addin org list|get|create|update|remove|assign|enable|disable|refresh"
        )
        "teams"   = @(
            "teams connect [--upn <user>] [--tenantId <id>]",
            "teams disconnect",
            "teams status",
            "teams cmdlets [--filter text]",
            "teams cmd <cmdlet> [--params key=value[,key=value]] [--json <payload>]",
            "teams list|get|create|delete",
            "teams user list|add|remove <groupId>",
            "teams channel list|create|remove <groupId>",
            "teams config get|update --set key=value[,key=value]",
            "teams policy list|get|create|update|delete --type messaging|meeting [name] [--set key=value]"
        )
        "spo"     = @(
            "spo connect [--url <adminUrl>] [--prefix <tenantPrefix>]",
            "spo disconnect",
            "spo status",
            "spo cmdlets [--filter text]",
            "spo cmd <cmdlet> [--params key=value[,key=value]] [--json <payload>]",
            "spo site list|get|create|update|delete",
            "spo tenant get|update --set key=value[,key=value]",
            "spo onedrive list|get|update [--filter <odata>]",
            "spo rename status [--current <prefix>] [--url <adminUrl>]",
            "spo rename start --new <prefix> [--current <prefix>] [--schedule <datetime>] [--adminUpn <upn>] [--interactive true|false] [--forceCredential true|false]"
        )
        "app"     = @(
            "app list|get|create|update|delete",
            "app redirect add|remove <appId|objectId> --uri <url> [--type spa|web|public]",
            "app secret list|add|remove <appId|objectId>",
            "app cert list|add|remove <appId|objectId>",
            "app perm list|add|remove <appId|objectId> [--type delegated|application] [--all]",
            "app consent",
            "app guide <appId|objectId> [--target react|node|dotnet|spa|daemon]"
        )
        "accessreview" = @(
            "accessreview list|get|create|update|delete",
            "accessreview instance list --def <definitionId>",
            "accessreview decision list --def <definitionId> --instance <instanceId>",
            "accessreview decision submit <decisionId> --def <definitionId> --instance <instanceId> --set key=value OR --json <payload>",
            "accessreview decision apply --def <definitionId> --instance <instanceId>",
            "accessreview history list|get|create|delete",
            "accessreview history instance list --id <historyDefinitionId>"
        )
        "graph"   = @(
            "graph cmdlets [--filter]",
            "graph perms [--type delegated|application] [--filter text]",
            "graph req <get|post|patch|put|delete> <path|url> [--body <json>] [--bodyFile <path>] [--headers <json>] [--beta|--v1|--auto] [--contentType <type>] [--out <file>]",
            "graph meta sync [--beta] [--v1] [--force]",
            "graph meta list [--type entityset|entity|action|function|enum|complex] [--filter text] [--beta]",
            "graph meta show <name> [--beta]",
            "graph meta paths <name> [--beta]",
            "graph meta diff [--type entityset|entity|action|function|enum|complex|property|nav] [--filter text] [--v1only] [--top <n>] [--json]",
            "graph list <path> [--filter <odata>] [--top <n>] [--select ...] [--expand ...] [--orderby ...] [--search text] [--beta]",
            "graph get <path> [--select ...] [--expand ...] [--beta]",
            "graph create <path> --json <payload> [--beta|--auto]",
            "graph update <path> --json <payload> [--beta|--auto]",
            "graph delete <path> [--force] [--beta|--auto]",
            "graph action <path> --json <payload> [--beta|--auto]",
            "graph batch --file <json> [--beta|--auto]"
        )
        "alias" = @(
            "alias list [--global|--local]",
            "alias get <name> [--global|--local]",
            "alias set <name> --value <command> [--global|--local]",
            "alias remove <name> [--global|--local]"
        )
        "preset" = @(
            "preset list",
            "preset get <name>",
            "preset set <name> --value <cmd1; cmd2; ...>",
            "preset remove <name>",
            "preset run <name> [args...]"
        )
        "manifest" = @(
            "manifest list [--type core|handlers|all]",
            "manifest sync [--type core|handlers|all]",
            "manifest set --type core|handlers --items a.ps1,b.ps1 OR --file <json>"
        )
        "forms"   = @(
            "forms open",
            "forms info",
            "forms admin get|update",
            "forms report list|run",
            "forms raw <get|post|patch|put|delete> <path> [--body <json>] [--bodyFile <path>] [--out <file>]",
            "forms excel tables --item <id>|--path <path> [--user <upn|id>]",
            "forms excel rows --item <id>|--path <path> [--table <name|id>] [--top <n>] [--skip <n>] [--json]",
            "forms excel watch --item <id>|--path <path> [--table <name|id>] [--interval <sec>] [--fromNow true|false] [--max <n>] [--json]",
            "forms flow list|get|runs|actions --env <environmentId> [--name <text>] [--workflowId <id>]"
        )
        "stream"  = @(
            "stream open",
            "stream list [--query <kql>] [--types mp4,mov,...] [--siteUrl <url>] [--path <url>] [--top <n>] [--from <n>] [--beta|--auto] [--json]",
            "stream search --query <kql> [--types ...] [--siteUrl <url>] [--path <url>] [--top <n>] [--from <n>] [--beta|--auto] [--json]",
            "stream file <list|get|create|update|delete|download|convert|preview|upload|copy|move|share> ... (uses file command args)"
        )
        "clipchamp" = @(
            "clipchamp open|info",
            "clipchamp list|search [--query <kql>] [--path <url>] [--siteUrl <url>]",
            "clipchamp project list|get|open|assets|exports [--path <projectFolder>]",
            "clipchamp file <list|get|create|update|delete|download|upload|share> ... (uses file command args)"
        )
        "copilot" = @(
            "copilot chat create",
            "copilot chat send <conversationId> --text <message> [--files url1,url2] [--useSearch 1,2] [--useSearchTop <n>] [--useRetrieve 1,2] [--useRetrieveTop <n>] [--tz <timezone>] [--web true|false] [--stream] [--ctx <text>] [--ctxFile <path>] [--ctxMax <n>] [--mail <id>] [--event <id>] [--meeting <id>] [--person <upn|id>] [--user <upn|id>] [--text]",
            "copilot chat stream <conversationId> --text <message> [--files url1,url2] [--useSearch 1,2] [--useSearchTop <n>] [--useRetrieve 1,2] [--useRetrieveTop <n>] [--tz <timezone>] [--web true|false] [--ctx <text>] [--ctxFile <path>] [--ctxMax <n>] [--mail <id>] [--event <id>] [--meeting <id>] [--person <upn|id>] [--user <upn|id>] [--text]",
            "copilot chat ask --text <message> [--files url1,url2] [--useSearch 1,2] [--useSearchTop <n>] [--useRetrieve 1,2] [--useRetrieveTop <n>] [--tz <timezone>] [--web true|false] [--stream] [--ctx <text>] [--ctxFile <path>] [--ctxMax <n>] [--mail <id>] [--event <id>] [--meeting <id>] [--person <upn|id>] [--user <upn|id>] [--text]",
            "copilot search --query <text> [--path <url>] [--paths url1,url2] [--pageSize <n>] [--metadata name,name] [--hits]",
            "copilot search list|open|download [--index <n>] [--out <file>]",
            "copilot search next --url <nextLink>",
            "copilot retrieve --query <text> --source sharePoint|oneDriveBusiness|externalItem [--max <n>] [--filter <odata>] [--metadata name,name] [--connections id1,id2] [--hits]",
            "copilot retrieve list|open|download [--index <n>] [--out <file>]",
            "copilot retrieve ask --query <text> --source <source> --prompt <text> [--top <n>] [--stream]"
        )
        "bookings" = @(
            "bookings business list|get|create|update|delete",
            "bookings service list|get|create|update|delete --business <id>",
            "bookings staff list|get|create|update|delete --business <id>",
            "bookings appointment list|get|create|update|delete --business <id>",
            "bookings customer list|get|create|update|delete --business <id>"
        )
        "orgx" = @(
            "orgx manager [--user <upn|id>]",
            "orgx reports [--user <upn|id>]",
            "orgx chain [--user <upn|id>] [--depth <n>] [--json]",
            "orgx tree [--user <upn|id>] [--depth <n>] [--max <n>] [--json]"
        )
        "whiteboard" = @(
            "whiteboard list [--user <upn|id>] [--path <path>]",
            "whiteboard get <itemId> OR --path <path>",
            "whiteboard download <itemId> --out <file> OR --path <path> --out <file>",
            "whiteboard share ... (same as: file share ...)"
        )
        "insights" = @(
            "insights list --type shared|trending|used [--user <upn|id>] [--top <n>]",
            "insights get <id> --type shared|trending|used [--user <upn|id>]"
        )
        "connections" = @(
            "connections open",
            "connections home",
            "connections site <list|get|...>",
            "connections news list|get|create|update|delete|publish --site <siteId|url|hostname:/path:>",
            "connections dashboard list|get|open|update|publish --site <siteId|url|hostname:/path:>"
        )
        "engage" = @(
            "engage open|info",
            "engage token set|show|clear [--value <token>|--file <path>]",
            "engage message list|post|delete [--group <id>] [--thread <id>] [--topic <name>] [--opengraph <id>] [--feed my|sent|received|private|algo] [--limit <n>]",
            "engage community list|get|create|update|delete|owners|group|members [--beta|--auto]",
            "engage community member add|remove <communityId> --user <upn|id>",
            "engage community owner add|remove <communityId> --user <upn|id>",
            "engage raw <method> <path> [--json <payload>] [--bodyFile <file>] [--token <token>] [--base <url>] [--out <file>]"
        )
        "lists" = @(
            "lists list|get|create|update|delete|delta --site <siteId|url|hostname:/path:>",
            "lists item list|get|create|update|delete --site <siteId> --list <listId>"
        )
        "learning" = @(
            "learning provider|content ... (aliases viva learning)",
            "learning activity list|get|create|delete [--user <upn|id>] [--beta|--auto]"
        )
        "loop" = @(
            "loop open|info",
            "loop list|search [--query <kql>]",
            "loop get|download|upload|create|update|delete|share <...> (uses file args)"
        )
        "sway" = @(
            "sway open [swayId|url]",
            "sway uri --command <cmd> [--param k=v[,k=v]] [--url <url>]"
        )
        "kaizala" = @(
            "kaizala (retired)"
        )
        "apps" = @(
            "apps list [--json]",
            "apps get <appId|name>",
            "apps open <appId|name>",
            "apps cli <appId|name> [--cmd <m365 prefix>] <args...>"
        )
        "m365" = @(
            "m365 status",
            "m365 login [interactive|device]",
            "m365 logout",
            "m365 request <get|post|patch|put|delete> <url|path> [--body <json>] [--bodyFile <path>] [--headers <json>] [--beta|--v1|--auto] [--out <file>]",
            "m365 search <...> (same as: search ...)",
            "m365 version",
            "m365 docs"
        )
        "m365cli" = @(
            "m365cli status|install|path",
            "m365cli source path|clone|update",
            "m365cli inventory [--refresh] [--area <name>] [--filter <text>] [--json]",
            "m365cli parity [--refresh] [--json]",
            "m365cli app list|set|remove|show|run",
            "m365cli run <m365 args...>",
            "m365cli <m365 args...> (pass-through)"
        )
        "module"  = @("module list|install|update|remove <name>")
        "webhook" = @(
            "webhook listen [--port <n>] [--out <file>] [--once true|false] [--prefix <url>]",
            "webhook start [--port <n>] [--out <file>] [--prefix <url>]",
            "webhook stop",
            "webhook status"
        )
        "tenant"  = @("/tenant show", "/tenant set prefix <value>", "/tenant set domain <value>", "/tenant set id <value>")
        "config"  = @("/config show", "/config get <path>", "/config set <path> <json-or-text>")
        "login"   = @("/login", "/login device")
    }

    if (-not $Topic) {
        Write-Host "Global commands (start with '/'):"
        $globalHelp | ForEach-Object { Write-Host ("  " + $_) }
        Write-Host ""
        Write-Host "Local commands (no '/' prefix):"
        $localHelp | ForEach-Object { Write-Host ("  " + $_) }
        Write-Host ""
        Write-Host "Use /help <topic> for details."
        return
    }

    $t = $Topic.TrimStart('/').ToLowerInvariant()
    if ($topics.ContainsKey($t)) {
        $topics[$t] | ForEach-Object { Write-Host ("  " + $_) }
    } else {
        Write-Warn "No help found for: $Topic"
    }
}




