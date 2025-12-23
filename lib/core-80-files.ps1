# Core: Files
# Purpose: Files shared utilities.
function Invoke-DriveUpload {
    param(
        [string]$Base,
        [string]$LocalPath,
        [string]$DestPath,
        [switch]$Beta
    )
    if (-not $Base) { return }
    if (-not $LocalPath -or -not (Test-Path $LocalPath)) {
        Write-Warn "Local file not found."
        return
    }
    $fileName = [System.IO.Path]::GetFileName($LocalPath)
    $dest = Normalize-DrivePath $DestPath
    $targetPath = if ($dest) { $dest } else { $fileName }
    if ($dest -and ($dest.EndsWith("/") -or $dest.EndsWith("\"))) {
        $targetPath = ($dest.TrimEnd("/", "\") + "/" + $fileName)
    }

    $size = (Get-Item $LocalPath).Length
    $maxSimple = 250MB
    if ($size -le $maxSimple) {
        $bytes = [System.IO.File]::ReadAllBytes($LocalPath)
        $uri = $Base + "/root:/" + $targetPath + ":/content"
        $resp = Invoke-GraphRequest -Method "PUT" -Uri $uri -Body $bytes -ContentType "application/octet-stream" -Beta:$Beta
        if ($resp) {
            Write-Info "Upload completed."
        }
        return
    }

    $sessionUri = $Base + "/root:/" + $targetPath + ":/createUploadSession"
    $body = @{
        item = @{
            "@microsoft.graph.conflictBehavior" = "rename"
            name = $fileName
        }
    }
    $session = Invoke-GraphRequest -Method "POST" -Uri $sessionUri -Body $body -Beta:$Beta
    if (-not $session -or -not $session.uploadUrl) {
        Write-Err "Failed to create upload session."
        return
    }

    $uploadUrl = $session.uploadUrl
    $chunkSize = 10MB
    $fs = [System.IO.File]::OpenRead($LocalPath)
    try {
        $buffer = New-Object byte[] $chunkSize
        $offset = 0
        while (($read = $fs.Read($buffer, 0, $chunkSize)) -gt 0) {
            $end = $offset + $read - 1
            $chunk = if ($read -eq $buffer.Length) { $buffer } else { $buffer[0..($read - 1)] }
            $headers = @{
                "Content-Length" = $read
                "Content-Range"  = ("bytes {0}-{1}/{2}" -f $offset, $end, $size)
            }
            try {
                Invoke-RestMethod -Method Put -Uri $uploadUrl -Headers $headers -Body $chunk | Out-Null
            } catch {
                Write-Err $_.Exception.Message
                return
            }
            $offset += $read
        }
        Write-Info "Upload completed."
    } finally {
        $fs.Dispose()
    }
}



function Get-ServicePlanAutomationInfo {
    param([string]$PlanName)
    if (-not $PlanName) { return $null }
    $name = $PlanName.ToUpperInvariant()
    $info = [ordered]@{
        Category     = "Other"
        Modules      = @("Microsoft.Graph")
        Capabilities = @("User/license reporting, directory objects, basic settings")
    }

    if ($name -match "EXCHANGE") {
        $info.Category = "Exchange Online"
        $info.Modules = @("ExchangeOnlineManagement", "Microsoft.Graph")
        $info.Capabilities = @("Mailbox settings", "Mailbox permissions", "Distribution groups", "Mail flow rules")
    } elseif ($name -match "SHAREPOINT") {
        $info.Category = "SharePoint Online"
        $info.Modules = @("Microsoft.Online.SharePoint.PowerShell", "Microsoft.Graph")
        $info.Capabilities = @("Site provisioning", "Permissions & sharing", "Storage quotas")
    } elseif ($name -match "ONEDRIVE") {
        $info.Category = "OneDrive for Business"
        $info.Modules = @("Microsoft.Online.SharePoint.PowerShell", "Microsoft.Graph")
        $info.Capabilities = @("Personal site provisioning", "Sharing policies", "Storage settings")
    } elseif ($name -match "TEAMS") {
        $info.Category = "Microsoft Teams"
        $info.Modules = @("MicrosoftTeams", "Microsoft.Graph")
        $info.Capabilities = @("Teams/channels", "Policies", "Meetings & voice settings")
    } elseif ($name -match "YAMMER|VIVA") {
        $info.Category = "Viva/Yammer"
        $info.Modules = @("Microsoft.Graph")
        $info.Capabilities = @("Communities and group management (limited)")
    } elseif ($name -match "POWERBI") {
        $info.Category = "Power BI"
        $info.Modules = @("Microsoft.Graph")
        $info.Capabilities = @("Workspaces, users, capacity (varies by API)")
    } elseif ($name -match "INTUNE|MOBILITY|MDM") {
        $info.Category = "Intune / Endpoint Manager"
        $info.Modules = @("Microsoft.Graph", "Microsoft.Graph.Beta")
        $info.Capabilities = @("Device enrollment", "Compliance policies", "App protection policies")
    } elseif ($name -match "PROJECT|PLANNER") {
        $info.Category = "Planner/Project"
        $info.Modules = @("Microsoft.Graph")
        $info.Capabilities = @("Plans, tasks, assignments")
    } elseif ($name -match "OFFICE|PROPLUS") {
        $info.Category = "Microsoft 365 Apps"
        $info.Modules = @("Microsoft.Graph")
        $info.Capabilities = @("License assignments, basic user settings")
    }

    return $info
}



