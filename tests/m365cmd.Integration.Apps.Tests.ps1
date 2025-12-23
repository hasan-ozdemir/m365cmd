function Get-TestRoot {
    $repoRoot = (Get-Location).Path
    if (-not (Test-Path (Join-Path $repoRoot "lib"))) {
        $repoRoot = Split-Path -Parent $repoRoot
    }
    if (-not (Test-Path (Join-Path $repoRoot "lib"))) {
        throw "Repo root not found. Run tests from repo root."
    }
    return $repoRoot
}

function Ensure-GraphOrSkip {
    $ctx = Get-MgContextSafe
    if (-not $ctx) {
        Set-ItResult -Skipped -Because "Not connected to Microsoft Graph. Run /login first."
        return $false
    }
    return $true
}

function Invoke-PortalProbe {
    param([string]$Url)
    try {
        $resp = Invoke-WebRequest -Uri $Url -Method Head -SkipHttpErrorCheck -TimeoutSec 15
        return [int]$resp.StatusCode
    } catch {
        try {
            $resp = Invoke-WebRequest -Uri $Url -Method Get -SkipHttpErrorCheck -TimeoutSec 15
            return [int]$resp.StatusCode
        } catch {
            return 0
        }
    }
}

function Invoke-GraphStrict {
    param(
        [string]$Method,
        [string]$Uri,
        [object]$Body,
        [switch]$Beta
    )
    $err = $null
    $resp = Invoke-GraphRequest -Method $Method -Uri $Uri -Body $Body -Beta:$Beta -SuppressError -ErrorRef ([ref]$err)
    if ($resp -eq $null) {
        $msg = if ($err -and $err.Exception) { $err.Exception.Message } else { "Unknown error" }
        throw ("Graph request failed: " + $msg)
    }
    return $resp
}

Describe "Integration app smoke tests" -Tag "integration" {
    BeforeAll {
        $repoRoot = Get-TestRoot
        $env:M365CMD_ROOT = Join-Path ([System.IO.Path]::GetTempPath()) ("m365cmd-tests-" + [guid]::NewGuid().ToString("n"))
        New-Item -ItemType Directory -Path $env:M365CMD_ROOT -Force | Out-Null
        . (Join-Path $repoRoot "lib\core.ps1")
        . (Join-Path $repoRoot "lib\handlers.ps1")
        . (Join-Path $repoRoot "lib\repl.ps1")
        $global:Config = Normalize-Config (Get-DefaultConfig)
    }

    It "Admin" {
        if (-not (Ensure-GraphOrSkip)) { return }
        (Invoke-GraphStrict -Method "GET" -Uri "/organization") | Should -Not -BeNullOrEmpty
    }

    It "Connections" {
        Invoke-PortalProbe "https://www.microsoft365.com/" | Should -BeGreaterThan 0
    }

    It "Bookings" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/solutions/bookingBusinesses?`$top=1" | Out-Null
    }

    It "Calendar" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/events?`$top=1" | Out-Null
    }

    It "Excel" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/drive/root" | Out-Null
    }

    It "Forms" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/admin/forms" -Beta | Out-Null
    }

    It "Clipchamp" {
        Invoke-PortalProbe "https://app.clipchamp.com/" | Should -BeGreaterThan 0
    }

    It "Loop" {
        Invoke-PortalProbe "https://loop.microsoft.com/" | Should -BeGreaterThan 0
    }

    It "Insights" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/insights/shared?`$top=1" | Out-Null
    }

    It "Engage" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/employeeExperience/communities?`$top=1" -Beta | Out-Null
    }

    It "Kaizala" {
        Invoke-PortalProbe "https://web.kaiza.la/" | Should -BeGreaterThan 0
    }

    It "Learning" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/employeeExperience/learningProviders" | Out-Null
    }

    It "Lists" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/sites?search=*" | Out-Null
    }

    It "OneDrive" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/drive" | Out-Null
    }

    It "OneNote" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/onenote/notebooks?`$top=1" | Out-Null
    }

    It "Outlook" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/messages?`$top=1" | Out-Null
    }

    It "People" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/people?`$top=1" | Out-Null
    }

    It "Planner" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/planner/plans" | Out-Null
    }

    It "Power Apps" {
        Invoke-PortalProbe "https://make.powerapps.com/" | Should -BeGreaterThan 0
    }

    It "Power Automate" {
        Invoke-PortalProbe "https://make.powerautomate.com/" | Should -BeGreaterThan 0
    }

    It "Power Pages" {
        Invoke-PortalProbe "https://make.powerpages.microsoft.com/" | Should -BeGreaterThan 0
    }

    It "PowerPoint" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/drive/root/children?`$top=1" | Out-Null
    }

    It "Purview" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/compliance/ediscovery/cases?`$top=1" -Beta | Out-Null
    }

    It "Security" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/security/alerts?`$top=1" | Out-Null
    }

    It "SharePoint" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/sites?search=*" | Out-Null
    }

    It "Stream" {
        if (-not (Ensure-GraphOrSkip)) { return }
        $body = @{ requests = @(@{ entityTypes = @("driveItem"); query = @{ queryString = "filetype:mp4" }; from = 0; size = 1 }) }
        Invoke-GraphStrict -Method "POST" -Uri "/search/query" -Body $body | Out-Null
    }

    It "Sway" {
        Invoke-PortalProbe "https://sway.office.com/" | Should -BeGreaterThan 0
    }

    It "Teams" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/chats?`$top=1" | Out-Null
    }

    It "To Do" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/todo/lists?`$top=1" | Out-Null
    }

    It "Visio" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/drive/root/children?`$top=1" | Out-Null
    }

    It "Viva" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/employeeExperience/learningProviders" | Out-Null
    }

    It "Org Explorer" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/manager" | Out-Null
    }

    It "Copilot" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "POST" -Uri "/copilot/conversations" -Beta | Out-Null
    }

    It "Whiteboard" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/drive/root/children?`$top=1" | Out-Null
    }

    It "Word" {
        if (-not (Ensure-GraphOrSkip)) { return }
        Invoke-GraphStrict -Method "GET" -Uri "/me/drive/root/children?`$top=1" | Out-Null
    }
}
