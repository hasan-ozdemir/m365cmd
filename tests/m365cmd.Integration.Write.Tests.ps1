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

function Invoke-GraphStrict {
    param(
        [string]$Method,
        [string]$Uri,
        [object]$Body,
        [switch]$Beta,
        [switch]$AllowNull
    )
    $err = $null
    $resp = Invoke-GraphRequest -Method $Method -Uri $Uri -Body $Body -Beta:$Beta -SuppressError -ErrorRef ([ref]$err)
    if ($err) {
        throw ("Graph request failed: " + $err.Exception.Message + " (" + $Method + " " + $Uri + ")")
    }
    if (-not $AllowNull -and $resp -eq $null) {
        throw ("Graph request returned null (" + $Method + " " + $Uri + ")")
    }
    return $resp
}

Describe "Write operations" -Tag "write" {
    BeforeAll {
        $repoRoot = Get-TestRoot
        $env:M365CMD_ROOT = Join-Path ([System.IO.Path]::GetTempPath()) ("m365cmd-tests-" + [guid]::NewGuid().ToString("n"))
        New-Item -ItemType Directory -Path $env:M365CMD_ROOT -Force | Out-Null
        . (Join-Path $repoRoot "lib\core.ps1")
        . (Join-Path $repoRoot "lib\handlers.ps1")
        . (Join-Path $repoRoot "lib\repl.ps1")
        $global:Config = Normalize-Config (Get-DefaultConfig)
    }

    It "Group create/update/delete" {
        if (-not (Ensure-GraphOrSkip)) { return }
        $id = $null
        $nick = ("m365cmdtest" + (Get-Random -Minimum 10000 -Maximum 99999))
        $body = @{ displayName = ("m365cmd-test-group-" + [guid]::NewGuid().ToString("n").Substring(0,8)); mailEnabled = $false; securityEnabled = $true; mailNickname = $nick }
        try {
            $grp = Invoke-GraphStrict -Method "POST" -Uri "/groups" -Body $body
            $id = $grp.id
            Invoke-GraphStrict -Method "PATCH" -Uri ("/groups/" + $id) -Body @{ description = "m365cmd test group" } -AllowNull | Out-Null
        } finally {
            if ($id) {
                Invoke-GraphStrict -Method "DELETE" -Uri ("/groups/" + $id) -AllowNull | Out-Null
            }
        }
    }

    It "Application create/update/delete" {
        if (-not (Ensure-GraphOrSkip)) { return }
        $id = $null
        $body = @{ displayName = ("m365cmd-test-app-" + [guid]::NewGuid().ToString("n").Substring(0,8)) }
        try {
            $app = Invoke-GraphStrict -Method "POST" -Uri "/applications" -Body $body
            $id = $app.id
            Invoke-GraphStrict -Method "PATCH" -Uri ("/applications/" + $id) -Body @{ notes = "m365cmd test" } -AllowNull | Out-Null
        } finally {
            if ($id) {
                Invoke-GraphStrict -Method "DELETE" -Uri ("/applications/" + $id) -AllowNull | Out-Null
            }
        }
    }

    It "Drive folder create/update/delete" {
        if (-not (Ensure-GraphOrSkip)) { return }
        $id = $null
        $name = "m365cmd-test-folder-" + [guid]::NewGuid().ToString("n").Substring(0,8)
        $body = @{ name = $name; folder = @{}; "@microsoft.graph.conflictBehavior" = "rename" }
        try {
            $item = Invoke-GraphStrict -Method "POST" -Uri "/me/drive/root/children" -Body $body
            $id = $item.id
            Invoke-GraphStrict -Method "PATCH" -Uri ("/me/drive/items/" + $id) -Body @{ name = ($name + "-upd") } -AllowNull | Out-Null
        } finally {
            if ($id) {
                Invoke-GraphStrict -Method "DELETE" -Uri ("/me/drive/items/" + $id) -AllowNull | Out-Null
            }
        }
    }

    It "To Do list create/update/delete" {
        if (-not (Ensure-GraphOrSkip)) { return }
        $id = $null
        $body = @{ displayName = ("m365cmd-test-todo-" + [guid]::NewGuid().ToString("n").Substring(0,8)) }
        try {
            $list = Invoke-GraphStrict -Method "POST" -Uri "/me/todo/lists" -Body $body
            $id = $list.id
            Invoke-GraphStrict -Method "PATCH" -Uri ("/me/todo/lists/" + $id) -Body @{ displayName = ($body.displayName + "-upd") } -AllowNull | Out-Null
        } finally {
            if ($id) {
                Invoke-GraphStrict -Method "DELETE" -Uri ("/me/todo/lists/" + $id) -AllowNull | Out-Null
            }
        }
    }

    It "Calendar event create/update/delete" {
        if (-not (Ensure-GraphOrSkip)) { return }
        $id = $null
        $start = (Get-Date).AddHours(1).ToString("s")
        $end = (Get-Date).AddHours(2).ToString("s")
        $body = @{ subject = ("m365cmd test event " + [guid]::NewGuid().ToString("n").Substring(0,6)); start = @{ dateTime = $start; timeZone = "UTC" }; end = @{ dateTime = $end; timeZone = "UTC" } }
        try {
            $event = Invoke-GraphStrict -Method "POST" -Uri "/me/events" -Body $body
            $id = $event.id
            Invoke-GraphStrict -Method "PATCH" -Uri ("/me/events/" + $id) -Body @{ subject = ($body.subject + "-upd") } -AllowNull | Out-Null
        } finally {
            if ($id) {
                Invoke-GraphStrict -Method "DELETE" -Uri ("/me/events/" + $id) -AllowNull | Out-Null
            }
        }
    }
}
