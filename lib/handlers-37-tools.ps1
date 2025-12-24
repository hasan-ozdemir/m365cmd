# Handler: Tools
# Purpose: Tools command handlers.
function Get-M365CliInstallRoot {
    return (Join-Path $Paths.Tools "m365")
}


function Get-M365CliPath {
    $root = Get-M365CliInstallRoot
    $bin = Join-Path $root "node_modules/.bin"
    $cmd = if ($IsWindows) { "m365.cmd" } else { "m365" }
    $local = Join-Path $bin $cmd
    if (Test-Path $local) { return $local }
    $found = Get-Command -Name "m365" -ErrorAction SilentlyContinue
    if ($found) { return $found.Source }
    return $null
}


function Handle-M365CliCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: m365cli status|install|path|app|run|inventory|parity|source <args...> OR m365cli <m365 args...>"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    switch ($sub) {
        "status" {
            $path = Get-M365CliPath
            if (-not $path) {
                Write-Warn "CLI for Microsoft 365 not found. Use: m365cli install"
                return
            }
            Write-Host ("m365 path: " + $path)
            try {
                & $path "--version"
            } catch {
                Write-Warn "Unable to run m365 --version."
            }
        }
        "path" {
            $path = Get-M365CliPath
            if ($path) { Write-Host $path } else { Write-Warn "CLI for Microsoft 365 not found." }
        }
        "install" {
            $npm = Get-Command -Name "npm" -ErrorAction SilentlyContinue
            if (-not $npm) {
                Write-Warn "npm not found. Install Node.js first."
                return
            }
            Ensure-Directories
            $root = Get-M365CliInstallRoot
            if (-not (Test-Path $root)) {
                New-Item -ItemType Directory -Path $root -Force | Out-Null
            }
            Write-Info "Installing CLI for Microsoft 365 locally..."
            try {
                & $npm.Source "install" "--prefix" $root "@pnp/cli-microsoft365"
                Write-Info "Install completed."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "app" {
            Handle-M365CliAppCommand $rest
        }
        "source" {
            Handle-M365CliSourceCommand $rest
        }
        "inventory" {
            Handle-M365CliInventoryCommand $rest
        }
        "parity" {
            Handle-M365CliParityCommand $rest
        }
        "run" {
            $path = Get-M365CliPath
            if (-not $path) {
                Write-Warn "CLI for Microsoft 365 not found. Use: m365cli install"
                return
            }
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: m365cli run <m365 args...>"
                return
            }
            & $path @rest
        }
        default {
            $path = Get-M365CliPath
            if (-not $path) {
                Write-Warn "CLI for Microsoft 365 not found. Use: m365cli install"
                return
            }
            & $path @Args
        }
    }
}


function Get-M365CliAppMapPath {
    return (Join-Path $Paths.Data "m365cli.appmap.json")
}


function Load-M365CliAppMap {
    $path = Get-M365CliAppMapPath
    if (-not (Test-Path $path)) { return @{} }
    try {
        $obj = Get-Content -Raw -Path $path | ConvertFrom-Json
        if (-not $obj) { return @{} }
        $map = @{}
        foreach ($p in $obj.PSObject.Properties) { $map[$p.Name] = [string]$p.Value }
        return $map
    } catch {
        Write-Warn "App map is invalid. Recreating on next save."
        return @{}
    }
}


function Save-M365CliAppMap {
    param([hashtable]$Map)
    if (-not $Map) { $Map = @{} }
    Ensure-Directories
    $path = Get-M365CliAppMapPath
    $Map | ConvertTo-Json -Depth 4 | Set-Content -Path $path -Encoding ASCII
}


function Handle-M365CliAppCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: m365cli app list|set|remove|show|run"
        return
    }
    $action = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $map = Load-M365CliAppMap

    switch ($action) {
        "list" {
            if (-not $map -or $map.Keys.Count -eq 0) {
                Write-Info "No app mappings configured."
                Write-Host "Use: m365cli app set <app> --cmd <m365 command prefix>"
                return
            }
            $out = @()
            foreach ($k in ($map.Keys | Sort-Object)) {
                $out += [pscustomobject]@{ App = $k; Command = $map[$k] }
            }
            $out | Format-Table -AutoSize
        }
        "show" {
            $app = $rest | Select-Object -First 1
            if (-not $app) {
                Write-Warn "Usage: m365cli app show <app>"
                return
            }
            if ($map.ContainsKey($app)) {
                Write-Host $map[$app]
            } else {
                Write-Warn "App mapping not found."
            }
        }
        "set" {
            $app = $rest | Select-Object -First 1
            $parsed = Parse-NamedArgs ($rest | Select-Object -Skip 1)
            $cmd = Get-ArgValue $parsed.Map "cmd"
            if (-not $app -or -not $cmd) {
                Write-Warn "Usage: m365cli app set <app> --cmd <m365 command prefix>"
                return
            }
            $map[$app] = $cmd
            Save-M365CliAppMap $map
            Write-Info "App mapping saved."
        }
        "remove" {
            $app = $rest | Select-Object -First 1
            if (-not $app) {
                Write-Warn "Usage: m365cli app remove <app>"
                return
            }
            if ($map.ContainsKey($app)) {
                $map.Remove($app)
                Save-M365CliAppMap $map
                Write-Info "App mapping removed."
            } else {
                Write-Warn "App mapping not found."
            }
        }
        "run" {
            $app = $rest | Select-Object -First 1
            $args2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            if (-not $app -or -not $map.ContainsKey($app)) {
                Write-Warn "Usage: m365cli app run <app> <args...> (app must be mapped)"
                return
            }
            $path = Get-M365CliPath
            if (-not $path) {
                Write-Warn "CLI for Microsoft 365 not found. Use: m365cli install"
                return
            }
            $prefix = Split-Args $map[$app]
            & $path @prefix @args2
        }
        default {
            Write-Warn "Usage: m365cli app list|set|remove|show|run"
        }
    }
}


function Get-M365CliRepoPath {
    return (Join-Path $Paths.Tools "cli-microsoft365")
}


function Get-M365CliInventoryPath {
    return (Join-Path $Paths.Data "m365cli.inventory.json")
}


function Resolve-M365CliCommandsFile {
    param([string]$StartDir, [string]$RepoRoot)
    $dir = $StartDir
    while ($dir -and ($dir -like (Join-Path $RepoRoot "*"))) {
        $candidate = Join-Path $dir "commands.ts"
        if (Test-Path $candidate) { return $candidate }
        $parent = Split-Path -Parent $dir
        if ($parent -eq $dir) { break }
        $dir = $parent
    }
    return $null
}


function Parse-M365CliCommandsMap {
    param([string]$CommandsFile)
    $vars = @{}
    $map = @{}
    $lines = Get-Content -Path $CommandsFile
    foreach ($line in $lines) {
        if ($line -match '^\s*const\s+([A-Za-z0-9_]+)\s*(?::\s*string)?\s*=\s*''([^'']+)''') {
            $vars[$matches[1]] = $matches[2]
            continue
        }
        if ($line -match '^\s*const\s+([A-Za-z0-9_]+)\s*(?::\s*string)?\s*=\s*"([^"]+)"') {
            $vars[$matches[1]] = $matches[2]
            continue
        }
        if ($line -match '^\s*const\s+([A-Za-z0-9_]+)\s*(?::\s*string)?\s*=\s*`([^`]+)`') {
            $vars[$matches[1]] = $matches[2]
            continue
        }

        $key = $null
        $val = $null
        if ($line -match '^\s*([A-Z0-9_]+)\s*:\s*''([^'']+)''') {
            $key = $matches[1]
            $val = $matches[2]
        } elseif ($line -match '^\s*([A-Z0-9_]+)\s*:\s*"([^"]+)"') {
            $key = $matches[1]
            $val = $matches[2]
        } elseif ($line -match '^\s*([A-Z0-9_]+)\s*:\s*`([^`]+)`') {
            $key = $matches[1]
            $val = $matches[2]
        }
        if ($key) {
            foreach ($vk in $vars.Keys) {
                $pattern = [regex]::Escape('${' + $vk + '}')
                $val = $val -replace $pattern, $vars[$vk]
            }
            $map[$key] = $val
        }
    }
    return $map
}


function Get-M365CliCommandInventory {
    param([string]$RepoPath)
    $root = Join-Path $RepoPath "src\\m365"
    if (-not (Test-Path $root)) {
        throw "CLI repo not found at: $RepoPath"
    }

    $commandsFiles = Get-ChildItem -Path $root -Recurse -Filter "commands.ts"
    $items = @()
    foreach ($cf in $commandsFiles) {
        $map = Parse-M365CliCommandsMap $cf.FullName
        if (-not $map -or $map.Keys.Count -eq 0) { continue }
        $area = Split-Path -Leaf (Split-Path -Parent $cf.FullName)
        foreach ($key in $map.Keys) {
            $cmdString = $map[$key]
            if (-not $cmdString) { continue }
            $items += [pscustomobject]@{
                Command     = $cmdString
                Description = $null
                Area        = $area
                SourceFile  = $cf.FullName
            }
        }
    }

    return ($items | Sort-Object Command -Unique)
}


function Handle-M365CliSourceCommand {
    param([string[]]$InputArgs)
    $action = if ($InputArgs.Count -gt 0) { $InputArgs[0].ToLowerInvariant() } else { "" }
    $repo = Get-M365CliRepoPath
    switch ($action) {
        "path" {
            Write-Host $repo
        }
        "clone" {
            if (Test-Path $repo) {
                Write-Warn "Repo already exists: $repo"
                return
            }
            $git = Get-Command -Name "git" -ErrorAction SilentlyContinue
            if (-not $git) {
                Write-Warn "git not found. Install Git first."
                return
            }
            Ensure-Directories
            Write-Info "Cloning CLI for Microsoft 365 source..."
            & $git.Source "clone" "--depth" "1" "https://github.com/pnp/cli-microsoft365" $repo
        }
        "update" {
            if (-not (Test-Path $repo)) {
                Write-Warn "Repo not found. Use: m365cli source clone"
                return
            }
            $git = Get-Command -Name "git" -ErrorAction SilentlyContinue
            if (-not $git) {
                Write-Warn "git not found. Install Git first."
                return
            }
            Write-Info "Updating CLI source..."
            & $git.Source "-C" $repo "pull"
        }
        default {
            Write-Warn "Usage: m365cli source path|clone|update"
        }
    }
}


function Load-M365CliInventory {
    $path = Get-M365CliInventoryPath
    if (-not (Test-Path $path)) { return @() }
    try {
        $data = Get-Content -Raw -Path $path | ConvertFrom-Json
        return @($data)
    } catch {
        Write-Warn "Inventory file is invalid. Rebuilding required."
        return @()
    }
}


function Save-M365CliInventory {
    param([object[]]$Items)
    Ensure-Directories
    $path = Get-M365CliInventoryPath
    $Items | ConvertTo-Json -Depth 6 | Set-Content -Path $path -Encoding ASCII
}


function Handle-M365CliInventoryCommand {
    param([string[]]$InputArgs)
    $parsed = Parse-NamedArgs $InputArgs
    $refresh = Parse-Bool (Get-ArgValue $parsed.Map "refresh") $false
    $area = Get-ArgValue $parsed.Map "area"
    $filter = Get-ArgValue $parsed.Map "filter"
    $asJson = Parse-Bool (Get-ArgValue $parsed.Map "json") $false

    $repo = Get-M365CliRepoPath
    $inventory = @()

    if ($refresh -or -not (Test-Path (Get-M365CliInventoryPath))) {
        if (-not (Test-Path $repo)) {
            Write-Warn "CLI source not found. Use: m365cli source clone"
            return
        }
        Write-Info "Building CLI command inventory..."
        $inventory = Get-M365CliCommandInventory -RepoPath $repo
        Save-M365CliInventory $inventory
    } else {
        $inventory = Load-M365CliInventory
    }

    if ($area) {
        $inventory = $inventory | Where-Object { $_.Area -eq $area }
    }
    if ($filter) {
        $inventory = $inventory | Where-Object { $_.Command -like "*$filter*" }
    }

    if ($asJson) {
        $inventory | ConvertTo-Json -Depth 6
    } else {
        $inventory | Select-Object Command, Description, Area | Format-Table -AutoSize
        Write-Host ("Total commands: " + ($inventory | Measure-Object).Count)
    }
}


function Get-M365CmdLocalCommands {
    $path = Join-Path $PSScriptRoot "handlers-99-dispatch.ps1"
    if (-not (Test-Path $path)) { return @() }
    $cmds = @()
    foreach ($line in (Get-Content -Path $path)) {
        if ($line -match '^\s*\"([a-z0-9]+)\"\s*\{') {
            $cmds += $matches[1]
        }
    }
    return ($cmds | Sort-Object -Unique)
}


function Get-M365CmdGlobalCommands {
    $path = Join-Path $PSScriptRoot "handlers-99-dispatch.ps1"
    if (-not (Test-Path $path)) { return @() }
    $cmds = @()
    $inGlobal = $false
    foreach ($line in (Get-Content -Path $path)) {
        if ($line -match 'function\s+Handle-GlobalCommand') { $inGlobal = $true; continue }
        if ($inGlobal -and $line -match 'function\s+Handle-LocalCommand') { break }
        if ($inGlobal -and $line -match '^\s*\"([a-z0-9]+)\"\s*\{') {
            $cmds += $matches[1]
        }
    }
    return ($cmds | Sort-Object -Unique)
}


function Get-M365CliParityPath {
    return (Join-Path $Paths.Data "m365cli.parity.json")
}


function Handle-M365CliParityCommand {
    param([string[]]$InputArgs)
    $parsed = Parse-NamedArgs $InputArgs
    $refresh = Parse-Bool (Get-ArgValue $parsed.Map "refresh") $false
    $asJson = Parse-Bool (Get-ArgValue $parsed.Map "json") $false

    $inventory = @()
    if ($refresh -or -not (Test-Path (Get-M365CliInventoryPath))) {
        $repo = Get-M365CliRepoPath
        if (-not (Test-Path $repo)) {
            Write-Warn "CLI source not found. Use: m365cli source clone"
            return
        }
        $inventory = Get-M365CliCommandInventory -RepoPath $repo
        Save-M365CliInventory $inventory
    } else {
        $inventory = Load-M365CliInventory
    }

    if (-not $inventory -or $inventory.Count -eq 0) {
        Write-Warn "Inventory is empty. Use: m365cli inventory --refresh"
        return
    }

    $cliAreas = @($inventory | ForEach-Object { ($_.Command -split '\s+')[0] } | Sort-Object -Unique)
    $local = Get-M365CmdLocalCommands
    $global = Get-M365CmdGlobalCommands

    $areaMap = @{
        "entra" = @("user","group","role","ca","device","authmethod","risk")
        "aad" = @("user","group","role","ca","device","authmethod","risk")
        "teams" = @("teams","chat","channelmsg","teamsapp","teamsappinst","teamstab")
        "spo" = @("spo","site","splist","spage","spcolumn","spctype","spperm","file","drive")
        "onedrive" = @("onedrive","file","drive")
        "graph" = @("graph","search")
        "outlook" = @("outlook","mail","calendar","contacts","people")
        "planner" = @("planner")
        "todo" = @("todo")
        "viva" = @("viva","insights","connections","engage","learning")
        "pp" = @("pp","powerapps","powerautomate","powerpages")
        "booking" = @("bookings")
    }

    $rows = @()
    foreach ($area in $cliAreas) {
        $coveredBy = @()
        if ($local -contains $area -or $global -contains $area) {
            $coveredBy += $area
        }
        if ($areaMap.ContainsKey($area)) {
            foreach ($mapped in $areaMap[$area]) {
                if ($local -contains $mapped) { $coveredBy += $mapped }
            }
        }
        $status = if ($coveredBy.Count -gt 0) { "partial" } else { "missing" }
        $rows += [pscustomobject]@{
            Area      = $area
            Status    = $status
            CoveredBy = ($coveredBy | Sort-Object -Unique) -join ","
        }
    }

    Ensure-Directories
    $rows | ConvertTo-Json -Depth 4 | Set-Content -Path (Get-M365CliParityPath) -Encoding ASCII

    if ($asJson) {
        $rows | ConvertTo-Json -Depth 4
    } else {
        $rows | Sort-Object Status, Area | Format-Table -AutoSize
        Write-Host ("Areas checked: " + $rows.Count)
        Write-Host ("Parity report saved: " + (Get-M365CliParityPath))
    }
}


function Handle-M365Command {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: m365 status|login|logout|request|search|version|docs"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    switch ($sub) {
        "status" {
            Show-Status
        }
        "login" {
            Invoke-Login ($rest | Select-Object -First 1)
        }
        "logout" {
            Invoke-Logout
        }
        "request" {
            if (-not $rest -or $rest.Count -lt 2) {
                Write-Warn "Usage: m365 request <get|post|patch|put|delete> <url|path> [--body <json>] [--bodyFile <path>] [--headers <json>] [--beta|--v1|--auto] [--out <file>]"
                return
            }
            Handle-GraphCommand (@("req") + $rest)
        }
        "search" {
            Handle-SearchCommand $rest
        }
        "version" {
            Write-Host "m365cmd (native CLI port)"
        }
        "docs" {
            Write-Host "Use /help or /help <topic> for command reference."
        }
        default {
            Write-Warn "Unknown m365 command. Use: m365 status|login|logout|request|search|version|docs"
        }
    }
}

