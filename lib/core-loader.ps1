param(
    [string]$Root,
    [switch]$SyncManifest
)

$rootPath = if ($Root) { $Root } else { $PSScriptRoot }
$manifest = Join-Path $rootPath "core.manifest.json"
$warn = {
    param([string]$Message)
    if (Get-Command Write-Warn -ErrorAction SilentlyContinue) {
        Write-Warn $Message
    } else {
        Write-Host ("WARN: " + $Message) -ForegroundColor Yellow
    }
}

$list = @()
if (Test-Path $manifest) {
    try {
        $list = Get-Content -Raw -Path $manifest | ConvertFrom-Json
    } catch {
        & $warn "core.manifest.json is invalid. Falling back to file scan."
    }
}
if (-not $list -or $list.Count -eq 0) {
    $list = Get-ChildItem -Path $rootPath -Filter "core-*.ps1" |
        Where-Object { $_.Name -notin @("core.ps1","core-loader.ps1") } |
        Sort-Object Name |
        Select-Object -ExpandProperty Name
}
if ($SyncManifest) {
    $current = Get-ChildItem -Path $rootPath -Filter "core-*.ps1" |
        Where-Object { $_.Name -notin @("core.ps1","core-loader.ps1") } |
        Sort-Object Name |
        Select-Object -ExpandProperty Name
    $missing = @($list | Where-Object { $current -notcontains $_ })
    $extra = @($current | Where-Object { $list -notcontains $_ })
    if ($missing.Count -gt 0 -or $extra.Count -gt 0) {
        $newList = @()
        foreach ($f in $list) { if ($current -contains $f) { $newList += $f } }
        foreach ($f in $current) { if ($newList -notcontains $f) { $newList += $f } }
        try {
            $json = $newList | ConvertTo-Json -Depth 3
            Set-Content -Path $manifest -Value $json -Encoding ASCII
        } catch {
            & $warn "Failed to sync core.manifest.json."
        }
        $list = $newList
    }
}

$verbose = $env:M365CMD_LOADER_VERBOSE
$trace = $env:M365CMD_LOADER_TRACE
$missing = @()
foreach ($f in $list) {
    $path = Join-Path $rootPath $f
    if (Test-Path $path) {
        if ($trace) { Write-Host ("[core] " + $f) }
        . $path
    } else {
        $missing += $f
    }
}
if ($verbose) {
    Write-Host ("[core] loaded " + $list.Count + " files")
}
if ($missing.Count -gt 0) {
    & $warn ("core.manifest.json references missing files: " + ($missing -join ", "))
}
