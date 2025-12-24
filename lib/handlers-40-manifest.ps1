# Handler: Manifest
# Purpose: Manifest management for loader order and sync.
function Get-ManifestPath {
    param([string]$Type)
    $t = if ($Type) { $Type.ToLowerInvariant() } else { "" }
    if ($t -in @("core","cores")) { return (Join-Path $PSScriptRoot "core.manifest.json") }
    if ($t -in @("handler","handlers")) { return (Join-Path $PSScriptRoot "handlers.manifest.json") }
    return $null
}


function Get-ManifestFiles {
    param([string]$Type)
    $t = if ($Type) { $Type.ToLowerInvariant() } else { "" }
    if ($t -in @("core","cores")) {
        return Get-ChildItem -Path $PSScriptRoot -Filter "core-*.ps1" |
            Where-Object { $_.Name -notin @("core.ps1","core-loader.ps1") } |
            Sort-Object Name | Select-Object -ExpandProperty Name
    }
    if ($t -in @("handler","handlers")) {
        return Get-ChildItem -Path $PSScriptRoot -Filter "handlers-*.ps1" |
            Where-Object { $_.Name -notin @("handlers.ps1","handlers-loader.ps1") } |
            Sort-Object Name | Select-Object -ExpandProperty Name
    }
    return @()
}


function Read-ManifestList {
    param([string]$Type)
    $path = Get-ManifestPath $Type
    if (-not $path -or -not (Test-Path $path)) { return @() }
    try {
        return (Get-Content -Raw -Path $path | ConvertFrom-Json)
    } catch {
        Write-Warn "Manifest JSON is invalid."
        return @()
    }
}


function Write-ManifestList {
    param([string]$Type, [string[]]$List)
    $path = Get-ManifestPath $Type
    if (-not $path) { return }
    $json = $List | ConvertTo-Json -Depth 3
    Set-Content -Path $path -Value $json -Encoding ASCII
}


function Handle-ManifestCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: manifest list|sync|set [--type core|handlers]"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $type = Get-ArgValue $parsed.Map "type"
    if (-not $type) { $type = "handlers" }
    $t = $type.ToLowerInvariant()

    switch ($sub) {
        "list" {
            if ($t -eq "all") {
                foreach ($kind in @("core","handlers")) {
                    Write-Host ("[" + $kind + "]")
                    $list = Read-ManifestList $kind
                    if (-not $list -or $list.Count -eq 0) {
                        Write-Host "  (empty)"
                        continue
                    }
                    $i = 1
                    foreach ($item in $list) {
                        Write-Host ("  {0,3}. {1}" -f $i, $item)
                        $i++
                    }
                }
                return
            }
            $list = Read-ManifestList $type
            if (-not $list -or $list.Count -eq 0) { Write-Info "No manifest entries."; return }
            $i = 1
            foreach ($item in $list) { Write-Host ("{0,3}. {1}" -f $i, $item); $i++ }
        }
        "sync" {
            if ($t -eq "all") {
                foreach ($kind in @("core","handlers")) {
                    $files = Get-ManifestFiles $kind
                    if (-not $files -or $files.Count -eq 0) { continue }
                    Write-ManifestList $kind $files
                }
                Write-Info "Manifests synced."
                return
            }
            $files = Get-ManifestFiles $type
            if (-not $files -or $files.Count -eq 0) { Write-Warn "No files found for manifest."; return }
            Write-ManifestList $type $files
            Write-Info "Manifest synced."
        }
        "set" {
            if ($t -eq "all") {
                Write-Warn "Use --type core or --type handlers with set."
                return
            }
            $itemsRaw = Get-ArgValue $parsed.Map "items"
            $file = Get-ArgValue $parsed.Map "file"
            $items = @()
            if ($file) {
                if (-not (Test-Path $file)) { Write-Warn "File not found."; return }
                try { $items = Get-Content -Raw -Path $file | ConvertFrom-Json } catch { Write-Warn "Invalid JSON file."; return }
            } elseif ($itemsRaw) {
                $items = Parse-CommaList $itemsRaw
            }
            if (-not $items -or $items.Count -eq 0) {
                Write-Warn "Usage: manifest set --type core|handlers --items a.ps1,b.ps1 OR --file <json>"
                return
            }
            $all = Get-ManifestFiles $type
            $valid = @()
            foreach ($i in $items) {
                if ($all -contains $i) { $valid += $i } else { Write-Warn ("Unknown file: " + $i) }
            }
            $remaining = @($all | Where-Object { $valid -notcontains $_ })
            $final = @($valid + $remaining)
            Write-ManifestList $type $final
            Write-Info "Manifest order updated."
        }
        default {
            Write-Warn "Usage: manifest list|sync|set [--type core|handlers]"
        }
    }
}

