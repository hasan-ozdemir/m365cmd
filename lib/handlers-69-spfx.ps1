# Handler: SPFx
# Purpose: SharePoint Framework helpers.
function Get-SpfxProjectInfo {
    $pkgPath = Join-Path (Get-Location) "package.json"
    if (-not (Test-Path $pkgPath)) { return $null }
    try {
        $pkg = Get-Content -Raw -Path $pkgPath | ConvertFrom-Json
    } catch {
        return $null
    }
    $deps = @{}
    if ($pkg.dependencies) {
        foreach ($p in $pkg.dependencies.PSObject.Properties) { $deps[$p.Name] = $p.Value }
    }
    if ($pkg.devDependencies) {
        foreach ($p in $pkg.devDependencies.PSObject.Properties) { $deps[$p.Name] = $p.Value }
    }
    $spfx = $deps.Keys | Where-Object { $_ -like "@microsoft/sp-*" }
    return [pscustomobject]@{
        Package = $pkg
        SpfxDeps = $spfx
        HasSpfx  = ($spfx.Count -gt 0)
        Path     = $pkgPath
    }
}


function Handle-SpfxCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: spfx doctor|package|project <args...>"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "doctor" {
            $node = Get-Command -Name "node" -ErrorAction SilentlyContinue
            $npm = Get-Command -Name "npm" -ErrorAction SilentlyContinue
            $yo = Get-Command -Name "yo" -ErrorAction SilentlyContinue
            $gulp = Get-Command -Name "gulp" -ErrorAction SilentlyContinue
            $info = Get-SpfxProjectInfo
            Write-Host ("Node     : " + $(if ($node) { "ok" } else { "missing" }))
            Write-Host ("npm      : " + $(if ($npm) { "ok" } else { "missing" }))
            Write-Host ("yo       : " + $(if ($yo) { "ok" } else { "missing" }))
            Write-Host ("gulp     : " + $(if ($gulp) { "ok" } else { "missing" }))
            if ($info -and $info.HasSpfx) {
                Write-Host ("SPFx deps: " + ($info.SpfxDeps -join ", "))
            } else {
                Write-Host "SPFx deps: not detected"
            }
        }
        "package" {
            $action = if ($rest.Count -gt 0) { $rest[0].ToLowerInvariant() } else { "" }
            if ($action -ne "generate") {
                Write-Warn "Usage: spfx package generate"
                return
            }
            $npx = Get-Command -Name "npx" -ErrorAction SilentlyContinue
            if (-not $npx) {
                Write-Warn "npx not found. Install Node.js first."
                return
            }
            try {
                & $npx.Source "gulp" "bundle" "--ship"
                & $npx.Source "gulp" "package-solution" "--ship"
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "project" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: spfx project doctor|rename|upgrade|externalize|github|azuredevops|permissions"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $parsed2 = Parse-NamedArgs $rest2
            switch ($action) {
                "doctor" {
                    Handle-SpfxCommand @("doctor")
                }
                "rename" {
                    $name = Get-ArgValue $parsed2.Map "name"
                    if (-not $name) { $name = $parsed2.Positionals | Select-Object -First 1 }
                    if (-not $name) {
                        Write-Warn "Usage: spfx project rename --name <newName>"
                        return
                    }
                    $info = Get-SpfxProjectInfo
                    if (-not $info) {
                        Write-Warn "package.json not found."
                        return
                    }
                    $pkg = $info.Package
                    $pkg.name = $name
                    $pkg | ConvertTo-Json -Depth 6 | Set-Content -Path $info.Path -Encoding ASCII
                    $solutionPath = Join-Path (Join-Path (Split-Path -Parent $info.Path) "config") "package-solution.json"
                    if (Test-Path $solutionPath) {
                        try {
                            $sol = Get-Content -Raw -Path $solutionPath | ConvertFrom-Json
                            if ($sol.solution) { $sol.solution.name = $name }
                            $sol | ConvertTo-Json -Depth 6 | Set-Content -Path $solutionPath -Encoding ASCII
                        } catch {}
                    }
                    Write-Info "Project renamed."
                }
                "upgrade" {
                    $version = Get-ArgValue $parsed2.Map "version"
                    if (-not $version) { $version = Get-ArgValue $parsed2.Map "to" }
                    if (-not $version) {
                        Write-Warn "Usage: spfx project upgrade --version <spfxVersion>"
                        return
                    }
                    $info = Get-SpfxProjectInfo
                    if (-not $info) {
                        Write-Warn "package.json not found."
                        return
                    }
                    $pkg = $info.Package
                    $updated = @()
                    foreach ($section in @("dependencies","devDependencies")) {
                        $deps = $pkg.$section
                        if ($deps) {
                            foreach ($p in @($deps.PSObject.Properties)) {
                                if ($p.Name -like "@microsoft/sp-*") {
                                    $deps.$($p.Name) = $version
                                    $updated += ($section + ":" + $p.Name)
                                }
                            }
                        }
                    }
                    $pkg | ConvertTo-Json -Depth 10 | Set-Content -Path $info.Path -Encoding ASCII
                    if ($updated.Count -gt 0) {
                        Write-Info ("Updated " + $updated.Count + " SPFx dependencies to " + $version)
                    } else {
                        Write-Warn "No SPFx dependencies found to update."
                    }
                }
                "externalize" {
                    $cdn = Get-ArgValue $parsed2.Map "cdnBasePath"
                    $includeAssetsRaw = Get-ArgValue $parsed2.Map "includeClientSideAssets"
                    if (-not $cdn) {
                        Write-Warn "Usage: spfx project externalize --cdnBasePath <url> [--includeClientSideAssets true|false]"
                        return
                    }
                    $includeAssets = if ($null -eq $includeAssetsRaw) { $false } else { Parse-Bool $includeAssetsRaw $false }
                    $info = Get-SpfxProjectInfo
                    if (-not $info) {
                        Write-Warn "package.json not found."
                        return
                    }
                    $solutionPath = Join-Path (Join-Path (Split-Path -Parent $info.Path) "config") "package-solution.json"
                    if (-not (Test-Path $solutionPath)) {
                        Write-Warn "package-solution.json not found."
                        return
                    }
                    try {
                        $sol = Get-Content -Raw -Path $solutionPath | ConvertFrom-Json
                    } catch {
                        Write-Warn "package-solution.json is invalid."
                        return
                    }
                    if (-not $sol.solution) { Write-Warn "solution block missing."; return }
                    $sol.solution.cdnBasePath = $cdn
                    $sol.solution.includeClientSideAssets = $includeAssets
                    $sol | ConvertTo-Json -Depth 10 | Set-Content -Path $solutionPath -Encoding ASCII
                    Write-Info "package-solution.json updated for externalized assets."
                }
                "github" {
                    $path = Get-ArgValue $parsed2.Map "path"
                    if (-not $path) {
                        $path = Join-Path (Join-Path (Join-Path (Get-Location) ".github") "workflows") "spfx-build.yml"
                    }
                    $force = Parse-Bool (Get-ArgValue $parsed2.Map "force") $false
                    $node = Get-ArgValue $parsed2.Map "nodeVersion"
                    if (-not $node) { $node = "18.x" }
                    if ((Test-Path $path) -and -not $force) {
                        Write-Warn "Workflow already exists. Use --force to overwrite."
                        return
                    }
                    $dir = Split-Path -Parent $path
                    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
                    $content = @"
name: SPFx Build
on:
  push:
    branches: [ main ]
  pull_request:
    branches: [ main ]
jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: '$node'
      - run: npm ci
      - run: npx gulp bundle --ship
      - run: npx gulp package-solution --ship
"@
                    $content | Set-Content -Path $path -Encoding ASCII
                    Write-Info ("GitHub workflow written: " + $path)
                }
                "azuredevops" {
                    $path = Get-ArgValue $parsed2.Map "path"
                    if (-not $path) { $path = Join-Path (Get-Location) "azure-pipelines.yml" }
                    $force = Parse-Bool (Get-ArgValue $parsed2.Map "force") $false
                    $node = Get-ArgValue $parsed2.Map "nodeVersion"
                    if (-not $node) { $node = "18.x" }
                    $branch = Get-ArgValue $parsed2.Map "branch"
                    if (-not $branch) { $branch = "main" }
                    if ((Test-Path $path) -and -not $force) {
                        Write-Warn "Pipeline already exists. Use --force to overwrite."
                        return
                    }
                    $content = @"
trigger:
- $branch
pool:
  vmImage: 'ubuntu-latest'
steps:
- task: NodeTool@0
  inputs:
    versionSpec: '$node'
- script: npm ci
  displayName: Install dependencies
- script: npx gulp bundle --ship
  displayName: Bundle
- script: npx gulp package-solution --ship
  displayName: Package
"@
                    $content | Set-Content -Path $path -Encoding ASCII
                    Write-Info ("Azure DevOps pipeline written: " + $path)
                }
                "permissions" {
                    if (-not $rest2 -or $rest2.Count -eq 0) {
                        Write-Warn "Usage: spfx project permissions list|grant [--force]"
                        return
                    }
                    $permAction = $rest2[0].ToLowerInvariant()
                    $info = Get-SpfxProjectInfo
                    if (-not $info) { Write-Warn "package.json not found."; return }
                    $solutionPath = Join-Path (Join-Path (Split-Path -Parent $info.Path) "config") "package-solution.json"
                    if (-not (Test-Path $solutionPath)) { Write-Warn "package-solution.json not found."; return }
                    try {
                        $sol = Get-Content -Raw -Path $solutionPath | ConvertFrom-Json
                    } catch {
                        Write-Warn "package-solution.json is invalid."
                        return
                    }
                    $requests = @()
                    if ($sol.solution -and $sol.solution.webApiPermissionRequests) {
                        $requests = @($sol.solution.webApiPermissionRequests)
                    }
                    if ($requests.Count -eq 0) {
                        Write-Info "No webApiPermissionRequests found."
                        return
                    }
                    if ($permAction -eq "list") {
                        $requests | ForEach-Object {
                            [pscustomobject]@{ Resource = $_.resource; Scope = $_.scope }
                        } | Format-Table -AutoSize
                        return
                    }
                    if ($permAction -ne "grant") {
                        Write-Warn "Usage: spfx project permissions list|grant [--force]"
                        return
                    }
                    if (-not (Require-SpoConnection)) { return }
                    if (-not (Get-Command Get-SPOTenantServicePrincipalPermissionRequests -ErrorAction SilentlyContinue)) {
                        Write-Warn "SPOTenantServicePrincipalPermissionRequests cmdlets not available. Update the SPO module."
                        return
                    }
                    $pending = @(Get-SPOTenantServicePrincipalPermissionRequests)
                    $matches = @()
                    foreach ($req in $requests) {
                        $matches += @($pending | Where-Object { $_.Resource -eq $req.resource -and $_.Scope -eq $req.scope })
                    }
                    if ($matches.Count -eq 0) {
                        Write-Warn "No matching pending permission requests found."
                        return
                    }
                    $force = Parse-Bool (Get-ArgValue $parsed2.Map "force") $false
                    if (-not $force) {
                        $confirm = Read-Host ("Approve " + $matches.Count + " permission request(s)? (y/N)")
                        if ($confirm.ToLowerInvariant() -notin @("y","yes")) { Write-Info "Canceled."; return }
                    }
                    foreach ($m in $matches) {
                        try {
                            Approve-SPOTenantServicePrincipalPermissionRequest -RequestId $m.Id | Out-Null
                            Write-Info ("Approved: " + $m.Resource + " / " + $m.Scope)
                        } catch {
                            Write-Err $_.Exception.Message
                        }
                    }
                }
                default {
                    Write-Warn "Usage: spfx project doctor|rename|upgrade|externalize|github|azuredevops|permissions"
                }
            }
        }
        default {
            Write-Warn "Usage: spfx doctor|package|project <args...>"
        }
    }
}
