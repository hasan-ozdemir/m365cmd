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
                    $solutionPath = Join-Path (Split-Path -Parent $info.Path) "config\\package-solution.json"
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
                    Write-Warn "spfx project upgrade is not implemented yet."
                }
                "externalize" {
                    Write-Warn "spfx project externalize is not implemented yet."
                }
                "github" {
                    Write-Warn "spfx project github workflow add is not implemented yet."
                }
                "azuredevops" {
                    Write-Warn "spfx project azuredevops pipeline add is not implemented yet."
                }
                "permissions" {
                    Write-Warn "spfx project permissions grant is not implemented yet."
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
