# Handler: Context
# Purpose: Manage CLI-like context file for option defaults.
function Handle-ContextCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: context init|remove|option <list|set|remove>"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "init" {
            $map = Load-Context
            if ($map.Count -eq 0) {
                Save-Context @{}
                Write-Info "Context initialized."
            } else {
                Write-Info "Context already initialized."
            }
        }
        "remove" {
            $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
            if (-not $force) {
                $confirm = Read-Host "Remove context file? (y/N)"
                if ($confirm.ToLowerInvariant() -notin @("y","yes")) {
                    Write-Info "Cancelled."
                    return
                }
            }
            if (Remove-Context) {
                Write-Info "Context removed."
            } else {
                Write-Info "Context not found."
            }
        }
        "option" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: context option list|set|remove"
                return
            }
            $optSub = $rest[0].ToLowerInvariant()
            $optRest = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
            $optParsed = Parse-NamedArgs $optRest
            switch ($optSub) {
                "list" {
                    $map = Load-Context
                    if ($map.Count -eq 0) {
                        Write-Info "No context options."
                        return
                    }
                    $map.GetEnumerator() | Sort-Object Name | ForEach-Object {
                        [pscustomobject]@{ Name = $_.Key; Value = $_.Value }
                    } | Format-Table -AutoSize
                }
                "set" {
                    $name = Get-ArgValue $optParsed.Map "name"
                    $valueRaw = Get-ArgValue $optParsed.Map "value"
                    if (-not $name) { $name = ($optParsed.Positionals | Select-Object -First 1) }
                    if ($null -eq $valueRaw -and $optParsed.Positionals.Count -ge 2) {
                        $valueRaw = $optParsed.Positionals[1]
                    }
                    if (-not $name -or $null -eq $valueRaw) {
                        Write-Warn "Usage: context option set --name <name> --value <value>"
                        return
                    }
                    $value = Parse-Value $valueRaw
                    if (Set-ContextOption -Name $name -Value $value) {
                        Write-Info "Context option set."
                    } else {
                        Write-Warn "Unable to set option."
                    }
                }
                "remove" {
                    $name = Get-ArgValue $optParsed.Map "name"
                    $force = Parse-Bool (Get-ArgValue $optParsed.Map "force") $false
                    if (-not $name -and $optParsed.Positionals.Count -gt 0) { $name = $optParsed.Positionals[0] }
                    if (-not $name) {
                        Write-Warn "Usage: context option remove --name <name> [--force]"
                        return
                    }
                    if (-not $force) {
                        $confirm = Read-Host "Remove context option '$name'? (y/N)"
                        if ($confirm.ToLowerInvariant() -notin @("y","yes")) {
                            Write-Info "Cancelled."
                            return
                        }
                    }
                    if (Remove-ContextOption $name) {
                        Write-Info "Context option removed."
                    } else {
                        Write-Warn "Option not found."
                    }
                }
                default {
                    Write-Warn "Usage: context option list|set|remove"
                }
            }
        }
        default {
            Write-Warn "Usage: context init|remove|option <list|set|remove>"
        }
    }
}
