# Handler: Connection
# Purpose: Manage multiple saved connections.
function Handle-ConnectionCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: connection list|use|set|remove"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "list" {
            $list = Load-Connections
            if (-not $list -or $list.Count -eq 0) {
                Write-Info "No saved connections. Use /login first."
                return
            }
            $out = $list | ForEach-Object {
                [pscustomobject]@{
                    name        = $_.name
                    connectedAs = $_.connectedAs
                    authType    = $_.authType
                    active      = $_.active
                    tenantId    = $_.tenantId
                    updatedAt   = $_.updatedAt
                }
            }
            $out | Format-Table -AutoSize
        }
        "use" {
            $name = Get-ArgValue $parsed.Map "name"
            if (-not $name -and $rest.Count -gt 0) { $name = $rest[0] }
            if (-not $name) {
                Write-Warn "Usage: connection use --name <name>"
                return
            }
            $conn = Get-ConnectionByName $name
            if (-not $conn) {
                Write-Warn "Connection not found."
                return
            }
            $scopes = $global:Config.auth.scopes
            $contextScope = $global:Config.auth.contextScope
            try {
                if ($conn.tenantId) {
                    Connect-MgGraph -Scopes $scopes -ContextScope $contextScope -TenantId $conn.tenantId | Out-Null
                } else {
                    Connect-MgGraph -Scopes $scopes -ContextScope $contextScope | Out-Null
                }
                Set-ActiveConnectionByName $conn.name | Out-Null
                Write-Info ("Active connection: " + $conn.name)
            } catch {
                Write-Warn "Unable to switch connection. Try /login."
            }
        }
        "set" {
            $name = Get-ArgValue $parsed.Map "name"
            $newName = Get-ArgValue $parsed.Map "newName"
            if (-not $name -or -not $newName) {
                Write-Warn "Usage: connection set --name <name> --newName <newName>"
                return
            }
            if ($name -eq $newName) {
                Write-Warn "Choose a name different from the current one."
                return
            }
            if (Get-ConnectionByName $newName) {
                Write-Warn "Target name already exists."
                return
            }
            if (Rename-Connection -Name $name -NewName $newName) {
                Write-Info "Connection renamed."
            } else {
                Write-Warn "Connection not found."
            }
        }
        "remove" {
            $name = Get-ArgValue $parsed.Map "name"
            $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
            if (-not $name -and $rest.Count -gt 0) { $name = $rest[0] }
            if (-not $name) {
                Write-Warn "Usage: connection remove --name <name> [--force]"
                return
            }
            if (-not $force) {
                $confirm = Read-Host "Remove connection '$name'? (y/N)"
                if ($confirm.ToLowerInvariant() -notin @("y","yes")) {
                    Write-Info "Cancelled."
                    return
                }
            }
            if (Remove-Connection $name) {
                Write-Info "Connection removed."
            } else {
                Write-Warn "Connection not found."
            }
        }
        default {
            Write-Warn "Usage: connection list|use|set|remove"
        }
    }
}
