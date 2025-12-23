# Handler: Compliance
# Purpose: Compliance command handlers.
function Handle-ComplianceCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: compliance connect|disconnect|status|cmd|cmdlets"
        return
    }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "connect" {
            if (-not (Ensure-ModuleLoaded "ExchangeOnlineManagement")) { return }
            $upn = Get-ArgValue $parsed.Map "upn"
            if (-not $upn) { $upn = $global:Config.admin.defaultUpn }
            $delegated = Get-ArgValue $parsed.Map "delegatedOrg"
            $disableWam = Parse-Bool (Get-ArgValue $parsed.Map "disableWam") $false
            $params = @{}
            if ($upn) { $params.UserPrincipalName = $upn }
            if ($delegated) { $params.DelegatedOrganization = $delegated }
            if ($disableWam) { $params.DisableWAM = $true }
            try {
                Connect-IPPSSession @params | Out-Null
                $global:ComplianceConnected = $true
                Write-Info "Connected to Compliance PowerShell."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "disconnect" {
            if (-not (Ensure-ModuleLoaded "ExchangeOnlineManagement")) { return }
            try {
                if (Get-Command Disconnect-IPPSSession -ErrorAction SilentlyContinue) {
                    Disconnect-IPPSSession -Confirm:$false | Out-Null
                }
                $global:ComplianceConnected = $false
                Write-Info "Disconnected from Compliance PowerShell."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "status" {
            $status = if ($global:ComplianceConnected) { "connected" } else { "not connected" }
            Write-Host ("Compliance: " + $status)
        }
        "cmdlets" {
            if (-not (Ensure-ModuleLoaded "ExchangeOnlineManagement")) { return }
            $filter = Get-ArgValue $parsed.Map "filter"
            $cmds = Get-Command -ErrorAction SilentlyContinue | Where-Object {
                $_.Name -match "Compliance|Retention|Sensitivity|Label|Dlp|Case|Search|Hold|Audit|Ediscovery|ReviewSet"
            }
            if ($filter) {
                $cmds = $cmds | Where-Object { $_.Name -like ("*" + $filter + "*") }
            }
            $cmds | Sort-Object Name | Select-Object -ExpandProperty Name | Format-Wide -Column 3
        }
        "cmd" {
            if (-not (Ensure-ModuleLoaded "ExchangeOnlineManagement")) { return }
            if (-not $global:ComplianceConnected) {
                Write-Warn "Not connected. Use: compliance connect"
                return
            }
            $cmdlet = $parsed.Positionals | Select-Object -First 1
            if (-not $cmdlet) {
                Write-Warn "Usage: compliance cmd <cmdlet> [--params key=value[,key=value]] [--json <payload>] [--bodyFile <file>] [--set key=value]"
                return
            }
            $paramObj = Resolve-CmdletParams $parsed
            try {
                if ($paramObj.Keys.Count -gt 0) {
                    & $cmdlet @paramObj
                } else {
                    & $cmdlet
                }
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        default {
            Write-Warn "Usage: compliance connect|disconnect|status|cmd|cmdlets"
        }
    }
}
