# Handler: Device
# Purpose: Device command handlers.
function Handle-DeviceCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: device list|get|update|delete"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($sub) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @("id", "displayName", "deviceId", "operatingSystem", "accountEnabled")
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ("/devices" + $qh.Query) -Headers $qh.Headers -AllowFallback:$true
            if ($resp -and $resp.value) {
                Write-GraphTable $resp.value @("Id", "DisplayName", "DeviceId", "OperatingSystem", "AccountEnabled")
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: device get <id>"
                return
            }
            $qh = Build-QueryAndHeaders $parsed.Map @()
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ("/devices/" + $id + $qh.Query) -Headers $qh.Headers -AllowFallback:$true
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "update" {
            $id = $parsed.Positionals | Select-Object -First 1
            $setRaw = Get-ArgValue $parsed.Map "set"
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            if (-not $id -or (-not $setRaw -and -not $jsonRaw)) {
                Write-Warn "Usage: device update <id> --set key=value[,key=value] OR --json <payload>"
                return
            }
            $body = if ($jsonRaw) { Parse-Value $jsonRaw } else { Parse-Value $setRaw }
            if ($body -is [string]) { $body = Parse-KvPairs $setRaw }
            if (-not $body -or $body.Keys.Count -eq 0) {
                Write-Warn "No properties to update."
                return
            }
            $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ("/devices/" + $id) -Body $body -AllowFallback:$true
            if ($resp -ne $null) { Write-Info "Device updated." }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: device delete <id> [--force]"
                return
            }
            $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
            if (-not $force) {
                $confirm = Read-Host "Type DELETE to confirm"
                if ($confirm -ne "DELETE") {
                    Write-Info "Canceled."
                    return
                }
            }
            $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri ("/devices/" + $id) -AllowFallback:$true -AllowNullResponse
            if ($resp -ne $null) { Write-Info "Device deleted." }
        }
        default {
            Write-Warn "Usage: device list|get|update|delete"
        }
    }
}

