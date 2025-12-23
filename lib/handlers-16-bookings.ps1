# Handler: Bookings
# Purpose: Bookings command handlers.
function Resolve-BookingsBusinessId {
    param([hashtable]$Map)
    $id = Get-ArgValue $Map "business"
    if (-not $id) { $id = Get-ArgValue $Map "biz" }
    if (-not $id) { $id = Get-ArgValue $Map "businessId" }
    return $id
}

function Invoke-BookingsList {
    param(
        [string]$Base,
        [hashtable]$Map,
        [string[]]$SelectDefaults
    )
    $qh = Build-QueryAndHeaders $Map $SelectDefaults
    $resp = Invoke-GraphRequest -Method "GET" -Uri ($Base + $qh.Query) -Headers $qh.Headers
    if ($resp -and $resp.value) {
        if ($SelectDefaults -and $SelectDefaults.Count -gt 0) {
            Write-GraphTable $resp.value $SelectDefaults
        } else {
            $resp.value | ConvertTo-Json -Depth 8
        }
    } elseif ($resp) {
        $resp | ConvertTo-Json -Depth 8
    }
}

function Invoke-BookingsGet {
    param([string]$Base, [string]$Id)
    $resp = Invoke-GraphRequest -Method "GET" -Uri ($Base + "/" + $Id)
    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
}

function Invoke-BookingsCreate {
    param([string]$Base, [object]$Body)
    $resp = Invoke-GraphRequest -Method "POST" -Uri $Base -Body $Body
    if ($resp) { $resp | ConvertTo-Json -Depth 8 }
}

function Invoke-BookingsUpdate {
    param([string]$Base, [string]$Id, [object]$Body)
    $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($Base + "/" + $Id) -Body $Body
    if ($resp -ne $null) { Write-Info "Updated." }
}

function Invoke-BookingsDelete {
    param([string]$Base, [string]$Id, [switch]$Force)
    if (-not $Force) {
        $confirm = Read-Host "Type DELETE to confirm"
        if ($confirm -ne "DELETE") {
            Write-Info "Canceled."
            return
        }
    }
    $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($Base + "/" + $Id)
    if ($resp -ne $null) { Write-Info "Deleted." }
}

function Handle-BookingsCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: bookings business|service|staff|appointment|customer ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    if (-not $rest -or $rest.Count -eq 0) {
        Write-Warn "Usage: bookings <area> list|get|create|update|delete ..."
        return
    }
    $action = $rest[0].ToLowerInvariant()
    $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest2

    switch ($sub) {
        "business" {
            $base = "/solutions/bookingBusinesses"
            switch ($action) {
                "list" {
                    Invoke-BookingsList $base $parsed.Map @("Id","DisplayName","Email","IsPublished")
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) { Write-Warn "Usage: bookings business get <businessId>"; return }
                    Invoke-BookingsGet $base $id
                }
                "create" {
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $body) { Write-Warn "Usage: bookings business create --json <payload>"; return }
                    Invoke-BookingsCreate $base $body
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) { Write-Warn "Usage: bookings business update <businessId> --json <payload> OR --set key=value"; return }
                    Invoke-BookingsUpdate $base $id $body
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) { Write-Warn "Usage: bookings business delete <businessId> [--force]"; return }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    Invoke-BookingsDelete $base $id -Force:$force
                }
                default { Write-Warn "Usage: bookings business list|get|create|update|delete" }
            }
        }
        "service" {
            $biz = Resolve-BookingsBusinessId $parsed.Map
            if (-not $biz) { Write-Warn "Usage: bookings service <action> --business <id>"; return }
            $base = "/solutions/bookingBusinesses/" + $biz + "/services"
            switch ($action) {
                "list" { Invoke-BookingsList $base $parsed.Map @("Id","DisplayName","DefaultDuration","DefaultPrice","IsLocationOnline") }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) { Write-Warn "Usage: bookings service get <serviceId> --business <id>"; return }
                    Invoke-BookingsGet $base $id
                }
                "create" {
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $body) { Write-Warn "Usage: bookings service create --business <id> --json <payload>"; return }
                    Invoke-BookingsCreate $base $body
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) { Write-Warn "Usage: bookings service update <serviceId> --business <id> --json <payload> OR --set key=value"; return }
                    Invoke-BookingsUpdate $base $id $body
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) { Write-Warn "Usage: bookings service delete <serviceId> --business <id> [--force]"; return }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    Invoke-BookingsDelete $base $id -Force:$force
                }
                default { Write-Warn "Usage: bookings service list|get|create|update|delete --business <id>" }
            }
        }
        "staff" {
            $biz = Resolve-BookingsBusinessId $parsed.Map
            if (-not $biz) { Write-Warn "Usage: bookings staff <action> --business <id>"; return }
            $base = "/solutions/bookingBusinesses/" + $biz + "/staffMembers"
            switch ($action) {
                "list" { Invoke-BookingsList $base $parsed.Map @("Id","DisplayName","EmailAddress","IsActive","Role") }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) { Write-Warn "Usage: bookings staff get <staffId> --business <id>"; return }
                    Invoke-BookingsGet $base $id
                }
                "create" {
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $body) { Write-Warn "Usage: bookings staff create --business <id> --json <payload>"; return }
                    Invoke-BookingsCreate $base $body
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) { Write-Warn "Usage: bookings staff update <staffId> --business <id> --json <payload> OR --set key=value"; return }
                    Invoke-BookingsUpdate $base $id $body
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) { Write-Warn "Usage: bookings staff delete <staffId> --business <id> [--force]"; return }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    Invoke-BookingsDelete $base $id -Force:$force
                }
                default { Write-Warn "Usage: bookings staff list|get|create|update|delete --business <id>" }
            }
        }
        "appointment" {
            $biz = Resolve-BookingsBusinessId $parsed.Map
            if (-not $biz) { Write-Warn "Usage: bookings appointment <action> --business <id>"; return }
            $base = "/solutions/bookingBusinesses/" + $biz + "/appointments"
            switch ($action) {
                "list" { Invoke-BookingsList $base $parsed.Map @("Id","StartDateTime","EndDateTime","CustomerName","ServiceName") }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) { Write-Warn "Usage: bookings appointment get <appointmentId> --business <id>"; return }
                    Invoke-BookingsGet $base $id
                }
                "create" {
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $body) { Write-Warn "Usage: bookings appointment create --business <id> --json <payload>"; return }
                    Invoke-BookingsCreate $base $body
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) { Write-Warn "Usage: bookings appointment update <appointmentId> --business <id> --json <payload> OR --set key=value"; return }
                    Invoke-BookingsUpdate $base $id $body
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) { Write-Warn "Usage: bookings appointment delete <appointmentId> --business <id> [--force]"; return }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    Invoke-BookingsDelete $base $id -Force:$force
                }
                default { Write-Warn "Usage: bookings appointment list|get|create|update|delete --business <id>" }
            }
        }
        "customer" {
            $biz = Resolve-BookingsBusinessId $parsed.Map
            if (-not $biz) { Write-Warn "Usage: bookings customer <action> --business <id>"; return }
            $base = "/solutions/bookingBusinesses/" + $biz + "/customers"
            switch ($action) {
                "list" { Invoke-BookingsList $base $parsed.Map @("Id","DisplayName","EmailAddress","Phone") }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) { Write-Warn "Usage: bookings customer get <customerId> --business <id>"; return }
                    Invoke-BookingsGet $base $id
                }
                "create" {
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $body) { Write-Warn "Usage: bookings customer create --business <id> --json <payload>"; return }
                    Invoke-BookingsCreate $base $body
                }
                "update" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed.Map "json") (Get-ArgValue $parsed.Map "bodyFile") (Get-ArgValue $parsed.Map "set")
                    if (-not $id -or -not $body) { Write-Warn "Usage: bookings customer update <customerId> --business <id> --json <payload> OR --set key=value"; return }
                    Invoke-BookingsUpdate $base $id $body
                }
                "delete" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) { Write-Warn "Usage: bookings customer delete <customerId> --business <id> [--force]"; return }
                    $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
                    Invoke-BookingsDelete $base $id -Force:$force
                }
                default { Write-Warn "Usage: bookings customer list|get|create|update|delete --business <id>" }
            }
        }
        default {
            Write-Warn "Usage: bookings business|service|staff|appointment|customer ..."
        }
    }
}
