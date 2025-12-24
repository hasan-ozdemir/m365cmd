# Handler: Defender
# Purpose: Defender command handlers.
function Handle-DefenderCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: defender incident|alert|hunt|machine ..."
        return
    }

    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    $baseUrl = $global:Config.defender.baseUrl
    $centerBaseUrl = $global:Config.defender.centerBaseUrl
    if (-not $centerBaseUrl) { $centerBaseUrl = $baseUrl }
    $scope = "https://api.security.microsoft.com/.default"
    $base = "/api"

    $invokeDefender = {
        param(
            [string]$Method,
            [string]$Url,
            [object]$Body,
            [string]$PreferBase = "base",
            [bool]$AllowNull = $false
        )
        $primary = if ($PreferBase -eq "center") { $centerBaseUrl } else { $baseUrl }
        $secondary = if ($PreferBase -eq "center") { $baseUrl } else { $centerBaseUrl }
        $resp = Invoke-ExternalApiRequest -Method $Method -Url $Url -Body $Body -Scope $scope -BaseUrl $primary -AllowNullResponse:$AllowNull
        if (-not $resp -and $secondary -and $secondary -ne $primary) {
            $resp = Invoke-ExternalApiRequest -Method $Method -Url $Url -Body $Body -Scope $scope -BaseUrl $secondary
        }
        return $resp
    }

    switch ($sub) {
        "incident" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: defender incident list|get|update"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $parsed2 = Parse-NamedArgs ($rest | Select-Object -Skip 1)
            $qh = Build-QueryAndHeaders $parsed2.Map @()
            switch ($action) {
                "list" {
                    $resp = & $invokeDefender "GET" ($base + "/incidents" + $qh.Query) $null "base"
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "get" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: defender incident get <id>"
                        return
                    }
                    $resp = & $invokeDefender "GET" ($base + "/incidents/" + $id) $null "base"
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "update" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed2.Map "json") (Get-ArgValue $parsed2.Map "bodyFile") (Get-ArgValue $parsed2.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: defender incident update <id> --json <payload> OR --set key=value"
                        return
                    }
                    $resp = & $invokeDefender "PATCH" ($base + "/incidents/" + $id) $body "base"
                    if ($resp -ne $null) { Write-Info "Incident updated." }
                }
                default {
                    Write-Warn "Usage: defender incident list|get|update"
                }
            }
        }
        "alert" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: defender alert list|get|update"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $parsed2 = Parse-NamedArgs ($rest | Select-Object -Skip 1)
            $qh = Build-QueryAndHeaders $parsed2.Map @()
            switch ($action) {
                "list" {
                    $resp = & $invokeDefender "GET" ($base + "/alerts" + $qh.Query) $null "base"
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "get" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: defender alert get <id>"
                        return
                    }
                    $resp = & $invokeDefender "GET" ($base + "/alerts/" + $id) $null "base"
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "update" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    $body = Read-JsonPayload (Get-ArgValue $parsed2.Map "json") (Get-ArgValue $parsed2.Map "bodyFile") (Get-ArgValue $parsed2.Map "set")
                    if (-not $id -or -not $body) {
                        Write-Warn "Usage: defender alert update <id> --json <payload> OR --set key=value"
                        return
                    }
                    $resp = & $invokeDefender "PATCH" ($base + "/alerts/" + $id) $body "base"
                    if ($resp -ne $null) { Write-Info "Alert updated." }
                }
                default {
                    Write-Warn "Usage: defender alert list|get|update"
                }
            }
        }
        "hunt" {
            $query = Get-ArgValue $parsed.Map "query"
            $queryFile = Get-ArgValue $parsed.Map "queryFile"
            if (-not $query -and -not $queryFile) {
                Write-Warn "Usage: defender hunt --query <kql> [--queryFile <path>]"
                return
            }
            if ($queryFile) {
                if (-not (Test-Path $queryFile)) {
                    Write-Warn "Query file not found."
                    return
                }
                $query = Get-Content -Raw -Path $queryFile
            }
            $body = @{ Query = $query }
            $resp = & $invokeDefender "POST" ($base + "/advancedhunting/run") $body "base"
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        { $_ -in @("machineaction", "machineactions") } {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: defender machineaction list|get [--filter <odata>] [--top <n>]"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $parsed2 = Parse-NamedArgs ($rest | Select-Object -Skip 1)
            $qh = Build-QueryAndHeaders $parsed2.Map @()
            switch ($action) {
                "list" {
                    $resp = & $invokeDefender "GET" ($base + "/machineactions" + $qh.Query) $null "base"
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "get" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: defender machineaction get <id>"
                        return
                    }
                    $resp = & $invokeDefender "GET" ($base + "/machineactions/" + $id) $null "base"
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                default {
                    Write-Warn "Usage: defender machineaction list|get [--filter <odata>] [--top <n>]"
                }
            }
        }
        "machine" {
            if (-not $rest -or $rest.Count -eq 0) {
                Write-Warn "Usage: defender machine list|get|findbytag|isolate|unisolate|collect|runscan|stopfile"
                return
            }
            $action = $rest[0].ToLowerInvariant()
            $parsed2 = Parse-NamedArgs ($rest | Select-Object -Skip 1)
            $qh = Build-QueryAndHeaders $parsed2.Map @()
            switch ($action) {
                "list" {
                    $resp = & $invokeDefender "GET" ($base + "/machines" + $qh.Query) $null "base" $true
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "get" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: defender machine get <id>"
                        return
                    }
                    $resp = & $invokeDefender "GET" ($base + "/machines/" + $id) $null "base" $true
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "findbytag" {
                    $tag = Get-ArgValue $parsed2.Map "tag"
                    if (-not $tag) {
                        Write-Warn "Usage: defender machine findbytag --tag <tag>"
                        return
                    }
                    $qs = "tag=" + (Encode-QueryValue $tag)
                    $resp = & $invokeDefender "GET" ($base + "/machines/findbytag?" + $qs) $null "center" $true
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "isolate" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    $comment = Get-ArgValue $parsed2.Map "comment"
                    $itype = Get-ArgValue $parsed2.Map "type"
                    if (-not $id -or -not $comment) {
                        Write-Warn "Usage: defender machine isolate <id> --comment <text> [--type Full|Selective]"
                        return
                    }
                    $body = @{ Comment = $comment }
                    if ($itype) { $body.IsolationType = $itype }
                    $resp = & $invokeDefender "POST" ($base + "/machines/" + $id + "/isolate") $body "center" $true
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "unisolate" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    $comment = Get-ArgValue $parsed2.Map "comment"
                    if (-not $id -or -not $comment) {
                        Write-Warn "Usage: defender machine unisolate <id> --comment <text>"
                        return
                    }
                    $body = @{ Comment = $comment }
                    $resp = & $invokeDefender "POST" ($base + "/machines/" + $id + "/unisolate") $body "base" $true
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "collect" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    $comment = Get-ArgValue $parsed2.Map "comment"
                    if (-not $id -or -not $comment) {
                        Write-Warn "Usage: defender machine collect <id> --comment <text>"
                        return
                    }
                    $body = @{ Comment = $comment }
                    $resp = & $invokeDefender "POST" ($base + "/machines/" + $id + "/collectInvestigationPackage") $body "center" $true
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "runscan" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    $comment = Get-ArgValue $parsed2.Map "comment"
                    $stype = Get-ArgValue $parsed2.Map "type"
                    if (-not $id -or -not $comment) {
                        Write-Warn "Usage: defender machine runscan <id> --comment <text> [--type Full|Quick]"
                        return
                    }
                    $body = @{ Comment = $comment }
                    if ($stype) { $body.ScanType = $stype }
                    $resp = & $invokeDefender "POST" ($base + "/machines/" + $id + "/runAntiVirusScan") $body "base" $true
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                "stopfile" {
                    $id = $parsed2.Positionals | Select-Object -First 1
                    $comment = Get-ArgValue $parsed2.Map "comment"
                    $sha1 = Get-ArgValue $parsed2.Map "sha1"
                    if (-not $id -or -not $comment -or -not $sha1) {
                        Write-Warn "Usage: defender machine stopfile <id> --comment <text> --sha1 <hash>"
                        return
                    }
                    $body = @{ Comment = $comment; Sha1 = $sha1 }
                    $resp = & $invokeDefender "POST" ($base + "/machines/" + $id + "/StopAndQuarantineFile") $body "center" $true
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                default {
                    Write-Warn "Usage: defender machine list|get|findbytag|isolate|unisolate|collect|runscan|stopfile"
                }
            }
        }
        default {
            Write-Warn "Usage: defender incident|alert|hunt|machine|machineaction"
        }
    }
}

