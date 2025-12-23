# Handler: Risk
# Purpose: Risk command handlers.
function Handle-RiskCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: risk detection|user ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $useBeta = $parsed.Map.ContainsKey("beta")
    $useV1 = $parsed.Map.ContainsKey("v1")
    $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
    if ($parsed.Map.ContainsKey("auto")) { $allowFallback = $true }
    if ($useBeta -or $useV1) { $allowFallback = $false }
    $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }

    if ($sub -in @("detection","detections","riskdetection","riskdetections")) {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: risk detection list|get"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed2 = Parse-NamedArgs $rest2
        $base = "/identityProtection/riskDetections"
        switch ($action) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed2.Map @()
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $api -AllowFallback:$allowFallback
                if ($resp -and $resp.value) {
                    $resp.value | ConvertTo-Json -Depth 8
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 8
                }
            }
            "get" {
                $id = $parsed2.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: risk detection get <id>"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
            }
            default {
                Write-Warn "Usage: risk detection list|get"
            }
        }
        return
    }

    if ($sub -in @("user","users","riskyuser","riskyusers")) {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: risk user list|get|history|confirm|dismiss"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed2 = Parse-NamedArgs $rest2
        $base = "/identityProtection/riskyUsers"
        switch ($action) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed2.Map @()
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $api -AllowFallback:$allowFallback
                if ($resp -and $resp.value) {
                    $resp.value | ConvertTo-Json -Depth 8
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 8
                }
            }
            "get" {
                $id = $parsed2.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: risk user get <riskyUserId>"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
            }
            "history" {
                $id = $parsed2.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: risk user history <riskyUserId>"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id + "/history") -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
            }
            "confirm" {
                $idsRaw = Get-ArgValue $parsed2.Map "ids"
                $userRaw = Get-ArgValue $parsed2.Map "user"
                $list = @()
                if ($idsRaw) { $list = Parse-CommaList $idsRaw }
                if ($userRaw) {
                    foreach ($u in (Parse-CommaList $userRaw)) {
                        $obj = Resolve-UserObject $u
                        if ($obj -and $obj.Id) { $list += $obj.Id }
                    }
                }
                if ($list.Count -eq 0) {
                    Write-Warn "Usage: risk user confirm --ids <id1,id2> OR --user <upn,id>"
                    return
                }
                $body = @{ userIds = @($list | Select-Object -Unique) }
                $resp = Invoke-GraphRequestAuto -Method "POST" -Uri ($base + "/confirmCompromised") -Body $body -Api $api -AllowFallback:$allowFallback
                if ($resp -ne $null) { Write-Info "Confirm compromised requested." }
            }
            "dismiss" {
                $idsRaw = Get-ArgValue $parsed2.Map "ids"
                $userRaw = Get-ArgValue $parsed2.Map "user"
                $list = @()
                if ($idsRaw) { $list = Parse-CommaList $idsRaw }
                if ($userRaw) {
                    foreach ($u in (Parse-CommaList $userRaw)) {
                        $obj = Resolve-UserObject $u
                        if ($obj -and $obj.Id) { $list += $obj.Id }
                    }
                }
                if ($list.Count -eq 0) {
                    Write-Warn "Usage: risk user dismiss --ids <id1,id2> OR --user <upn,id>"
                    return
                }
                $body = @{ userIds = @($list | Select-Object -Unique) }
                $resp = Invoke-GraphRequestAuto -Method "POST" -Uri ($base + "/dismiss") -Body $body -Api $api -AllowFallback:$allowFallback
                if ($resp -ne $null) { Write-Info "Dismiss requested." }
            }
            default {
                Write-Warn "Usage: risk user list|get|history|confirm|dismiss"
            }
        }
        return
    }

    Write-Warn "Usage: risk detection|user ..."
}
