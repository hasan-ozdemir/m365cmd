# Handler: Auth
# Purpose: Auth command handlers.
function Resolve-AuthUserSegment {
    param([hashtable]$Map)
    $seg = Resolve-UserSegment (Get-ArgValue $Map "user")
    if (-not $seg) { return $null }
    return $seg
}

function Handle-AuthMethodCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: authmethod list|get|delete|phone|email|tap ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $seg = Resolve-AuthUserSegment $parsed.Map
    if (-not $seg) { return }

    if ($sub -in @("list","get","delete")) {
        $base = $seg + "/authentication/methods"
        switch ($sub) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed.Map @("id")
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
                if ($resp -and $resp.value) {
                    $resp.value | ConvertTo-Json -Depth 8
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 8
                }
            }
            "get" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: authmethod get <methodId> [--user <upn|id>]"
                    return
                }
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
            }
            "delete" {
                $id = $parsed.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: authmethod delete <methodId> [--user <upn|id>] [--force]"
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
                $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($base + "/" + $id)
                if ($resp -ne $null) { Write-Info "Authentication method deleted." }
            }
        }
        return
    }

    if ($sub -eq "phone") {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: authmethod phone list|get|add|update|delete"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed2 = Parse-NamedArgs $rest2
        $base = $seg + "/authentication/phoneMethods"
        switch ($action) {
            "list" {
                $qh = Build-QueryAndHeaders $parsed2.Map @("id","phoneNumber","phoneType","smsSignInState")
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers
                if ($resp -and $resp.value) {
                    $resp.value | Select-Object Id, PhoneType, PhoneNumber, SmsSignInState | Format-Table -AutoSize
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 8
                }
            }
            "get" {
                $id = $parsed2.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: authmethod phone get <methodId> [--user <upn|id>]"
                    return
                }
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($base + "/" + $id)
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
            }
            "add" {
                $number = Get-ArgValue $parsed2.Map "number"
                if (-not $number) { $number = Get-ArgValue $parsed2.Map "phoneNumber" }
                if (-not $number) { $number = Get-ArgValue $parsed2.Map "phone" }
                $type = Get-ArgValue $parsed2.Map "type"
                if (-not $type) { $type = "mobile" }
                $jsonRaw = Get-ArgValue $parsed2.Map "json"
                $body = if ($jsonRaw) { Parse-Value $jsonRaw } else { @{ phoneNumber = $number; phoneType = $type } }
                if (-not $body -or -not $body.phoneNumber) {
                    Write-Warn "Usage: authmethod phone add --number <phone> [--type mobile|alternateMobile|office] OR --json <payload>"
                    return
                }
                $resp = Invoke-GraphRequest -Method "POST" -Uri $base -Body $body
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
            }
            "update" {
                $id = $parsed2.Positionals | Select-Object -First 1
                $body = Read-JsonPayload (Get-ArgValue $parsed2.Map "json") (Get-ArgValue $parsed2.Map "bodyFile") (Get-ArgValue $parsed2.Map "set")
                if (-not $body) {
                    $number = Get-ArgValue $parsed2.Map "number"
                    if (-not $number) { $number = Get-ArgValue $parsed2.Map "phoneNumber" }
                    $type = Get-ArgValue $parsed2.Map "type"
                    $tmp = @{}
                    if ($number) { $tmp.phoneNumber = $number }
                    if ($type) { $tmp.phoneType = $type }
                    if ($tmp.Keys.Count -gt 0) { $body = $tmp }
                }
                if (-not $id -or -not $body) {
                    Write-Warn "Usage: authmethod phone update <methodId> --number <phone> [--type mobile|alternateMobile|office] OR --json <payload>"
                    return
                }
                $resp = Invoke-GraphRequest -Method "PATCH" -Uri ($base + "/" + $id) -Body $body
                if ($resp -ne $null) { Write-Info "Phone method updated." }
            }
            "delete" {
                $id = $parsed2.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: authmethod phone delete <methodId> [--force]"
                    return
                }
                $force = Parse-Bool (Get-ArgValue $parsed2.Map "force") $false
                if (-not $force) {
                    $confirm = Read-Host "Type DELETE to confirm"
                    if ($confirm -ne "DELETE") {
                        Write-Info "Canceled."
                        return
                    }
                }
                $resp = Invoke-GraphRequest -Method "DELETE" -Uri ($base + "/" + $id)
                if ($resp -ne $null) { Write-Info "Phone method deleted." }
            }
            default {
                Write-Warn "Usage: authmethod phone list|get|add|update|delete"
            }
        }
        return
    }

    if ($sub -eq "email") {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: authmethod email list|get|create|update|delete"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed2 = Parse-NamedArgs $rest2
        $base = $seg + "/authentication/emailMethods"
        $useBeta = $parsed2.Map.ContainsKey("beta")
        $useV1 = $parsed2.Map.ContainsKey("v1")
        $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
        if ($useBeta -or $useV1) { $allowFallback = $false }
        $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }

        switch ($action) {
            "list" {
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri $base -Api $api -AllowFallback:$allowFallback
                if ($resp -and $resp.value) {
                    $resp.value | Select-Object Id, EmailAddress | Format-Table -AutoSize
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 8
                }
            }
            "get" {
                $id = $parsed2.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: authmethod email get <methodId>"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
            }
            "create" {
                $email = Get-ArgValue $parsed2.Map "email"
                if (-not $email) { $email = Get-ArgValue $parsed2.Map "emailAddress" }
                $jsonRaw = Get-ArgValue $parsed2.Map "json"
                $body = if ($jsonRaw) { Parse-Value $jsonRaw } else { @{ emailAddress = $email } }
                if (-not $body -or -not $body.emailAddress) {
                    Write-Warn "Usage: authmethod email create --email <address> OR --json <payload>"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
            }
            "update" {
                $id = $parsed2.Positionals | Select-Object -First 1
                $body = Read-JsonPayload (Get-ArgValue $parsed2.Map "json") (Get-ArgValue $parsed2.Map "bodyFile") (Get-ArgValue $parsed2.Map "set")
                if (-not $body) {
                    $email = Get-ArgValue $parsed2.Map "email"
                    if (-not $email) { $email = Get-ArgValue $parsed2.Map "emailAddress" }
                    if ($email) { $body = @{ emailAddress = $email } }
                }
                if (-not $id -or -not $body) {
                    Write-Warn "Usage: authmethod email update <methodId> --email <address> OR --json <payload> [--beta]"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "PATCH" -Uri ($base + "/" + $id) -Body $body -Api $api -AllowFallback:$allowFallback
                if ($resp -ne $null) { Write-Info "Email method updated." }
            }
            "delete" {
                $id = $parsed2.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: authmethod email delete <methodId> [--force]"
                    return
                }
                $force = Parse-Bool (Get-ArgValue $parsed2.Map "force") $false
                if (-not $force) {
                    $confirm = Read-Host "Type DELETE to confirm"
                    if ($confirm -ne "DELETE") {
                        Write-Info "Canceled."
                        return
                    }
                }
                $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback -AllowNullResponse
                if ($resp -ne $null) { Write-Info "Email method deleted." } else { Write-Info "Delete requested." }
            }
            default {
                Write-Warn "Usage: authmethod email list|get|create|update|delete"
            }
        }
        return
    }

    if ($sub -eq "tap") {
        if (-not $rest -or $rest.Count -eq 0) {
            Write-Warn "Usage: authmethod tap list|get|create|delete"
            return
        }
        $action = $rest[0].ToLowerInvariant()
        $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
        $parsed2 = Parse-NamedArgs $rest2
        $base = $seg + "/authentication/temporaryAccessPassMethods"
        $useBeta = $parsed2.Map.ContainsKey("beta")
        $useV1 = $parsed2.Map.ContainsKey("v1")
        $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
        if ($useBeta -or $useV1) { $allowFallback = $false }
        $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }

        switch ($action) {
            "list" {
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri $base -Api $api -AllowFallback:$allowFallback
                if ($resp -and $resp.value) {
                    $resp.value | Select-Object Id, CreatedDateTime, LifetimeInMinutes, IsUsable, IsUsableOnce | Format-Table -AutoSize
                } elseif ($resp) {
                    $resp | ConvertTo-Json -Depth 8
                }
            }
            "get" {
                $id = $parsed2.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: authmethod tap get <methodId>"
                    return
                }
                $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
            }
            "create" {
                $jsonRaw = Get-ArgValue $parsed2.Map "json"
                $bodyFile = Get-ArgValue $parsed2.Map "bodyFile"
                $body = Read-JsonPayload $jsonRaw $bodyFile $null
                if (-not $body) {
                    $start = Get-ArgValue $parsed2.Map "start"
                    $lifetime = Get-ArgValue $parsed2.Map "lifetime"
                    $once = Get-ArgValue $parsed2.Map "once"
                    $tmp = @{}
                    if ($start) { $tmp.startDateTime = $start }
                    if ($lifetime) { $tmp.lifetimeInMinutes = [int]$lifetime }
                    if ($once -ne $null) { $tmp.isUsableOnce = (Parse-Bool $once $false) }
                    if ($tmp.Keys.Count -gt 0) { $body = $tmp } else { $body = @{} }
                }
                $resp = Invoke-GraphRequestAuto -Method "POST" -Uri $base -Body $body -Api $api -AllowFallback:$allowFallback
                if ($resp) { $resp | ConvertTo-Json -Depth 8 }
            }
            "delete" {
                $id = $parsed2.Positionals | Select-Object -First 1
                if (-not $id) {
                    Write-Warn "Usage: authmethod tap delete <methodId> [--force]"
                    return
                }
                $force = Parse-Bool (Get-ArgValue $parsed2.Map "force") $false
                if (-not $force) {
                    $confirm = Read-Host "Type DELETE to confirm"
                    if ($confirm -ne "DELETE") {
                        Write-Info "Canceled."
                        return
                    }
                }
                $resp = Invoke-GraphRequestAuto -Method "DELETE" -Uri ($base + "/" + $id) -Api $api -AllowFallback:$allowFallback -AllowNullResponse
                if ($resp -ne $null) { Write-Info "Temporary Access Pass deleted." } else { Write-Info "Delete requested." }
            }
            default {
                Write-Warn "Usage: authmethod tap list|get|create|delete"
            }
        }
        return
    }

    Write-Warn "Usage: authmethod list|get|delete|phone|email|tap ..."
}
