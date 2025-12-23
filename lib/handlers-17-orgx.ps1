# Handler: Orgx
# Purpose: Orgx command handlers.
function Resolve-OrgXUserSegment {
    param(
        [hashtable]$Map,
        [string[]]$Positionals
    )
    $user = Get-ArgValue $Map "user"
    if (-not $user -and $Positionals -and $Positionals.Count -gt 0) {
        $user = $Positionals[0]
    }
    if (-not $user) { $user = "me" }
    return Resolve-UserSegment $user
}

function Handle-OrgXCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: orgx manager|reports|chain|tree ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $seg = Resolve-OrgXUserSegment $parsed.Map $parsed.Positionals
    if (-not $seg) { return }

    switch ($action) {
        "manager" {
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/manager")
            if ($resp) { $resp | ConvertTo-Json -Depth 8 }
        }
        "reports" {
            $resp = Invoke-GraphRequest -Method "GET" -Uri ($seg + "/directReports")
            if ($resp -and $resp.value) {
                $resp.value | Select-Object Id, DisplayName, UserPrincipalName, Mail, "@odata.type" | Format-Table -AutoSize
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 8
            }
        }
        "chain" {
            $depth = 10
            $depthRaw = Get-ArgValue $parsed.Map "depth"
            if ($depthRaw) { try { $depth = [int]$depthRaw } catch {} }
            if ($depth -le 0) { $depth = 10 }
            $cur = $seg
            $list = @()
            for ($i = 0; $i -lt $depth; $i++) {
                $mgr = Invoke-GraphRequest -Method "GET" -Uri ($cur + "/manager")
                if (-not $mgr) { break }
                $list += $mgr
                if ($mgr.id) {
                    $cur = "/users/" + $mgr.id
                } else {
                    break
                }
            }
            if ($list.Count -gt 0) {
                if ($parsed.Map.ContainsKey("json")) {
                    $list | ConvertTo-Json -Depth 8
                } else {
                    $list | Select-Object Id, DisplayName, UserPrincipalName, Mail | Format-Table -AutoSize
                }
            } else {
                Write-Info "No manager chain found."
            }
        }
        "tree" {
            $depth = 2
            $depthRaw = Get-ArgValue $parsed.Map "depth"
            if ($depthRaw) { try { $depth = [int]$depthRaw } catch {} }
            if ($depth -le 0) { $depth = 2 }
            $max = $null
            $maxRaw = Get-ArgValue $parsed.Map "max"
            if ($maxRaw) { try { $max = [int]$maxRaw } catch {} }
            $asJson = $parsed.Map.ContainsKey("json")

            $queue = New-Object System.Collections.Queue
            $queue.Enqueue(@{ Seg = $seg; Level = 0 })
            $results = @()

            while ($queue.Count -gt 0) {
                $node = $queue.Dequeue()
                if ($node.Level -ge $depth) { continue }
                $resp = Invoke-GraphRequest -Method "GET" -Uri ($node.Seg + "/directReports")
                foreach ($r in @($resp.value)) {
                    $results += [pscustomobject]@{
                        Id                = $r.id
                        DisplayName       = $r.displayName
                        UserPrincipalName = $r.userPrincipalName
                        Mail              = $r.mail
                        Level             = ($node.Level + 1)
                        Type              = $r.'@odata.type'
                    }
                    if ($r.id) {
                        $queue.Enqueue(@{ Seg = ("/users/" + $r.id); Level = ($node.Level + 1) })
                    }
                    if ($max -and $results.Count -ge $max) { break }
                }
                if ($max -and $results.Count -ge $max) { break }
            }

            if ($results.Count -eq 0) {
                Write-Info "No direct reports found."
            } elseif ($asJson) {
                $results | ConvertTo-Json -Depth 8
            } else {
                $results | Format-Table -AutoSize
            }
        }
        default {
            Write-Warn "Usage: orgx manager|reports|chain|tree ..."
        }
    }
}
