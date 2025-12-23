# Handler: Label
# Purpose: Label command handlers.
function Handle-LabelCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: label list|get [--user <upn|id>|--me|--org] [--beta]"
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest
    $useBeta = $true
    if ($parsed.Map.ContainsKey("v1")) { $useBeta = $false }
    $api = if ($useBeta) { "beta" } else { "v1" }

    $seg = $null
    if ($parsed.Map.ContainsKey("me")) {
        $seg = "/me"
    } elseif ($parsed.Map.ContainsKey("user")) {
        $seg = Resolve-UserSegment (Get-ArgValue $parsed.Map "user")
        if (-not $seg) { return }
    } elseif ($parsed.Map.ContainsKey("org")) {
        $seg = ""
    } else {
        $seg = "/me"
    }
    $path = if ($seg -eq "") { "/security/informationProtection/sensitivityLabels" } else { $seg + "/security/informationProtection/sensitivityLabels" }

    switch ($sub) {
        "list" {
            $qh = Build-QueryAndHeaders $parsed.Map @()
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($path + $qh.Query) -Api $api -AllowFallback:$false
            if ($resp -and $resp.value) {
                $resp.value | ConvertTo-Json -Depth 10
            } elseif ($resp) {
                $resp | ConvertTo-Json -Depth 10
            }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: label get <id> [--user <upn|id>|--me|--org] [--beta]"
                return
            }
            $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($path + "/" + $id) -Api $api -AllowFallback:$false
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        default {
            Write-Warn "Usage: label list|get [--user <upn|id>|--me|--org] [--beta]"
        }
    }
}
