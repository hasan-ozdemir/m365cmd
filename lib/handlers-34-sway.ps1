# Handler: Sway
# Purpose: Sway command handlers.
function Handle-SwayCommand {
    param([string[]]$InputArgs)
    if (-not $InputArgs -or $InputArgs.Count -eq 0) {
        Write-Warn "Usage: sway open [swayId|url] | sway uri --command <cmd> [--param k=v[,k=v]] [--url <url>]"
        return
    }
    $sub = $InputArgs[0].ToLowerInvariant()
    $rest = if ($InputArgs.Count -gt 1) { $InputArgs[1..($InputArgs.Count - 1)] } else { @() }
    switch ($sub) {
        "open" {
            $id = $rest | Select-Object -First 1
            if ($id) {
                if ($id -match "^https?://") {
                    Write-Host $id
                } else {
                    Write-Host ("https://sway.office.com/" + $id)
                }
            } else {
                Write-Host "https://sway.office.com/"
            }
        }
        "uri" {
            $parsed = Parse-NamedArgs $rest
            $cmd = Get-ArgValue $parsed.Map "command"
            if (-not $cmd) { $cmd = "open" }
            $paramRaw = Get-ArgValue $parsed.Map "param"
            $url = Get-ArgValue $parsed.Map "url"
            $kv = @{}
            if ($paramRaw) { $kv = Parse-KvPairs $paramRaw }
            if ($url) { $kv["url"] = $url }
            $qs = ""
            if ($kv.Keys.Count -gt 0) { $qs = ConvertTo-QueryString $kv }
            $uri = "ms-sway:" + $cmd + $qs
            Write-Host $uri
        }
        default { Write-Warn "Usage: sway open [swayId|url] | sway uri --command <cmd> [--param k=v[,k=v]] [--url <url>]" }
    }
}

