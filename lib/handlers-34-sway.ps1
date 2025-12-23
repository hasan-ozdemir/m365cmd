# Handler: Sway
# Purpose: Sway command handlers.
function Handle-SwayCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: sway open [swayId|url] | sway uri --command <cmd> [--param k=v[,k=v]] [--url <url>]"
        return
    }
    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
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
