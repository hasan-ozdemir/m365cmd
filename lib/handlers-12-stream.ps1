# Handler: Stream
# Purpose: Stream command handlers.
function Build-StreamQuery {
    param(
        [string]$Query,
        [string]$Types,
        [string]$SiteUrl,
        [string]$Path
    )
    $q = $Query
    if (-not $q) {
        $list = @()
        $custom = Parse-CommaList $Types
        if ($custom.Count -gt 0) {
            $list = $custom
        } else {
            $list = @("mp4","mov","wmv","webm","m4v","avi","mpg","mpeg","mkv")
        }
        $clauses = @()
        foreach ($t in $list) {
            if (-not $t) { continue }
            $clauses += ("filetype:" + $t)
        }
        $q = ($clauses -join " OR ")
    }
    if ($SiteUrl) {
        $q = $q + " path:`"" + $SiteUrl.TrimEnd("/") + "`""
    }
    if ($Path) {
        $q = $q + " path:`"" + $Path + "`""
    }
    return $q
}

function Invoke-StreamSearch {
    param(
        [string]$Query,
        [int]$From,
        [int]$Size,
        [string]$Api,
        [switch]$AllowFallback
    )
    $req = @{
        requests = @(
            @{
                entityTypes = @("driveItem")
                query       = @{ queryString = $Query }
                from        = $From
                size        = $Size
            }
        )
    }
    return Invoke-GraphRequestAuto -Method "POST" -Uri "/search/query" -Body $req -Api $Api -AllowFallback:$AllowFallback
}

function Handle-StreamCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: stream open|list|search|file ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }

    if ($action -eq "open") {
        Write-Host "https://stream.microsoft.com/"
        return
    }

    $fileOps = @("file","item","get","list","create","update","delete","download","convert","preview","upload","copy","move","share")
    if ($fileOps -contains $action) {
        $sub = if ($action -in @("file","item")) { $rest } else { @($action) + $rest }
        Handle-FileCommand $sub
        return
    }

    switch ($action) {
        "list" { }
        "search" { }
        default {
            Write-Warn "Usage: stream open|list|search|file ..."
            return
        }
    }

    $parsed = Parse-NamedArgs $rest
    $query = Get-ArgValue $parsed.Map "query"
    $types = Get-ArgValue $parsed.Map "types"
    $siteUrl = Get-ArgValue $parsed.Map "siteUrl"
    $path = Get-ArgValue $parsed.Map "path"
    $fromRaw = Get-ArgValue $parsed.Map "from"
    $topRaw = Get-ArgValue $parsed.Map "top"
    $useBeta = $parsed.Map.ContainsKey("beta")
    $useV1 = $parsed.Map.ContainsKey("v1")
    $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
    if ($parsed.Map.ContainsKey("auto")) { $allowFallback = $true }
    if ($useBeta -or $useV1) { $allowFallback = $false }
    $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }

    $from = 0
    if ($fromRaw) { try { $from = [int]$fromRaw } catch {} }
    $size = 25
    if ($topRaw) { try { $size = [int]$topRaw } catch {} }
    if ($size -le 0) { $size = 25 }

    $q = Build-StreamQuery -Query $query -Types $types -SiteUrl $siteUrl -Path $path
    $resp = Invoke-StreamSearch -Query $q -From $from -Size $size -Api $api -AllowFallback:$allowFallback
    if (-not $resp) { return }

    $hits = @()
    foreach ($v in @($resp.value)) {
        foreach ($hc in @($v.hitsContainers)) {
            foreach ($h in @($hc.hits)) {
                if ($h.resource) { $hits += $h.resource }
            }
        }
    }
    if (-not $hits -or $hits.Count -eq 0) {
        Write-Info "No items found."
        return
    }
    $asJson = $parsed.Map.ContainsKey("json")
    if ($asJson) {
        $hits | ConvertTo-Json -Depth 8
    } else {
        Write-GraphTable $hits @("Name","Id","Size","LastModifiedDateTime","WebUrl")
    }
}
