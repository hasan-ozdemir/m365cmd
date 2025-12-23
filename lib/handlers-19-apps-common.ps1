# Handler: Apps Common
# Purpose: Apps Common command handlers.
function Flatten-SearchHits {
    param([object]$Response)
    $hits = @()
    foreach ($v in @($Response.value)) {
        foreach ($hc in @($v.hitsContainers)) {
            foreach ($h in @($hc.hits)) {
                if ($h.resource) { $hits += $h.resource }
            }
        }
    }
    return $hits
}

function Invoke-FileTypeSearch {
    param(
        [string]$Types,
        [hashtable]$Map
    )
    $query = Get-ArgValue $Map "query"
    $siteUrl = Get-ArgValue $Map "siteUrl"
    $path = Get-ArgValue $Map "path"
    $fromRaw = Get-ArgValue $Map "from"
    $topRaw = Get-ArgValue $Map "top"
    $useBeta = $Map.ContainsKey("beta")
    $useV1 = $Map.ContainsKey("v1")
    $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
    if ($Map.ContainsKey("auto")) { $allowFallback = $true }
    if ($useBeta -or $useV1) { $allowFallback = $false }
    $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }

    $from = 0
    if ($fromRaw) { try { $from = [int]$fromRaw } catch {} }
    $size = 25
    if ($topRaw) { try { $size = [int]$topRaw } catch {} }
    if ($size -le 0) { $size = 25 }

    $q = Build-StreamQuery -Query $query -Types $Types -SiteUrl $siteUrl -Path $path
    $resp = Invoke-StreamSearch -Query $q -From $from -Size $size -Api $api -AllowFallback:$allowFallback
    if (-not $resp) { return }
    $hits = Flatten-SearchHits $resp
    if (-not $hits -or $hits.Count -eq 0) {
        Write-Info "No items found."
        return
    }
    $asJson = $Map.ContainsKey("json")
    if ($asJson) {
        $hits | ConvertTo-Json -Depth 8
    } else {
        Write-GraphTable $hits @("Name","Id","Size","LastModifiedDateTime","WebUrl")
    }
}

function Build-FileArgsFromMap {
    param([hashtable]$Map)
    $args = @()
    if (-not $Map) { return $args }
    foreach ($k in @("user","drive","site","group","top","skip","select","filter","orderby","search","expand")) {
        $v = Get-ArgValue $Map $k
        if ($v) { $args += @("--" + $k, $v) }
    }
    if ($Map.ContainsKey("beta")) { $args += "--beta" }
    return $args
}
