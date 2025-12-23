# Handler: Common
# Purpose: Common command handlers.
function Resolve-CAApiSettings {
    param([object]$Parsed)
    $useBeta = $Parsed.Map.ContainsKey("beta")
    $useV1 = $Parsed.Map.ContainsKey("v1")
    $allowFallback = Parse-Bool $global:Config.graph.fallbackToBeta $true
    if ($Parsed.Map.ContainsKey("auto")) { $allowFallback = $true }
    if ($useBeta -or $useV1) { $allowFallback = $false }
    $api = if ($useBeta) { "beta" } elseif ($useV1) { "v1" } else { "" }
    return [pscustomobject]@{ Api = $api; AllowFallback = $allowFallback }
}
