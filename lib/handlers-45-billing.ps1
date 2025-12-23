# Handler: Billing
# Purpose: Billing command handlers.
function Handle-BillingCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: billing sku|subscription ..."
        return
    }
    if (-not (Require-GraphConnection)) { return }

    $sub = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    if (-not $rest -or $rest.Count -eq 0) {
        Write-Warn "Usage: billing sku|subscription ..."
        return
    }
    $action = $rest[0].ToLowerInvariant()
    $rest2 = if ($rest.Count -gt 1) { $rest[1..($rest.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest2
    $apiInfo = Resolve-CAApiSettings $parsed

    switch ($sub) {
        "sku" {
            $base = "/subscribedSkus"
            switch ($action) {
                "list" {
                    $qh = Build-QueryAndHeaders $parsed.Map @("skuPartNumber", "skuId", "consumedUnits", "capabilityStatus")
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp -and $resp.value) {
                        Write-GraphTable $resp.value @("SkuPartNumber", "SkuId", "ConsumedUnits", "CapabilityStatus")
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 10
                    }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    if (-not $id) {
                        Write-Warn "Usage: billing sku get <skuId>"
                        return
                    }
                    $qh = Build-QueryAndHeaders $parsed.Map @()
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + "/" + $id + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                default {
                    Write-Warn "Usage: billing sku list|get"
                }
            }
        }
        "subscription" {
            $base = "/directory/subscriptions"
            switch ($action) {
                "list" {
                    $qh = Build-QueryAndHeaders $parsed.Map @("id", "skuId", "serviceStatus", "totalLicenses")
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($base + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp -and $resp.value) {
                        $resp.value | ConvertTo-Json -Depth 10
                    } elseif ($resp) {
                        $resp | ConvertTo-Json -Depth 10
                    }
                }
                "get" {
                    $id = $parsed.Positionals | Select-Object -First 1
                    $commerce = Get-ArgValue $parsed.Map "commerce"
                    if (-not $id -and -not $commerce) {
                        Write-Warn "Usage: billing subscription get <id> [--commerce <commerceSubscriptionId>]"
                        return
                    }
                    $path = if ($commerce) { $base + "(commerceSubscriptionId='" + $commerce + "')" } else { $base + "/" + $id }
                    $qh = Build-QueryAndHeaders $parsed.Map @()
                    $resp = Invoke-GraphRequestAuto -Method "GET" -Uri ($path + $qh.Query) -Headers $qh.Headers -Api $apiInfo.Api -AllowFallback:$apiInfo.AllowFallback
                    if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                }
                default {
                    Write-Warn "Usage: billing subscription list|get"
                }
            }
        }
        default {
            Write-Warn "Usage: billing sku|subscription ..."
        }
    }
}
