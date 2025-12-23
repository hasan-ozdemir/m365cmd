# Core: Query
# Purpose: Query shared utilities.
function Encode-QueryValue {
    param([string]$Value)
    if ($null -eq $Value) { return "" }
    return [System.Uri]::EscapeDataString([string]$Value)
}

function ConvertTo-QueryString {
    param([hashtable]$Params)
    if (-not $Params -or $Params.Count -eq 0) { return "" }
    $parts = @()
    foreach ($k in $Params.Keys) {
        $v = $Params[$k]
        if ($null -eq $v -or $v -eq "") { continue }
        $parts += ($k + "=" + (Encode-QueryValue $v))
    }
    if ($parts.Count -eq 0) { return "" }
    return "?" + ($parts -join "&")
}

function Build-QueryAndHeaders {
    param(
        [hashtable]$Map,
        [string[]]$SelectDefaults
    )
    $params = @{}
    $headers = @{}
    $top = Get-ArgValue $Map "top"
    $skip = Get-ArgValue $Map "skip"
    $filter = Get-ArgValue $Map "filter"
    $select = Get-ArgValue $Map "select"
    $orderby = Get-ArgValue $Map "orderby"
    $search = Get-ArgValue $Map "search"
    $expand = Get-ArgValue $Map "expand"

    if ($top) { $params['$top'] = $top }
    if ($skip) { $params['$skip'] = $skip }
    if ($filter) { $params['$filter'] = $filter }
    if ($orderby) { $params['$orderby'] = $orderby }
    if ($select) {
        $params['$select'] = $select
    } elseif ($SelectDefaults -and $SelectDefaults.Count -gt 0) {
        $params['$select'] = ($SelectDefaults -join ",")
    }
    if ($expand) { $params['$expand'] = $expand }
    if ($search) {
        if ($search -notmatch '^\".*\"$') {
            $search = '"' + $search + '"'
        }
        $params['$search'] = $search
        $params['$count'] = "true"
        $headers["ConsistencyLevel"] = "eventual"
    }

    return [pscustomobject]@{
        Query   = ConvertTo-QueryString $params
        Headers = $headers
    }
}

function Get-ArgValue {
    param(
        [hashtable]$Map,
        [string]$Name
    )
    foreach ($k in $Map.Keys) {
        if ($k.ToLowerInvariant() -eq $Name.ToLowerInvariant()) {
            return $Map[$k]
        }
    }
    return $null
}

function Escape-ODataString {
    param([string]$Value)
    if ($null -eq $Value) { return "" }
    return ($Value -replace "'", "''")
}
