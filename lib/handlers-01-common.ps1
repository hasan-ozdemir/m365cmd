# Handler: Common
# Purpose: Common command handlers.
function Get-GraphEtag {
    param([string]$Path)
    try {
        $resp = Invoke-GraphRequest -Method "GET" -Uri $Path
        if ($resp -and $resp."@odata.etag") { return $resp."@odata.etag" }
    } catch {}
    return $null
}



function Read-JsonPayload {
    param(
        [string]$JsonRaw,
        [string]$BodyFile,
        [string]$SetRaw
    )
    if ($BodyFile) {
        if (-not (Test-Path $BodyFile)) {
            Write-Warn "Body file not found."
            return $null
        }
        $raw = Get-Content -Raw -Path $BodyFile
        $obj = Parse-Value $raw
        return $obj
    }
    if ($JsonRaw) {
        return (Parse-Value $JsonRaw)
    }
    if ($SetRaw) {
        $obj = Parse-Value $SetRaw
        if ($obj -is [string]) { $obj = Parse-KvPairs $SetRaw }
        return $obj
    }
    return $null
}


function Resolve-CmdletParams {
    param([object]$Parsed)
    if (-not $Parsed) { return @{} }
    $json = Get-ArgValue $Parsed.Map "json"
    $bodyFile = Get-ArgValue $Parsed.Map "bodyFile"
    $paramsRaw = Get-ArgValue $Parsed.Map "params"
    $setRaw = Get-ArgValue $Parsed.Map "set"
    $body = $null
    if ($json -or $bodyFile -or $setRaw) {
        $body = Read-JsonPayload $json $bodyFile $setRaw
    } elseif ($paramsRaw) {
        $body = Parse-KvPairs $paramsRaw
    } else {
        $body = @{}
    }
    if (-not $body) { $body = @{} }
    if ($body -isnot [hashtable]) {
        $tmp = @{}
        foreach ($p in $body.PSObject.Properties) {
            $tmp[$p.Name] = $p.Value
        }
        $body = $tmp
    }
    return $body
}



