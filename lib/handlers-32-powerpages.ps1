# Handler: Powerpages
# Purpose: Powerpages command handlers.
function Invoke-PpRequestWithHeaders {
    param(
        [string]$Method,
        [string]$Path,
        [object]$Body,
        [hashtable]$Headers
    )
    $token = Get-PpToken
    if (-not $token) { return $null }
    $base = $global:Config.pp.baseUrl.TrimEnd("/")
    $url = if ($Path -match "^https?://") { $Path } else { $base + $Path }
    $apiVer = $global:Config.pp.apiVersion
    if ($url -notmatch "api-version=") {
        $join = if ($url -match "\\?") { "&" } else { "?" }
        $url = $url + $join + "api-version=" + $apiVer
    }
    $hdr = @{ Authorization = "Bearer " + $token }
    if ($Headers) {
        foreach ($k in $Headers.Keys) { $hdr[$k] = $Headers[$k] }
    }
    $params = @{ Method = $Method; Uri = $url; Headers = $hdr }
    if ($Body -ne $null) {
        $params.ContentType = "application/json"
        if ($Body -is [string]) {
            $params.Body = $Body
        } else {
            $params.Body = ($Body | ConvertTo-Json -Depth 10)
        }
    }
    try {
        $resp = Invoke-WebRequest @params
    } catch {
        Write-Err $_.Exception.Message
        return $null
    }
    $obj = [ordered]@{
        StatusCode = $resp.StatusCode
        Headers    = $resp.Headers
        Raw        = $resp.Content
    }
    if ($resp.Content) {
        try {
            $obj["Body"] = ($resp.Content | ConvertFrom-Json)
        } catch {
            $obj["Body"] = $null
        }
    }
    return [pscustomobject]$obj
}

function Handle-PowerPagesCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: powerpages site list|get|create|delete|restart --env <environmentId> OR powerpages op status --url <operationUrl>"
        return
    }
    $sub = $Args[0].ToLowerInvariant()
    $action = $sub
    $rest = @()
    if ($sub -eq "op" -or $sub -eq "operation") {
        if ($Args.Count -lt 2) {
            Write-Warn "Usage: powerpages op status --url <operationUrl>"
            return
        }
        $action = $Args[1].ToLowerInvariant()
        $rest = if ($Args.Count -gt 2) { $Args[2..($Args.Count - 1)] } else { @() }
        switch ($action) {
            "status" {
                $parsed = Parse-NamedArgs $rest
                $url = Get-ArgValue $parsed.Map "url"
                if (-not $url) {
                    $url = $parsed.Positionals | Select-Object -First 1
                }
                if (-not $url) {
                    Write-Warn "Usage: powerpages op status --url <operationUrl>"
                    return
                }
                $resp = Invoke-PpRequest -Method "GET" -Path $url
                if ($resp) { $resp | ConvertTo-Json -Depth 10 }
                return
            }
            default {
                Write-Warn "Usage: powerpages op status --url <operationUrl>"
                return
            }
        }
    }
    if ($sub -eq "site" -or $sub -eq "sites") {
        if ($Args.Count -lt 2) {
            Write-Warn "Usage: powerpages site list|get|create|delete|restart --env <environmentId>"
            return
        }
        $action = $Args[1].ToLowerInvariant()
        $rest = if ($Args.Count -gt 2) { $Args[2..($Args.Count - 1)] } else { @() }
    } else {
        $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    }
    $parsed = Parse-NamedArgs $rest
    $env = Get-ArgValue $parsed.Map "env"
    if (-not $env) { $env = Get-ArgValue $parsed.Map "environment" }
    if (-not $env) {
        Write-Warn "Usage: powerpages site list|get|create|delete|restart --env <environmentId>"
        return
    }
    $base = "/powerpages/environments/" + $env + "/websites"
    switch ($action) {
        "list" {
            $query = @()
            $skip = Get-ArgValue $parsed.Map "skip"
            if (-not $skip) { $skip = Get-ArgValue $parsed.Map "skiptoken" }
            if ($skip) { $query += "skip=" + (Encode-QueryValue $skip) }
            $path = $base + (if ($query.Count -gt 0) { "?" + ($query -join "&") } else { "" })
            $resp = Invoke-PpRequest -Method "GET" -Path $path
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "get" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: powerpages site get <siteId> --env <environmentId>"
                return
            }
            $resp = Invoke-PpRequest -Method "GET" -Path ($base + "/" + $id)
            if ($resp) { $resp | ConvertTo-Json -Depth 10 }
        }
        "create" {
            $jsonRaw = Get-ArgValue $parsed.Map "json"
            $body = $null
            if ($jsonRaw) {
                $body = Parse-Value $jsonRaw
            } else {
                $name = Get-ArgValue $parsed.Map "name"
                $subdomain = Get-ArgValue $parsed.Map "subdomain"
                $language = Get-ArgValue $parsed.Map "language"
                $template = Get-ArgValue $parsed.Map "template"
                $orgId = Get-ArgValue $parsed.Map "dataverseOrgId"
                if (-not $orgId) { $orgId = Get-ArgValue $parsed.Map "orgId" }
                $websiteRecordId = Get-ArgValue $parsed.Map "websiteRecordId"
                if (-not $name -or -not $subdomain -or -not $language -or -not $template -or -not $orgId) {
                    Write-Warn "Usage: powerpages site create --env <environmentId> --name <name> --subdomain <name> --language <id> --template <name> --dataverseOrgId <id> [--websiteRecordId <id>] OR --json <payload>"
                    return
                }
                $body = @{
                    name                    = $name
                    subdomain               = $subdomain
                    selectedBaseLanguage    = [int]$language
                    templateName            = $template
                    dataverseOrganizationId = $orgId
                }
                if ($websiteRecordId) { $body.websiteRecordId = $websiteRecordId }
            }
            $resp = Invoke-PpRequestWithHeaders -Method "POST" -Path $base -Body $body
            if ($resp) {
                if ($resp.Body) {
                    $resp.Body | ConvertTo-Json -Depth 10
                } elseif ($resp.Raw) {
                    Write-Host $resp.Raw
                } else {
                    Write-Info "Create request submitted."
                }
                $op = $resp.Headers["Operation-Location"]
                if ($op) { Write-Host ("Operation-Location: " + $op) }
            }
        }
        "delete" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: powerpages site delete <siteId> --env <environmentId> [--force] [--operation <url>]"
                return
            }
            $force = Parse-Bool (Get-ArgValue $parsed.Map "force") $false
            if (-not $force) {
                $confirm = Read-Host "Type DELETE to confirm"
                if ($confirm -ne "DELETE") {
                    Write-Info "Canceled."
                    return
                }
            }
            $headers = @{}
            $op = Get-ArgValue $parsed.Map "operation"
            if ($op) { $headers["Operation-Location"] = $op }
            $resp = Invoke-PpRequestWithHeaders -Method "DELETE" -Path ($base + "/" + $id) -Headers $headers
            if ($resp) {
                Write-Info "Delete request submitted."
                $op2 = $resp.Headers["Operation-Location"]
                if ($op2) { Write-Host ("Operation-Location: " + $op2) }
            }
        }
        "restart" {
            $id = $parsed.Positionals | Select-Object -First 1
            if (-not $id) {
                Write-Warn "Usage: powerpages site restart <siteId> --env <environmentId>"
                return
            }
            $resp = Invoke-PpRequest -Method "POST" -Path ($base + "/" + $id + "/restart") -AllowNullResponse
            if ($resp -ne $null) { Write-Info "Restart requested." }
        }
        default {
            Write-Warn "Usage: powerpages site list|get|create|delete|restart --env <environmentId> OR powerpages op status --url <operationUrl>"
        }
    }
}
