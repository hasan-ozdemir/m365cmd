# Handler: Webhook
# Purpose: Webhook command handlers.
function Start-WebhookListener {
    param(
        [string]$Prefix,
        [string]$OutFile,
        [bool]$Once = $false,
        [bool]$Quiet = $false
    )
    if (-not $Prefix) { $Prefix = "http://+:5000/" }
    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add($Prefix)
    try {
        $listener.Start()
    } catch {
        Write-Err $_.Exception.Message
        return
    }
    if (-not $Quiet) { Write-Info ("Webhook listening on " + $Prefix) }
    if ($OutFile -and -not $Quiet) { Write-Info ("Output file: " + $OutFile) }

    try {
        while ($listener.IsListening) {
            $ctx = $listener.GetContext()
            $req = $ctx.Request
            $res = $ctx.Response
            $token = $req.QueryString["validationToken"]
            if ($token) {
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($token)
                $res.StatusCode = 200
                $res.ContentType = "text/plain"
                $res.OutputStream.Write($bytes, 0, $bytes.Length)
                $res.OutputStream.Close()
                if ($Once) { break }
                continue
            }
            $reader = New-Object System.IO.StreamReader($req.InputStream)
            $body = $reader.ReadToEnd()
            $reader.Close()
            $entry = [ordered]@{
                time    = (Get-Date).ToString("o")
                method  = $req.HttpMethod
                url     = $req.Url.AbsoluteUri
                headers = @{}
                body    = $body
            }
            foreach ($k in $req.Headers.AllKeys) {
                $entry.headers[$k] = $req.Headers[$k]
            }
            if ($OutFile) {
                ($entry | ConvertTo-Json -Depth 6) | Add-Content -Path $OutFile -Encoding ASCII
            }
            $res.StatusCode = 202
            $res.OutputStream.Close()
            if ($Once) { break }
        }
    } catch {
        Write-Err $_.Exception.Message
    } finally {
        try { $listener.Stop() } catch {}
        try { $listener.Close() } catch {}
        if (-not $Quiet) { Write-Info "Webhook listener stopped." }
    }
}


function Handle-WebhookCommand {
    param([string[]]$Args)
    if (-not $Args -or $Args.Count -eq 0) {
        Write-Warn "Usage: webhook listen|start|stop|status"
        return
    }
    $action = $Args[0].ToLowerInvariant()
    $rest = if ($Args.Count -gt 1) { $Args[1..($Args.Count - 1)] } else { @() }
    $parsed = Parse-NamedArgs $rest

    switch ($action) {
        "listen" {
            $port = Get-ArgValue $parsed.Map "port"
            $prefix = Get-ArgValue $parsed.Map "prefix"
            $out = Get-ArgValue $parsed.Map "out"
            $once = Parse-Bool (Get-ArgValue $parsed.Map "once") $false
            if (-not $prefix) {
                $p = if ($port) { [int]$port } else { 5000 }
                $prefix = ("http://+:" + $p + "/")
            }
            if (-not $out) { $out = Join-Path $Paths.Data "webhook-notifications.jsonl" }
            Start-WebhookListener -Prefix $prefix -OutFile $out -Once:$once
        }
        "start" {
            if ($global:WebhookJob -and $global:WebhookJob.State -eq "Running") {
                Write-Warn "Webhook listener already running. Use: webhook stop"
                return
            }
            $port = Get-ArgValue $parsed.Map "port"
            $prefix = Get-ArgValue $parsed.Map "prefix"
            $out = Get-ArgValue $parsed.Map "out"
            if (-not $prefix) {
                $p = if ($port) { [int]$port } else { 5000 }
                $prefix = ("http://+:" + $p + "/")
            }
            if (-not $out) { $out = Join-Path $Paths.Data "webhook-notifications.jsonl" }
            $sb = {
                param($Prefix, $OutFile)
                $listener = [System.Net.HttpListener]::new()
                $listener.Prefixes.Add($Prefix)
                $listener.Start()
                try {
                    while ($listener.IsListening) {
                        $ctx = $listener.GetContext()
                        $req = $ctx.Request
                        $res = $ctx.Response
                        $token = $req.QueryString["validationToken"]
                        if ($token) {
                            $bytes = [System.Text.Encoding]::UTF8.GetBytes($token)
                            $res.StatusCode = 200
                            $res.ContentType = "text/plain"
                            $res.OutputStream.Write($bytes, 0, $bytes.Length)
                            $res.OutputStream.Close()
                            continue
                        }
                        $reader = New-Object System.IO.StreamReader($req.InputStream)
                        $body = $reader.ReadToEnd()
                        $reader.Close()
                        $entry = [ordered]@{
                            time    = (Get-Date).ToString("o")
                            method  = $req.HttpMethod
                            url     = $req.Url.AbsoluteUri
                            headers = @{}
                            body    = $body
                        }
                        foreach ($k in $req.Headers.AllKeys) {
                            $entry.headers[$k] = $req.Headers[$k]
                        }
                        ($entry | ConvertTo-Json -Depth 6) | Add-Content -Path $OutFile -Encoding ASCII
                        $res.StatusCode = 202
                        $res.OutputStream.Close()
                    }
                } finally {
                    try { $listener.Stop() } catch {}
                    try { $listener.Close() } catch {}
                }
            }
            $global:WebhookJob = Start-Job -ScriptBlock $sb -ArgumentList $prefix, $out
            Write-Info ("Webhook job started (ID " + $global:WebhookJob.Id + ").")
        }
        "stop" {
            if (-not $global:WebhookJob) {
                Write-Warn "No webhook job running."
                return
            }
            try {
                Stop-Job -Job $global:WebhookJob -Force | Out-Null
                Remove-Job -Job $global:WebhookJob -Force | Out-Null
                $global:WebhookJob = $null
                Write-Info "Webhook job stopped."
            } catch {
                Write-Err $_.Exception.Message
            }
        }
        "status" {
            if ($global:WebhookJob) {
                Write-Host ("Webhook job: " + $global:WebhookJob.State + " (ID " + $global:WebhookJob.Id + ")")
            } else {
                Write-Host "Webhook job: not running"
            }
        }
        default {
            Write-Warn "Usage: webhook listen|start|stop|status"
        }
    }
}
