function Invoke-CommandLine {
    param(
        [string]$Line,
        [int]$Depth = 0
    )
    if ($Depth -gt 5) {
        Write-Warn "Alias expansion too deep."
        return $true
    }
    if (-not $Line) { return $true }
    $line = $Line.Trim()
    if ($line.Length -eq 0) { return $true }

    $isGlobal = $line.StartsWith("/")
    $body = if ($isGlobal) { $line.Substring(1) } else { $line }
    $parts = Split-Args $body
    if (-not $parts -or $parts.Count -eq 0) { return $true }
    $cmd = $parts[0].ToLowerInvariant()
    $args = if ($parts.Count -gt 1) { $parts[1..($parts.Count - 1)] } else { @() }

    $expanded = Expand-AliasCommand -Cmd $cmd -Args $args -IsGlobal:$isGlobal
    if ($expanded -and $expanded.Count -gt 0) {
        foreach ($lineItem in $expanded) {
            $cont = Invoke-CommandLine -Line $lineItem -Depth ($Depth + 1)
            if (-not $cont) { return $false }
        }
        return $true
    }

    if ($isGlobal) {
        return Handle-GlobalCommand $cmd $args
    }
    Handle-LocalCommand $cmd $args
    return $true
}


function Start-M365Cmd {
    Ensure-Directories
    Set-LocalModulePath
    $global:Config = Normalize-Config (Load-Config)

    Sync-GraphMetadataIfNeeded

    Write-Host "m365cmd REPL started. Type /help for commands."

    $running = $true
    while ($running) {
        try {
            $line = Read-Host "m365cmd"
        } catch {
            break
        }
        if ($null -eq $line) { continue }
        $line = $line.Trim()
        if ($line.Length -eq 0) { continue }

        foreach ($segment in (Split-CommandSequence $line)) {
            $running = Invoke-CommandLine -Line $segment -Depth 0
            if (-not $running) { break }
        }
    }
}
