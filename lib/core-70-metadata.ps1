# Core: Metadata
# Purpose: Metadata shared utilities.
function Get-GraphMetadataInfo {
    param([switch]$Beta)
    $label = if ($Beta) { "beta" } else { "v1" }
    $dir = Join-Path $Paths.Data "graph"
    $file = Join-Path $dir ("metadata-" + $label + ".xml")
    $meta = Join-Path $dir ("metadata-" + $label + ".json")
    $url = if ($Beta) { "https://graph.microsoft.com/beta/`$metadata" } else { "https://graph.microsoft.com/v1.0/`$metadata" }
    return [pscustomobject]@{
        Label = $label
        Dir   = $dir
        File  = $file
        Meta  = $meta
        Url   = $url
    }
}



function Read-GraphMetadataManifest {
    param([switch]$Beta)
    $info = Get-GraphMetadataInfo -Beta:$Beta
    if (Test-Path $info.Meta) {
        try {
            return (Get-Content -Raw $info.Meta | ConvertFrom-Json)
        } catch {
            return $null
        }
    }
    return $null
}



function Write-GraphMetadataManifest {
    param(
        [switch]$Beta,
        [object]$Manifest
    )
    $info = Get-GraphMetadataInfo -Beta:$Beta
    $json = $Manifest | ConvertTo-Json -Depth 6
    Set-Content -Path $info.Meta -Value $json -Encoding ASCII
}



function Sync-GraphMetadata {
    param(
        [switch]$Beta,
        [switch]$Force,
        [switch]$Quiet
    )
    $info = Get-GraphMetadataInfo -Beta:$Beta
    if (-not (Test-Path $info.Dir)) {
        New-Item -ItemType Directory -Path $info.Dir -Force | Out-Null
    }

    $manifest = Read-GraphMetadataManifest -Beta:$Beta
    $etag = $null
    if ($manifest -and $manifest.ETag) { $etag = $manifest.ETag }

    $headers = @{}
    if ($etag -and -not $Force) { $headers["If-None-Match"] = $etag }

    try {
        $resp = Invoke-WebRequest -Uri $info.Url -Headers $headers -Method Get -SkipHttpErrorCheck
    } catch {
        if (-not $Quiet) {
            Write-Warn ("Metadata sync failed (" + $info.Label + "): " + $_.Exception.Message)
        }
        return $false
    }

    $status = [int]$resp.StatusCode
    $now = (Get-Date).ToString("o")
    if ($status -eq 304) {
        $newManifest = if ($manifest) { $manifest } else { [ordered]@{} }
        $newManifest.LastChecked = $now
        Write-GraphMetadataManifest -Beta:$Beta -Manifest $newManifest
        if (-not $Quiet) { Write-Info ("Metadata up-to-date: " + $info.Label) }
        return $true
    }
    if ($status -ge 200 -and $status -lt 300) {
        $content = $resp.Content
        Set-Content -Path $info.File -Value $content -Encoding UTF8
        $newManifest = [ordered]@{
            ETag        = $resp.Headers.ETag
            LastModified = $resp.Headers["Last-Modified"]
            LastChecked = $now
            Downloaded  = $now
        }
        Write-GraphMetadataManifest -Beta:$Beta -Manifest $newManifest
        if ($global:GraphMetadataCache -and $global:GraphMetadataCache.ContainsKey($info.Label)) {
            $global:GraphMetadataCache.Remove($info.Label) | Out-Null
        }
        if (-not $Quiet) { Write-Info ("Metadata synced: " + $info.Label) }
        return $true
    }

    if (-not $Quiet) {
        Write-Warn ("Metadata sync failed (" + $info.Label + "): HTTP " + $status)
    }
    return $false
}



function Sync-GraphMetadataIfNeeded {
    $cfg = $global:Config
    if (-not $cfg -or -not $cfg.graph -or -not (Parse-Bool $cfg.graph.autoSyncMetadata $true)) {
        return
    }
    $hours = [int]$cfg.graph.metadataRefreshHours
    if ($hours -le 0) { $hours = 24 }
    foreach ($beta in @($false, $true)) {
        $info = Get-GraphMetadataInfo -Beta:$beta
        $manifest = Read-GraphMetadataManifest -Beta:$beta
        $needs = $false
        if (-not (Test-Path $info.File)) {
            $needs = $true
        } elseif (-not $manifest -or -not $manifest.LastChecked) {
            $needs = $true
        } else {
            try {
                $last = [datetime]::Parse($manifest.LastChecked)
                if ((Get-Date) -gt $last.AddHours($hours)) { $needs = $true }
            } catch {
                $needs = $true
            }
        }
        if ($needs) {
            Sync-GraphMetadata -Beta:$beta -Quiet
        }
    }
}



function Get-GraphMetadataXml {
    param([switch]$Beta)
    $info = Get-GraphMetadataInfo -Beta:$Beta
    if (-not (Test-Path $info.File)) {
        Write-Warn ("Metadata file missing for " + $info.Label + ". Run: graph meta sync")
        return $null
    }
    try {
        $raw = Get-Content -Raw -Path $info.File
        return [xml]$raw
    } catch {
        Write-Warn ("Failed to parse metadata for " + $info.Label + ": " + $_.Exception.Message)
        return $null
    }
}



function Get-GraphMetadataIndex {
    param([switch]$Beta)
    if (-not $global:GraphMetadataCache) { $global:GraphMetadataCache = @{} }
    $key = if ($Beta) { "beta" } else { "v1" }
    if ($global:GraphMetadataCache.ContainsKey($key)) { return $global:GraphMetadataCache[$key] }

    $xml = Get-GraphMetadataXml -Beta:$Beta
    if (-not $xml) { return $null }

    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("edmx", "http://docs.oasis-open.org/odata/ns/edmx")
    $ns.AddNamespace("edm", "http://docs.oasis-open.org/odata/ns/edm")

    $entitySets = @()
    $entityTypes = @()
    $actions = @()
    $functions = @()
    $enums = @()
    $complex = @()

    $schemas = $xml.SelectNodes("//edm:Schema", $ns)
    foreach ($schema in $schemas) {
        $namespace = $schema.GetAttribute("Namespace")

        foreach ($et in $schema.SelectNodes("edm:EntityType", $ns)) {
            $props = @()
            foreach ($p in $et.SelectNodes("edm:Property", $ns)) {
                $props += [pscustomobject]@{
                    Name = $p.GetAttribute("Name")
                    Type = $p.GetAttribute("Type")
                    Nullable = $p.GetAttribute("Nullable")
                }
            }
            $navs = @()
            foreach ($n in $et.SelectNodes("edm:NavigationProperty", $ns)) {
                $navs += [pscustomobject]@{
                    Name = $n.GetAttribute("Name")
                    Type = $n.GetAttribute("Type")
                    ContainsTarget = $n.GetAttribute("ContainsTarget")
                }
            }
            $entityTypes += [pscustomobject]@{
                Name      = $et.GetAttribute("Name")
                Namespace = $namespace
                FullName  = ($namespace + "." + $et.GetAttribute("Name"))
                BaseType  = $et.GetAttribute("BaseType")
                Properties = $props
                Navigation = $navs
            }
        }

        foreach ($ct in $schema.SelectNodes("edm:ComplexType", $ns)) {
            $complex += [pscustomobject]@{
                Name = $ct.GetAttribute("Name")
                Namespace = $namespace
                FullName = ($namespace + "." + $ct.GetAttribute("Name"))
            }
        }

        foreach ($en in $schema.SelectNodes("edm:EnumType", $ns)) {
            $members = @()
            foreach ($m in $en.SelectNodes("edm:Member", $ns)) {
                $members += $m.GetAttribute("Name")
            }
            $enums += [pscustomobject]@{
                Name = $en.GetAttribute("Name")
                Namespace = $namespace
                FullName = ($namespace + "." + $en.GetAttribute("Name"))
                Members = $members
            }
        }

        foreach ($a in $schema.SelectNodes("edm:Action", $ns)) {
            $params = @()
            foreach ($p in $a.SelectNodes("edm:Parameter", $ns)) {
                $params += [pscustomobject]@{
                    Name = $p.GetAttribute("Name")
                    Type = $p.GetAttribute("Type")
                    Nullable = $p.GetAttribute("Nullable")
                }
            }
            $ret = $a.SelectSingleNode("edm:ReturnType", $ns)
            $actions += [pscustomobject]@{
                Name = $a.GetAttribute("Name")
                Namespace = $namespace
                FullName = ($namespace + "." + $a.GetAttribute("Name"))
                IsBound = $a.GetAttribute("IsBound")
                Parameters = $params
                ReturnType = if ($ret) { $ret.GetAttribute("Type") } else { "" }
            }
        }

        foreach ($f in $schema.SelectNodes("edm:Function", $ns)) {
            $params = @()
            foreach ($p in $f.SelectNodes("edm:Parameter", $ns)) {
                $params += [pscustomobject]@{
                    Name = $p.GetAttribute("Name")
                    Type = $p.GetAttribute("Type")
                    Nullable = $p.GetAttribute("Nullable")
                }
            }
            $ret = $f.SelectSingleNode("edm:ReturnType", $ns)
            $functions += [pscustomobject]@{
                Name = $f.GetAttribute("Name")
                Namespace = $namespace
                FullName = ($namespace + "." + $f.GetAttribute("Name"))
                IsBound = $f.GetAttribute("IsBound")
                Parameters = $params
                ReturnType = if ($ret) { $ret.GetAttribute("Type") } else { "" }
            }
        }

        $container = $schema.SelectSingleNode("edm:EntityContainer", $ns)
        if ($container) {
            foreach ($es in $container.SelectNodes("edm:EntitySet", $ns)) {
                $entitySets += [pscustomobject]@{
                    Name = $es.GetAttribute("Name")
                    EntityType = $es.GetAttribute("EntityType")
                    Namespace = $namespace
                }
            }
        }
    }

    $index = [pscustomobject]@{
        EntitySets = $entitySets
        EntityTypes = $entityTypes
        Actions = $actions
        Functions = $functions
        Enums = $enums
        ComplexTypes = $complex
    }
    $global:GraphMetadataCache[$key] = $index
    return $index
}

function Compare-GraphMetadata {
    param(
        [string]$Kind,
        [switch]$V1Only
    )
    $src = if ($V1Only) { Get-GraphMetadataIndex } else { Get-GraphMetadataIndex -Beta }
    $dst = if ($V1Only) { Get-GraphMetadataIndex -Beta } else { Get-GraphMetadataIndex }
    if (-not $src -or -not $dst) { return $null }

    $kind = if ($Kind) { $Kind.ToLowerInvariant() } else { "entityset" }
    $items = @()

    switch ($kind) {
        { $_ -in @("entityset", "entitysets") } {
            $dstNames = @{}
            foreach ($e in $dst.EntitySets) { $dstNames[$e.Name] = $true }
            foreach ($e in $src.EntitySets) {
                if (-not $dstNames.ContainsKey($e.Name)) {
                    $items += [pscustomobject]@{
                        Kind       = "EntitySet"
                        Name       = $e.Name
                        EntityType = $e.EntityType
                    }
                }
            }
        }
        { $_ -in @("entity", "entitytype", "entitytypes") } {
            $dstNames = @{}
            foreach ($e in $dst.EntityTypes) { $dstNames[$e.FullName] = $true }
            foreach ($e in $src.EntityTypes) {
                if (-not $dstNames.ContainsKey($e.FullName)) {
                    $items += [pscustomobject]@{
                        Kind     = "EntityType"
                        Name     = $e.FullName
                        BaseType = $e.BaseType
                    }
                }
            }
        }
        { $_ -in @("action", "actions") } {
            $dstNames = @{}
            foreach ($a in $dst.Actions) { $dstNames[$a.FullName] = $true }
            foreach ($a in $src.Actions) {
                if (-not $dstNames.ContainsKey($a.FullName)) {
                    $items += [pscustomobject]@{
                        Kind      = "Action"
                        Name      = $a.FullName
                        IsBound   = $a.IsBound
                        ReturnType = $a.ReturnType
                    }
                }
            }
        }
        { $_ -in @("function", "functions") } {
            $dstNames = @{}
            foreach ($f in $dst.Functions) { $dstNames[$f.FullName] = $true }
            foreach ($f in $src.Functions) {
                if (-not $dstNames.ContainsKey($f.FullName)) {
                    $items += [pscustomobject]@{
                        Kind      = "Function"
                        Name      = $f.FullName
                        IsBound   = $f.IsBound
                        ReturnType = $f.ReturnType
                    }
                }
            }
        }
        { $_ -in @("enum", "enums") } {
            $dstNames = @{}
            foreach ($e in $dst.Enums) { $dstNames[$e.FullName] = $true }
            foreach ($e in $src.Enums) {
                if (-not $dstNames.ContainsKey($e.FullName)) {
                    $items += [pscustomobject]@{
                        Kind = "Enum"
                        Name = $e.FullName
                    }
                }
            }
        }
        { $_ -in @("complex", "complextype", "complextypes") } {
            $dstNames = @{}
            foreach ($c in $dst.ComplexTypes) { $dstNames[$c.FullName] = $true }
            foreach ($c in $src.ComplexTypes) {
                if (-not $dstNames.ContainsKey($c.FullName)) {
                    $items += [pscustomobject]@{
                        Kind = "ComplexType"
                        Name = $c.FullName
                    }
                }
            }
        }
        { $_ -in @("property", "properties") } {
            $dstTypes = @{}
            foreach ($t in $dst.EntityTypes) { $dstTypes[$t.FullName] = $t }
            foreach ($t in $src.EntityTypes) {
                if (-not $dstTypes.ContainsKey($t.FullName)) { continue }
                $dstProps = @{}
                foreach ($p in $dstTypes[$t.FullName].Properties) { $dstProps[$p.Name] = $true }
                foreach ($p in $t.Properties) {
                    if (-not $dstProps.ContainsKey($p.Name)) {
                        $items += [pscustomobject]@{
                            Kind      = "Property"
                            Name      = $p.Name
                            EntityType = $t.FullName
                            Type      = $p.Type
                            Nullable  = $p.Nullable
                        }
                    }
                }
            }
        }
        { $_ -in @("nav", "navigation", "navigationproperty", "navigationproperties") } {
            $dstTypes = @{}
            foreach ($t in $dst.EntityTypes) { $dstTypes[$t.FullName] = $t }
            foreach ($t in $src.EntityTypes) {
                if (-not $dstTypes.ContainsKey($t.FullName)) { continue }
                $dstNav = @{}
                foreach ($n in $dstTypes[$t.FullName].Navigation) { $dstNav[$n.Name] = $true }
                foreach ($n in $t.Navigation) {
                    if (-not $dstNav.ContainsKey($n.Name)) {
                        $items += [pscustomobject]@{
                            Kind      = "Navigation"
                            Name      = $n.Name
                            EntityType = $t.FullName
                            Type      = $n.Type
                            ContainsTarget = $n.ContainsTarget
                        }
                    }
                }
            }
        }
        default {
            return $null
        }
    }

    return $items
}


