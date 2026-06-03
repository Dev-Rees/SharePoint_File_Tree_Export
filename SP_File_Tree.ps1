
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$AdminUrl,
    [Parameter(Mandatory = $true)] [string]$ClientId,
    [Parameter()] [string]$Tenant,
    [Parameter()] [string]$CertificateThumbprint,
    [Parameter()] [string[]]$SiteUrl,
    [Parameter()] [string]$OutputPath = (Get-Location).Path,
    [Parameter()] [int]$MaxDepth = 0,
    [Parameter()] [switch]$AlsoExportCsv
)
 
# ---------------------------------------------------------------------------
# Module check
# ---------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "PnP.PowerShell not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}
Import-Module PnP.PowerShell -ErrorAction Stop
 
# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Connect-Spo([string]$Url) {
    if ($CertificateThumbprint) {
        Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $Tenant `
            -Thumbprint $CertificateThumbprint -ErrorAction Stop
    }
    else {
        Connect-PnPOnline -Url $Url -ClientId $ClientId -Interactive -ErrorAction Stop
    }
}
 
function Format-Size([long]$bytes) {
    if ($bytes -ge 1TB) { return ("{0:N2} TB" -f ($bytes / 1TB)) }
    elseif ($bytes -ge 1GB) { return ("{0:N2} GB" -f ($bytes / 1GB)) }
    elseif ($bytes -ge 1MB) { return ("{0:N2} MB" -f ($bytes / 1MB)) }
    elseif ($bytes -ge 1KB) { return ("{0:N2} KB" -f ($bytes / 1KB)) }
    else { return "$bytes B" }
}
 
function New-Node([string]$name, [bool]$isFolder) {
    [pscustomobject]@{
        Name     = $name
        IsFolder = $isFolder
        Size     = [long]0   # own file size (files only)
        Total    = [long]0   # computed roll-up
        Children = [System.Collections.Specialized.OrderedDictionary]::new()
    }
}
 
function Add-Item($root, [string[]]$segments, [long]$size, [bool]$isFolder) {
    $node = $root
    for ($i = 0; $i -lt $segments.Count; $i++) {
        $seg = $segments[$i]
        $isLast = ($i -eq $segments.Count - 1)
        if (-not $node.Children.Contains($seg)) {
            $segIsFolder = if ($isLast) { $isFolder } else { $true }
            $node.Children[$seg] = New-Node $seg $segIsFolder
        }
        $node = $node.Children[$seg]
    }
    if (-not $isFolder) { $node.Size = $size }
}
 
function Resolve-Totals($node) {
    if (-not $node.IsFolder) { $node.Total = $node.Size; return $node.Total }
    $sum = [long]0
    foreach ($k in $node.Children.Keys) { $sum += (Resolve-Totals $node.Children[$k]) }
    $node.Total = $sum
    return $sum
}
 
function Write-Tree($node, [string]$prefix, [System.Text.StringBuilder]$sb, [int]$depth) {
    $keys = @($node.Children.Keys)
    $sorted = $keys | Sort-Object `
        @{ Expression = { -not $node.Children[$_].IsFolder } }, `
        @{ Expression = { $_ } }
 
    for ($i = 0; $i -lt $sorted.Count; $i++) {
        $key = $sorted[$i]
        $child = $node.Children[$key]
        $isLast = ($i -eq $sorted.Count - 1)
        $connector = if ($isLast) { "`u{2514}`u{2500}`u{2500} " } else { "`u{251C}`u{2500}`u{2500} " }
        $label = if ($child.IsFolder) { "$($child.Name)/" } else { $child.Name }
        [void]$sb.AppendLine("$prefix$connector$label  ($(Format-Size $child.Total))")
 
        $atDepthLimit = ($MaxDepth -gt 0 -and $depth + 1 -ge $MaxDepth)
        if ($child.IsFolder -and $child.Children.Count -gt 0 -and -not $atDepthLimit) {
            $childPrefix = $prefix + $(if ($isLast) { "    " } else { "`u{2502}   " })
            Write-Tree $child $childPrefix $sb ($depth + 1)
        }
    }
}
 
# ---------------------------------------------------------------------------
# Discover sites
# ---------------------------------------------------------------------------
try {
    Write-Host "Connecting to $AdminUrl ..." -ForegroundColor Cyan
    Connect-Spo $AdminUrl
}
catch {
    Write-Error "Failed to connect to the admin endpoint: $_"
    return
}
 
if ($SiteUrl) {
    $targets = $SiteUrl
}
else {
    Write-Host "Enumerating site collections..." -ForegroundColor Cyan
    $targets = (Get-PnPTenantSite | Where-Object { $_.Template -notlike "RedirectSite*" }).Url
}
Write-Host "Processing $($targets.Count) site(s)." -ForegroundColor Cyan
 
# ---------------------------------------------------------------------------
# Walk each site
# ---------------------------------------------------------------------------
$sb = [System.Text.StringBuilder]::new()
$csvRows = [System.Collections.Generic.List[object]]::new()
$grandTotal = [long]0
 
foreach ($url in $targets) {
    try {
        Connect-Spo $url
    }
    catch {
        Write-Warning "Skipping $url (connect failed): $_"
        continue
    }
 
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("=" * 100)
    [void]$sb.AppendLine("SITE: $url")
    [void]$sb.AppendLine("=" * 100)
 
    $lists = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden }
    $siteTotal = [long]0
 
    foreach ($list in $lists) {
        Get-PnPProperty -ClientObject $list -Property RootFolder | Out-Null
        $rootUrl = $list.RootFolder.ServerRelativeUrl
 
        $items = Get-PnPListItem -List $list -PageSize 2000 `
            -Fields FileRef, FileLeafRef, FSObjType, SMTotalFileStreamSize, File_x0020_Size, Modified
 
        if (-not $items) { continue }
 
        $libRoot = New-Node $list.Title $true
 
        foreach ($item in $items) {
            $fv = $item.FieldValues
            $ref = [string]$fv.FileRef
            if ([string]::IsNullOrEmpty($ref)) { continue }
 
            $rel = $ref
            if ($rel.StartsWith($rootUrl)) { $rel = $rel.Substring($rootUrl.Length) }
            $rel = $rel.TrimStart('/')
            if ([string]::IsNullOrEmpty($rel)) { continue }  # the library root itself
 
            $segments = $rel -split '/'
            $isFolder = ([int]$fv.FSObjType -eq 1)
 
            $size = [long]0
            if ($fv.SMTotalFileStreamSize) { $size = [long]$fv.SMTotalFileStreamSize }
            elseif ($fv.File_x0020_Size) { $size = [long]$fv.File_x0020_Size }
 
            Add-Item $libRoot $segments $size $isFolder
 
            if (-not $isFolder -and $AlsoExportCsv) {
                $csvRows.Add([pscustomobject]@{
                        Site         = $url
                        Library      = $list.Title
                        Path         = $rel
                        FileName     = [string]$fv.FileLeafRef
                        SizeBytes    = $size
                        Size         = Format-Size $size
                        LastModified = $fv.Modified
                    })
            }
        }
 
        Resolve-Totals $libRoot | Out-Null
        $siteTotal += $libRoot.Total
 
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("$($libRoot.Name)/  ($(Format-Size $libRoot.Total))")
        Write-Tree $libRoot "" $sb 0
    }
 
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine(">> Site total: $(Format-Size $siteTotal)")
    Write-Host ("  {0,-60} {1}" -f $url, (Format-Size $siteTotal)) -ForegroundColor Gray
    $grandTotal += $siteTotal
}
 
# ---------------------------------------------------------------------------
# Write outputs
# ---------------------------------------------------------------------------
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$treePath = Join-Path $OutputPath "SPO-StorageTree-$timestamp.txt"
[void]$sb.AppendLine("")
[void]$sb.AppendLine("=" * 100)
[void]$sb.AppendLine("GRAND TOTAL across $($targets.Count) site(s): $(Format-Size $grandTotal)")
$sb.ToString() | Out-File -FilePath $treePath -Encoding UTF8
 
Write-Host "`nTree written to: $treePath" -ForegroundColor Green
 
if ($AlsoExportCsv -and $csvRows.Count -gt 0) {
    $csvPath = Join-Path $OutputPath "SPO-FileInventory-$timestamp.csv"
    $csvRows | Sort-Object SizeBytes -Descending |
        Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "CSV written to:  $csvPath  ($($csvRows.Count) files)" -ForegroundColor Green
}
 
Write-Host "Grand total: $(Format-Size $grandTotal)" -ForegroundColor Green
Disconnect-PnPOnline
