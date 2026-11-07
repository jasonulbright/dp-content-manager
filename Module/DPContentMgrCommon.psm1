<#
.SYNOPSIS
    Core module for DP Content Manager.

.DESCRIPTION
    Import this module to get:
      - Structured logging and CM site connection management via the
        vendored SuiteCommon module (Lib\SuiteCommon)
      - Distribution point and content object retrieval
      - Bulk distribution status queries via WMI
      - Redistribution, removal, and validation actions
      - Orphaned content detection and storage analysis
      - Export to CSV, HTML, and clipboard summary

.EXAMPLE
    Import-Module "$PSScriptRoot\Module\DPContentMgrCommon.psd1" -Force
    Initialize-Logging -LogPath "C:\temp\dpcm.log"
    Connect-CMSite -SiteCode 'MCM' -SMSProvider 'sccm.domain.com'
#>

# ---------------------------------------------------------------------------
# Shared core (vendored SuiteCommon)
# ---------------------------------------------------------------------------
# Logging (Initialize-Logging, Write-Log), CM connection (Connect-CMSite,
# Disconnect-CMSite, Test-CMConnection, Get-CMConnectionInfo), and settings
# persistence come from the vendored copy at Lib\SuiteCommon\. -Global makes
# the functions resolvable from the shell script and from this module alike;
# the guard keeps a -Force reimport of this module (the background runspace
# does one) from resetting SuiteCommon state mid-session.
if (-not (Get-Module SuiteCommon)) {
    Import-Module (Join-Path $PSScriptRoot '..\Lib\SuiteCommon\SuiteCommon.psd1') -Global -DisableNameChecking
}

# ---------------------------------------------------------------------------
# Content Retrieval
# ---------------------------------------------------------------------------

function Get-AllDistributionPoints {
    <#
    .SYNOPSIS
        Returns all distribution points as normalized objects.
    #>
    Write-Log "Querying distribution points..."

    $dps = Get-CMDistributionPoint -ErrorAction Stop

    $results = foreach ($dp in $dps) {
        $serverName = ''
        if ($dp.NetworkOSPath -match '\\\\(.+)') {
            $serverName = $Matches[1]
        }

        [PSCustomObject]@{
            Name      = $serverName
            ServerName = $serverName
            SiteCode  = $dp.SiteCode
            IsPullDP  = [bool]$dp.IsPullDP
            NALPath   = $dp.NALPath
        }
    }

    Write-Log "Found $($results.Count) distribution points"
    return $results
}

function Get-ContentApplications {
    <#
    .SYNOPSIS
        Returns all applications with content properties.
    #>
    Write-Log "Loading applications..."
    $apps = Get-CMApplication -ErrorAction SilentlyContinue

    $results = foreach ($app in $apps) {
        [PSCustomObject]@{
            ContentName   = $app.LocalizedDisplayName
            ContentType   = 'Application'
            PackageID     = $app.PackageID
            SourceSize    = $app.PackageSize
            SourceVersion = $app.SourceDate
            ObjectID      = $app.CI_ID
        }
    }

    Write-Log "  $($results.Count) applications"
    return $results
}

function Get-ContentPackages {
    <#
    .SYNOPSIS
        Returns all legacy packages with content properties.
    #>
    Write-Log "Loading packages..."
    $pkgs = Get-CMPackage -Fast -ErrorAction SilentlyContinue

    $results = foreach ($pkg in $pkgs) {
        [PSCustomObject]@{
            ContentName   = $pkg.Name
            ContentType   = 'Package'
            PackageID     = $pkg.PackageID
            SourceSize    = $pkg.PackageSize
            SourceVersion = $pkg.SourceDate
            ObjectID      = $pkg.PackageID
        }
    }

    Write-Log "  $($results.Count) packages"
    return $results
}

function Get-ContentSUDPs {
    <#
    .SYNOPSIS
        Returns all software update deployment packages.
    #>
    Write-Log "Loading software update deployment packages..."
    $sudps = Get-CMSoftwareUpdateDeploymentPackage -ErrorAction SilentlyContinue

    $results = foreach ($sudp in $sudps) {
        [PSCustomObject]@{
            ContentName   = $sudp.Name
            ContentType   = 'SU Deployment Pkg'
            PackageID     = $sudp.PackageID
            SourceSize    = $sudp.PackageSize
            SourceVersion = $sudp.SourceDate
            ObjectID      = $sudp.PackageID
        }
    }

    Write-Log "  $($results.Count) software update deployment packages"
    return $results
}

function Get-ContentBootImages {
    <#
    .SYNOPSIS
        Returns all boot images with content properties.
    #>
    Write-Log "Loading boot images..."
    $bimgs = Get-CMBootImage -ErrorAction SilentlyContinue

    $results = foreach ($bi in $bimgs) {
        [PSCustomObject]@{
            ContentName   = $bi.Name
            ContentType   = 'Boot Image'
            PackageID     = $bi.PackageID
            SourceSize    = $bi.PackageSize
            SourceVersion = $bi.SourceDate
            ObjectID      = $bi.PackageID
        }
    }

    Write-Log "  $($results.Count) boot images"
    return $results
}

function Get-ContentOSImages {
    <#
    .SYNOPSIS
        Returns all OS images with content properties.
    #>
    Write-Log "Loading OS images..."
    $osimgs = Get-CMOperatingSystemImage -ErrorAction SilentlyContinue

    $results = foreach ($osi in $osimgs) {
        [PSCustomObject]@{
            ContentName   = $osi.Name
            ContentType   = 'OS Image'
            PackageID     = $osi.PackageID
            SourceSize    = $osi.PackageSize
            SourceVersion = $osi.SourceDate
            ObjectID      = $osi.PackageID
        }
    }

    Write-Log "  $($results.Count) OS images"
    return $results
}

function Get-ContentDriverPackages {
    <#
    .SYNOPSIS
        Returns all driver packages with content properties.
    #>
    Write-Log "Loading driver packages..."
    $dpkgs = Get-CMDriverPackage -ErrorAction SilentlyContinue

    $results = foreach ($dpkg in $dpkgs) {
        [PSCustomObject]@{
            ContentName   = $dpkg.Name
            ContentType   = 'Driver Package'
            PackageID     = $dpkg.PackageID
            SourceSize    = $dpkg.PackageSize
            SourceVersion = $dpkg.SourceDate
            ObjectID      = $dpkg.PackageID
        }
    }

    Write-Log "  $($results.Count) driver packages"
    return $results
}

function Get-ContentTaskSequenceRefs {
    <#
    .SYNOPSIS
        Returns task sequence referenced content (cross-reference only).

    .DESCRIPTION
        Task sequences don't have content directly. This returns the TS name
        and the PackageIDs it references so they can be cross-referenced
        in the content view.
    #>
    Write-Log "Loading task sequence references..."
    $tslist = Get-CMTaskSequence -ErrorAction SilentlyContinue

    $results = foreach ($ts in $tslist) {
        $refIds = @()
        if ($ts.References) {
            $refIds = @($ts.References | ForEach-Object { $_.Package })
        }

        [PSCustomObject]@{
            ContentName   = $ts.Name
            ContentType   = 'Task Sequence'
            PackageID     = $ts.PackageID
            SourceSize    = $ts.PackageSize
            SourceVersion = $ts.SourceDate
            ObjectID      = $ts.PackageID
            ReferencedIDs = $refIds
        }
    }

    Write-Log "  $($results.Count) task sequences"
    return $results
}

function Get-AllContentObjects {
    <#
    .SYNOPSIS
        Retrieves all content objects across all types, merged into a unified array.

    .PARAMETER Types
        Optional string array to filter content types. If omitted, retrieves all types.
        Valid values: Application, Package, 'SU Deployment Pkg', 'Boot Image', 'OS Image',
                      'Driver Package', 'Task Sequence'
    #>
    param(
        [string[]]$Types
    )

    $all = @()
    $typeMap = [ordered]@{
        'Application'       = { Get-ContentApplications }
        'Package'           = { Get-ContentPackages }
        'SU Deployment Pkg' = { Get-ContentSUDPs }
        'Boot Image'        = { Get-ContentBootImages }
        'OS Image'          = { Get-ContentOSImages }
        'Driver Package'    = { Get-ContentDriverPackages }
        'Task Sequence'     = { Get-ContentTaskSequenceRefs }
    }

    foreach ($typeName in $typeMap.Keys) {
        if ($Types -and $Types -notcontains $typeName) { continue }
        $items = & $typeMap[$typeName]
        if ($items) { $all += $items }
    }

    Write-Log "Total content objects: $($all.Count)"
    return $all
}

# ---------------------------------------------------------------------------
# Status Queries
# ---------------------------------------------------------------------------

function Get-BulkDistributionStatus {
    <#
    .SYNOPSIS
        Bulk WMI query for all content-to-DP status rows.

    .DESCRIPTION
        Uses SMS_PackageStatusDistPointsSummarizer for a single query that returns
        all content distribution status across all DPs. Much faster than iterating
        Get-CMDistributionStatus per content object.

    .PARAMETER SMSProvider
        The SMS Provider server hostname.

    .PARAMETER SiteCode
        The CM site code.
    #>
    param(
        [Parameter(Mandatory)][string]$SMSProvider,
        [Parameter(Mandatory)][string]$SiteCode
    )

    Write-Log "Running bulk distribution status query against $SMSProvider..."

    $query = "SELECT PackageID, ServerNALPath, State, SourceVersion, SummaryDate FROM SMS_PackageStatusDistPointsSummarizer"

    $raw = Get-CimInstance -ComputerName $SMSProvider `
        -Namespace "root\SMS\site_$SiteCode" `
        -Query $query -ErrorAction Stop

    $results = foreach ($row in $raw) {
        $dpName = ''
        if ($row.ServerNALPath -match '\\\\([^\\]+)\\') {
            $dpName = $Matches[1]
        }

        [PSCustomObject]@{
            PackageID     = $row.PackageID
            DPName        = $dpName
            ServerNALPath = $row.ServerNALPath
            State         = $row.State
            StatusName    = switch ($row.State) {
                0 { 'Installed' }
                1 { 'Install Pending' }
                2 { 'Install Retrying' }
                3 { 'Install Failed' }
                4 { 'Removal Pending' }
                5 { 'Removal Retrying' }
                6 { 'Removal Failed' }
                7 { 'Content Validating' }
                8 { 'Content Valid' }
                default { "Unknown ($($row.State))" }
            }
            SourceVersion = $row.SourceVersion
            SummaryDate   = $row.SummaryDate
        }
    }

    Write-Log "Bulk status query returned $($results.Count) rows"
    return $results
}

function Get-ContentDistributionStatus {
    <#
    .SYNOPSIS
        Gets detailed distribution status for a single content object via CM cmdlet.

    .DESCRIPTION
        Used for drill-down detail, not bulk queries.
    #>
    param(
        [Parameter(Mandatory)]$ContentObject
    )

    return ($ContentObject | Get-CMDistributionStatus -ErrorAction SilentlyContinue)
}

function ConvertTo-ContentStatusSummary {
    <#
    .SYNOPSIS
        Aggregates bulk status rows into per-content-object summaries.

    .DESCRIPTION
        Uses hashtable-based O(n) aggregation for performance at scale.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][PSCustomObject[]]$StatusRows
    )

    $byPackage = @{}

    foreach ($row in $StatusRows) {
        $pkgId = $row.PackageID
        if (-not $byPackage.ContainsKey($pkgId)) {
            $byPackage[$pkgId] = @{ TotalDPs = 0; Installed = 0; InProgress = 0; Failed = 0 }
        }
        $byPackage[$pkgId].TotalDPs++

        switch ($row.State) {
            0       { $byPackage[$pkgId].Installed++ }
            8       { $byPackage[$pkgId].Installed++ }
            { $_ -in 1, 2, 7 } { $byPackage[$pkgId].InProgress++ }
            { $_ -in 3, 6 }    { $byPackage[$pkgId].Failed++ }
            { $_ -in 4, 5 }    { $byPackage[$pkgId].InProgress++ }
        }
    }

    $results = foreach ($pkgId in $byPackage.Keys) {
        $s = $byPackage[$pkgId]
        $pct = if ($s.TotalDPs -gt 0) { [math]::Round(($s.Installed / $s.TotalDPs) * 100, 1) } else { 0 }

        [PSCustomObject]@{
            PackageID      = $pkgId
            TotalDPs       = $s.TotalDPs
            InstalledCount = $s.Installed
            InProgressCount = $s.InProgress
            FailedCount    = $s.Failed
            PctComplete    = "$pct%"
        }
    }

    return $results
}

function ConvertTo-DPStatusSummary {
    <#
    .SYNOPSIS
        Aggregates bulk status rows into per-DP summaries.

    .DESCRIPTION
        Uses hashtable-based O(n) aggregation for performance at scale.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][PSCustomObject[]]$StatusRows
    )

    $byDP = @{}

    foreach ($row in $StatusRows) {
        $dp = $row.DPName
        if (-not $dp) { continue }
        if (-not $byDP.ContainsKey($dp)) {
            $byDP[$dp] = @{ TotalContent = 0; Installed = 0; InProgress = 0; Failed = 0 }
        }
        $byDP[$dp].TotalContent++

        switch ($row.State) {
            0       { $byDP[$dp].Installed++ }
            8       { $byDP[$dp].Installed++ }
            { $_ -in 1, 2, 7 } { $byDP[$dp].InProgress++ }
            { $_ -in 3, 6 }    { $byDP[$dp].Failed++ }
            { $_ -in 4, 5 }    { $byDP[$dp].InProgress++ }
        }
    }

    $results = foreach ($dp in $byDP.Keys) {
        $s = $byDP[$dp]
        $pct = if ($s.TotalContent -gt 0) { [math]::Round(($s.Installed / $s.TotalContent) * 100, 1) } else { 0 }

        [PSCustomObject]@{
            DPName          = $dp
            TotalContent    = $s.TotalContent
            InstalledCount  = $s.Installed
            InProgressCount = $s.InProgress
            FailedCount     = $s.Failed
            PctComplete     = "$pct%"
        }
    }

    return $results
}

# ---------------------------------------------------------------------------
# Actions
# ---------------------------------------------------------------------------

function Invoke-RedistributeContent {
    <#
    .SYNOPSIS
        Redistributes content to specified distribution points.

    .DESCRIPTION
        Sets RefreshNow on the SMS_DistributionPoint instance of each named DP.
        A DP name matches its server name exactly, or as the short name of a
        fully qualified server name. 'dp1' never matches 'dp10'. Returns one
        result per DP (PackageID, DPName, Success, Error); a failure on one DP
        does not stop the others. Under -WhatIf it changes nothing and returns
        no result.

    .PARAMETER PackageID
        The content package ID to redistribute.

    .PARAMETER DPName
        Array of DP server names to redistribute to.

    .PARAMETER ProgressCallback
        Optional scriptblock called with (currentIndex, totalCount, dpName) for progress reporting.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$PackageID,
        [Parameter(Mandatory)][string[]]$DPName,
        [scriptblock]$ProgressCallback
    )

    $results = @()
    $total = $DPName.Count

    $conn = Get-CMConnectionInfo
    if (-not $conn) { throw 'Not connected to a CM site; call Connect-CMSite first.' }

    for ($i = 0; $i -lt $total; $i++) {
        $dp = $DPName[$i]
        if ($ProgressCallback) { & $ProgressCallback $i $total $dp }

        try {
            # WMI RefreshNow triggers redistribution (same mechanism as the CM console)
            # Start-CMContentDistribution is for initial distribution only and fails on existing DPs
            $ns = "root\SMS\site_$($conn.SiteCode)"
            # WQL escapes a backslash and a quote with a backslash.
            $pkgLiteral = $PackageID.Replace('\', '\\').Replace("'", "\'")
            $wmiQuery = "SELECT * FROM SMS_DistributionPoint WHERE PackageID = '$pkgLiteral'"
            $candidates = @(Get-CimInstance -Namespace $ns -Query $wmiQuery -ComputerName $conn.SMSProvider -ErrorAction Stop)
            # Match the server name on the client. A LIKE '%dp1%' query also matches dp10 and dp11.
            $dpInst = @($candidates | Where-Object {
                $server = if ([string]$_.ServerNALPath -match '\\\\([^\\]+)\\') { $Matches[1] } else { '' }
                $server -ieq $dp -or $server.StartsWith("$dp.", [System.StringComparison]::OrdinalIgnoreCase)
            })
            if ($dpInst.Count -gt 0) {
                if ($PSCmdlet.ShouldProcess("$PackageID on $dp", 'Redistribute content')) {
                    foreach ($inst in $dpInst) {
                        Set-CimInstance -InputObject $inst -Property @{ RefreshNow = $true } -ErrorAction Stop
                    }
                    Write-Log "Redistributed $PackageID to $dp"
                    $results += [PSCustomObject]@{ PackageID = $PackageID; DPName = $dp; Success = $true; Error = '' }
                }
            } else {
                Write-Log "No SMS_DistributionPoint instance found for $PackageID on $dp" -Level WARN
                $results += [PSCustomObject]@{ PackageID = $PackageID; DPName = $dp; Success = $false; Error = 'DP instance not found' }
            }
        }
        catch {
            Write-Log "Failed to redistribute $PackageID to $dp : $_" -Level ERROR
            $results += [PSCustomObject]@{ PackageID = $PackageID; DPName = $dp; Success = $false; Error = $_.ToString() }
        }
    }

    return $results
}

function Remove-ContentFromDP {
    <#
    .SYNOPSIS
        Removes content from specified distribution points.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$PackageID,
        [Parameter(Mandatory)][string[]]$DPName,
        [scriptblock]$ProgressCallback
    )

    $results = @()
    $total = $DPName.Count

    for ($i = 0; $i -lt $total; $i++) {
        $dp = $DPName[$i]
        if ($ProgressCallback) { & $ProgressCallback $i $total $dp }

        try {
            if ($PSCmdlet.ShouldProcess("$PackageID on $dp", 'Remove content from DP')) {
            Remove-CMContentDistribution -PackageId $PackageID -DistributionPointName $dp -Force -ErrorAction Stop
            Write-Log "Removed $PackageID from $dp"
            $results += [PSCustomObject]@{ PackageID = $PackageID; DPName = $dp; Success = $true; Error = '' }
            }
        }
        catch {
            Write-Log "Failed to remove $PackageID from $dp : $_" -Level ERROR
            $results += [PSCustomObject]@{ PackageID = $PackageID; DPName = $dp; Success = $false; Error = $_.ToString() }
        }
    }

    return $results
}

function Invoke-ContentValidation {
    <#
    .SYNOPSIS
        Triggers content validation for specified content on specified DPs.

    .DESCRIPTION
        Content validation is server-side and asynchronous. Results will appear
        in the next status refresh.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$PackageID,
        [Parameter(Mandatory)][string[]]$DPName,
        [scriptblock]$ProgressCallback
    )

    $results = @()
    $total = $DPName.Count

    for ($i = 0; $i -lt $total; $i++) {
        $dp = $DPName[$i]
        if ($ProgressCallback) { & $ProgressCallback $i $total $dp }

        try {
            if ($PSCmdlet.ShouldProcess("$PackageID on $dp", 'Validate content on DP')) {
            Invoke-CMContentValidation -PackageId $PackageID -DistributionPointName $dp -ErrorAction Stop
            Write-Log "Validation initiated for $PackageID on $dp"
            $results += [PSCustomObject]@{ PackageID = $PackageID; DPName = $dp; Success = $true; Error = '' }
            }
        }
        catch {
            Write-Log "Failed to validate $PackageID on $dp : $_" -Level ERROR
            $results += [PSCustomObject]@{ PackageID = $PackageID; DPName = $dp; Success = $false; Error = $_.ToString() }
        }
    }

    return $results
}

function Assert-ContentActionResult {
    <#
    .SYNOPSIS
        Throws when any result of a content action reports Success = $false.

    .DESCRIPTION
        Invoke-RedistributeContent, Remove-ContentFromDP and Invoke-ContentValidation
        report a failure in the result object and do not throw. A caller that treats
        "did not throw" as success (the bulk operation dialog) passes the result
        through this function. The message lists each failed DP and its error.
        A null or empty result passes: a -WhatIf run returns no result.
    #>
    param(
        [Parameter(ValueFromPipeline)][AllowNull()][AllowEmptyCollection()]$Result
    )
    begin { $failed = @() }
    process {
        foreach ($r in @($Result)) {
            if ($r -and $r.PSObject.Properties['Success'] -and -not $r.Success) { $failed += $r }
        }
    }
    end {
        if ($failed.Count -gt 0) {
            throw (($failed | ForEach-Object { '{0}: {1}' -f $_.DPName, $_.Error }) -join '; ')
        }
    }
}

function Find-OrphanedContent {
    <#
    .SYNOPSIS
        Identifies content on DPs that no longer has a matching content object in ConfigMgr.

    .DESCRIPTION
        Compares PackageIDs in status rows against known content objects.
        Content present on DPs but absent from the content object list is orphaned.
    #>
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [PSCustomObject[]]$StatusRows,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [PSCustomObject[]]$ContentObjects
    )

    $knownIDs = @{}
    foreach ($obj in $ContentObjects) {
        $knownIDs[$obj.PackageID] = $true
    }

    $orphanedRows = $StatusRows | Where-Object { -not $knownIDs.ContainsKey($_.PackageID) }

    $byPkg = @{}
    foreach ($row in $orphanedRows) {
        if (-not $byPkg.ContainsKey($row.PackageID)) {
            $byPkg[$row.PackageID] = @{ DPNames = @(); States = @() }
        }
        $byPkg[$row.PackageID].DPNames += $row.DPName
        $byPkg[$row.PackageID].States += $row.StatusName
    }

    $results = foreach ($pkgId in $byPkg.Keys) {
        [PSCustomObject]@{
            PackageID = $pkgId
            DPCount   = $byPkg[$pkgId].DPNames.Count
            DPNames   = $byPkg[$pkgId].DPNames
        }
    }

    Write-Log "Found $($results.Count) orphaned content items"
    return $results
}

function Get-DPStorageAnalysis {
    <#
    .SYNOPSIS
        Calculates estimated content storage per DP.

    .DESCRIPTION
        Cross-references status rows with content objects to sum SourceSize per DP.
        Returns results sorted by total size descending.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][PSCustomObject[]]$StatusRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][PSCustomObject[]]$ContentObjects
    )

    $sizeLookup = @{}
    foreach ($obj in $ContentObjects) {
        $sizeLookup[$obj.PackageID] = $obj.SourceSize
    }

    $byDP = @{}
    foreach ($row in $StatusRows) {
        $dp = $row.DPName
        if (-not $dp) { continue }
        if (-not $byDP.ContainsKey($dp)) {
            $byDP[$dp] = @{ TotalSizeKB = 0; ContentCount = 0; FailedCount = 0 }
        }
        $byDP[$dp].ContentCount++
        if ($row.State -in 3, 6) { $byDP[$dp].FailedCount++ }

        $sizeKB = $sizeLookup[$row.PackageID]
        if ($sizeKB) { $byDP[$dp].TotalSizeKB += $sizeKB }
    }

    $results = foreach ($dp in $byDP.Keys) {
        $s = $byDP[$dp]
        [PSCustomObject]@{
            DPName       = $dp
            TotalSizeGB  = [math]::Round($s.TotalSizeKB / 1024 / 1024, 2)
            ContentCount = $s.ContentCount
            FailedCount  = $s.FailedCount
        }
    }

    return ($results | Sort-Object -Property TotalSizeGB -Descending)
}

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

function Export-ContentStatusCsv {
    <#
    .SYNOPSIS
        Exports a DataTable to CSV.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $rows = @()
    foreach ($row in $DataTable.Rows) {
        $obj = [ordered]@{}
        foreach ($col in $DataTable.Columns) {
            $obj[$col.ColumnName] = $row[$col.ColumnName]
        }
        $rows += [PSCustomObject]$obj
    }

    $rows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Log "Exported CSV to $OutputPath"
}

function Export-ContentStatusHtml {
    <#
    .SYNOPSIS
        Exports a DataTable to a self-contained HTML report with color-coded status.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$ReportTitle = 'DP Content Status Report'
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $css = @(
        '<style>',
        'body { font-family: "Segoe UI", Arial, sans-serif; margin: 20px; background: #fafafa; }',
        'h1 { color: #0078D4; margin-bottom: 4px; }',
        '.summary { color: #555; margin-bottom: 16px; }',
        'table { border-collapse: collapse; width: 100%; margin-top: 12px; }',
        'th { background: #0078D4; color: #fff; padding: 8px 12px; text-align: left; }',
        'td { padding: 6px 12px; border-bottom: 1px solid #e0e0e0; }',
        'tr:nth-child(even) { background: #f5f5f5; }',
        '.failed { color: #c00; font-weight: bold; }',
        '.inprog { color: #b87800; }',
        '.ok { color: #228b22; }',
        '</style>'
    ) -join "`r`n"

    # HtmlEncode every site-derived string: names with angle brackets or
    # ampersands otherwise corrupt the report markup.
    $headerRow = ($DataTable.Columns | ForEach-Object { "<th>$([System.Net.WebUtility]::HtmlEncode($_.ColumnName))</th>" }) -join ''
    $bodyRows = foreach ($row in $DataTable.Rows) {
        $cells = foreach ($col in $DataTable.Columns) {
            $val = [string]$row[$col.ColumnName]
            $cssClass = ''
            if ($col.ColumnName -match 'Failed' -and $val -match '^\d+$' -and [int]$val -gt 0) {
                $cssClass = ' class="failed"'
            }
            elseif ($col.ColumnName -match 'InProgress' -and $val -match '^\d+$' -and [int]$val -gt 0) {
                $cssClass = ' class="inprog"'
            }
            "<td$cssClass>$([System.Net.WebUtility]::HtmlEncode($val))</td>"
        }
        "<tr>$($cells -join '')</tr>"
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $html = @(
        '<!DOCTYPE html>',
        '<html><head><meta charset="utf-8"><title>' + [System.Net.WebUtility]::HtmlEncode($ReportTitle) + '</title>',
        $css,
        '</head><body>',
        "<h1>$([System.Net.WebUtility]::HtmlEncode($ReportTitle))</h1>",
        "<div class=`"summary`">Generated: $timestamp | Rows: $($DataTable.Rows.Count)</div>",
        "<table><thead><tr>$headerRow</tr></thead>",
        "<tbody>$($bodyRows -join "`r`n")</tbody></table>",
        '</body></html>'
    ) -join "`r`n"

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
    Write-Log "Exported HTML to $OutputPath"
}

function New-ContentStatusSummary {
    <#
    .SYNOPSIS
        Generates a plain-text summary suitable for clipboard or email.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable
    )

    $totalRows = $DataTable.Rows.Count
    $failedRows = @($DataTable.Select("FailedCount > 0")).Count

    $lines = @()
    $lines += "=== DP Content Status Summary ==="
    $lines += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += ""
    $lines += "Total content objects: $totalRows"
    $lines += "Objects with failures: $failedRows"
    $lines += ""

    if ($failedRows -gt 0) {
        $lines += "--- Failed Content ---"
        $failed = $DataTable.Select("FailedCount > 0")
        foreach ($row in $failed | Select-Object -First 20) {
            $lines += "  $($row['ContentName']) ($($row['PackageID'])) - $($row['FailedCount']) DPs failed"
        }
        if ($failedRows -gt 20) {
            $lines += "  ... and $($failedRows - 20) more"
        }
    }

    return ($lines -join "`r`n")
}

# ---------------------------------------------------------------------------
# Row shaping for the main window
# ---------------------------------------------------------------------------

function Get-RollupGlyph {
    <#
    .SYNOPSIS
        Returns the status glyph of a rollup: failed, in progress or complete.
    #>
    param([int]$Failed, [int]$InProgress, [int]$Total)
    if ($Total -eq 0)        { return [char]0x22EF }
    if ($Failed -gt 0)       { return [char]0x2717 }
    if ($InProgress -gt 0)   { return [char]0x22EF }
    return [char]0x2713
}

function Get-StatusRowGlyph {
    <#
    .SYNOPSIS
        Returns the status glyph of one distribution status state.
    #>
    param([int]$State)
    switch ($State) {
        0 { return [char]0x2713 }
        8 { return [char]0x2713 }
        3 { return [char]0x2717 }
        6 { return [char]0x2717 }
        default { return [char]0x22EF }
    }
}

function Format-ContentSize {
    <#
    .SYNOPSIS
        Formats a size in KB as KB, MB or GB text.
    #>
    param([double]$SizeKB)
    if ($SizeKB -ge 1048576) { return ('{0:N2} GB' -f ($SizeKB / 1048576)) }
    if ($SizeKB -ge 1024)    { return ('{0:N1} MB' -f ($SizeKB / 1024)) }
    return ('{0} KB' -f [long]$SizeKB)
}

function ConvertTo-DPRows {
    <#
    .SYNOPSIS
        Joins the DP list with the per-DP status and storage rollups into grid rows.
    #>
    param(
        [AllowNull()][AllowEmptyCollection()]$DPs,
        [AllowNull()][AllowEmptyCollection()]$ByDP,
        [AllowNull()][AllowEmptyCollection()]$StorageByDP
    )
    $byDPLookup = @{}
    foreach ($d in @($ByDP)) { if ($d) { $byDPLookup[[string]$d.DPName] = $d } }
    $storLookup = @{}
    foreach ($s in @($StorageByDP)) { if ($s) { $storLookup[[string]$s.DPName] = $s } }

    @(foreach ($dp in @($DPs)) {
        if (-not $dp) { continue }
        $name = [string]$dp.ServerName
        $stat = $byDPLookup[$name]
        $stor = $storLookup[$name]
        $total      = if ($stat) { [int]$stat.TotalContent }    else { 0 }
        $installed  = if ($stat) { [int]$stat.InstalledCount }  else { 0 }
        $inProgress = if ($stat) { [int]$stat.InProgressCount } else { 0 }
        $failed     = if ($stat) { [int]$stat.FailedCount }     else { 0 }
        $pct        = if ($stat) { [string]$stat.PctComplete }  else { '0%' }
        $sizeGB     = if ($stor) { [decimal]$stor.TotalSizeGB } else { 0 }
        [PSCustomObject]@{
            StatusGlyph     = Get-RollupGlyph -Failed $failed -InProgress $inProgress -Total $total
            ServerName      = $name
            SiteCode        = $dp.SiteCode
            IsPullDP        = $dp.IsPullDP
            TotalContent    = $total
            InstalledCount  = $installed
            InProgressCount = $inProgress
            FailedCount     = $failed
            PctComplete     = $pct
            TotalSizeGB     = $sizeGB
        }
    })
}

function ConvertTo-ContentRows {
    <#
    .SYNOPSIS
        Joins the content list with the per-content status rollup into grid rows.
    #>
    param(
        [AllowNull()][AllowEmptyCollection()]$Content,
        [AllowNull()][AllowEmptyCollection()]$ByContent
    )
    $byContentLookup = @{}
    foreach ($c in @($ByContent)) { if ($c) { $byContentLookup[[string]$c.PackageID] = $c } }

    @(foreach ($item in @($Content)) {
        if (-not $item) { continue }
        $packageId = [string]$item.PackageID
        $stat = $byContentLookup[$packageId]
        $total      = if ($stat) { [int]$stat.TotalDPs }        else { 0 }
        $installed  = if ($stat) { [int]$stat.InstalledCount }  else { 0 }
        $inProgress = if ($stat) { [int]$stat.InProgressCount } else { 0 }
        $failed     = if ($stat) { [int]$stat.FailedCount }     else { 0 }
        $pct        = if ($stat) { [string]$stat.PctComplete }  else { '0%' }
        $sizeKB = [long]$item.SourceSize
        [PSCustomObject]@{
            StatusGlyph     = Get-RollupGlyph -Failed $failed -InProgress $inProgress -Total $total
            Name            = $item.ContentName
            ContentType     = $item.ContentType
            PackageID       = $item.PackageID
            SizeDisplay     = Format-ContentSize -SizeKB $sizeKB
            SourceSize      = $sizeKB
            TotalDPs        = $total
            InstalledCount  = $installed
            InProgressCount = $inProgress
            FailedCount     = $failed
            PctComplete     = $pct
        }
    })
}

function ConvertTo-StatusIssueRows {
    <#
    .SYNOPSIS
        Returns one grid row for every status row that is not installed or valid.
    #>
    param(
        [AllowNull()][AllowEmptyCollection()]$Status,
        [AllowNull()][AllowEmptyCollection()]$Content
    )
    $contentLookup = @{}
    foreach ($c in @($Content)) { if ($c) { $contentLookup[[string]$c.PackageID] = $c } }

    @(foreach ($row in @($Status)) {
        if (-not $row) { continue }
        if ([int]$row.State -in 0, 8) { continue }
        $packageId = [string]$row.PackageID
        $info = $contentLookup[$packageId]
        [PSCustomObject]@{
            StatusGlyph = Get-StatusRowGlyph -State ([int]$row.State)
            DPName      = $row.DPName
            ContentName = if ($info) { $info.ContentName } else { '(unknown / orphaned)' }
            ContentType = if ($info) { $info.ContentType } else { '?' }
            State       = $row.State
            StatusName  = $row.StatusName
            PackageID   = $packageId
            LastUpdate  = $row.LastUpdate
        }
    })
}
