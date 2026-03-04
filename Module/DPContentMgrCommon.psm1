<#
.SYNOPSIS
    Core module for DP Content Manager.

.DESCRIPTION
    Import this module to get:
      - Structured logging (Initialize-Logging, Write-Log)
      - CM site connection management (Connect-CMSite, Disconnect-CMSite, Test-CMConnection)
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
# Module-scoped state
# ---------------------------------------------------------------------------

$script:__DPCMLogPath          = $null
$script:OriginalLocation       = $null
$script:ConnectedSiteCode      = $null
$script:ConnectedSMSProvider   = $null

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

function Initialize-Logging {
    param([string]$LogPath)

    $script:__DPCMLogPath = $LogPath

    if ($LogPath) {
        $parentDir = Split-Path -Path $LogPath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }

        $header = "[{0}] [INFO ] === Log initialized ===" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Set-Content -LiteralPath $LogPath -Value $header -Encoding UTF8
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped, severity-tagged log message.

    .DESCRIPTION
        INFO  -> Write-Host (stdout)
        WARN  -> Write-Host (stdout)
        ERROR -> Write-Host (stdout) + $host.UI.WriteErrorLine (stderr)

        -Quiet suppresses all console output but still writes to the log file.
    #>
    param(
        [AllowEmptyString()]
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO',

        [switch]$Quiet
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $formatted = "[{0}] [{1,-5}] {2}" -f $timestamp, $Level, $Message

    if (-not $Quiet) {
        Write-Host $formatted

        if ($Level -eq 'ERROR') {
            $host.UI.WriteErrorLine($formatted)
        }
    }

    if ($script:__DPCMLogPath) {
        Add-Content -LiteralPath $script:__DPCMLogPath -Value $formatted -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# CM Connection
# ---------------------------------------------------------------------------

function Connect-CMSite {
    <#
    .SYNOPSIS
        Imports the ConfigurationManager module, creates a PSDrive, and sets location.

    .DESCRIPTION
        Saves original location for restoration via Disconnect-CMSite.
        Returns $true on success, $false on failure.
    #>
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$SMSProvider
    )

    $script:OriginalLocation = Get-Location

    # Import CM module if not already loaded
    if (-not (Get-Module ConfigurationManager -ErrorAction SilentlyContinue)) {
        $cmModulePath = $null
        if ($env:SMS_ADMIN_UI_PATH) {
            $cmModulePath = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
        }

        if (-not $cmModulePath -or -not (Test-Path -LiteralPath $cmModulePath)) {
            Write-Log "ConfigurationManager module not found. Ensure the CM console is installed." -Level ERROR
            return $false
        }

        try {
            Import-Module $cmModulePath -ErrorAction Stop
            Write-Log "Imported ConfigurationManager module"
        }
        catch {
            Write-Log "Failed to import ConfigurationManager module: $_" -Level ERROR
            return $false
        }
    }

    # Create PSDrive if needed
    if (-not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        try {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SMSProvider -ErrorAction Stop | Out-Null
            Write-Log "Created PSDrive for site $SiteCode"
        }
        catch {
            Write-Log "Failed to create PSDrive for site $SiteCode : $_" -Level ERROR
            return $false
        }
    }

    try {
        Set-Location "${SiteCode}:" -ErrorAction Stop
        $site = Get-CMSite -SiteCode $SiteCode -ErrorAction Stop
        Write-Log "Connected to site $SiteCode ($($site.SiteName))"
        $script:ConnectedSiteCode    = $SiteCode
        $script:ConnectedSMSProvider = $SMSProvider
        return $true
    }
    catch {
        Write-Log "Failed to connect to site $SiteCode : $_" -Level ERROR
        return $false
    }
}

function Disconnect-CMSite {
    <#
    .SYNOPSIS
        Restores the original location before CM connection.
    #>
    if ($script:OriginalLocation) {
        try { Set-Location $script:OriginalLocation -ErrorAction SilentlyContinue } catch { }
    }
    $script:ConnectedSiteCode    = $null
    $script:ConnectedSMSProvider = $null
    Write-Log "Disconnected from CM site"
}

function Test-CMConnection {
    <#
    .SYNOPSIS
        Returns $true if currently connected to a CM site.
    #>
    if (-not $script:ConnectedSiteCode) { return $false }

    try {
        $drive = Get-PSDrive -Name $script:ConnectedSiteCode -PSProvider CMSite -ErrorAction Stop
        return ($null -ne $drive)
    }
    catch {
        return $false
    }
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
        [Parameter(Mandatory)][PSCustomObject[]]$StatusRows
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
        [Parameter(Mandatory)][PSCustomObject[]]$StatusRows
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

    .PARAMETER PackageID
        The content package ID to redistribute.

    .PARAMETER DPName
        Array of DP server names to redistribute to.

    .PARAMETER ProgressCallback
        Optional scriptblock called with (currentIndex, totalCount, dpName) for progress reporting.
    #>
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
            # WMI RefreshNow triggers redistribution (same mechanism as the CM console)
            # Start-CMContentDistribution is for initial distribution only and fails on existing DPs
            $ns = "root\SMS\site_$($script:ConnectedSiteCode)"
            $wmiQuery = "SELECT * FROM SMS_DistributionPoint WHERE PackageID = '$PackageID' AND ServerNALPath LIKE '%$dp%'"
            $dpInst = Get-CimInstance -Namespace $ns -Query $wmiQuery -ComputerName $script:ConnectedSMSProvider -ErrorAction Stop
            if ($dpInst) {
                $dpInst | Set-CimInstance -Property @{ RefreshNow = $true } -ErrorAction Stop
                Write-Log "Redistributed $PackageID to $dp"
                $results += [PSCustomObject]@{ PackageID = $PackageID; DPName = $dp; Success = $true; Error = '' }
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
            Remove-CMContentDistribution -PackageId $PackageID -DistributionPointName $dp -Force -ErrorAction Stop
            Write-Log "Removed $PackageID from $dp"
            $results += [PSCustomObject]@{ PackageID = $PackageID; DPName = $dp; Success = $true; Error = '' }
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
            Invoke-CMContentValidation -Id $PackageID -DistributionPointName $dp -ErrorAction Stop
            Write-Log "Validation initiated for $PackageID on $dp"
            $results += [PSCustomObject]@{ PackageID = $PackageID; DPName = $dp; Success = $true; Error = '' }
        }
        catch {
            Write-Log "Failed to validate $PackageID on $dp : $_" -Level ERROR
            $results += [PSCustomObject]@{ PackageID = $PackageID; DPName = $dp; Success = $false; Error = $_.ToString() }
        }
    }

    return $results
}

function Find-OrphanedContent {
    <#
    .SYNOPSIS
        Identifies content on DPs that no longer has a matching content object in MECM.

    .DESCRIPTION
        Compares PackageIDs in status rows against known content objects.
        Content present on DPs but absent from the content object list is orphaned.
    #>
    param(
        [Parameter(Mandatory)][PSCustomObject[]]$StatusRows,
        [Parameter(Mandatory)][PSCustomObject[]]$ContentObjects
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
        [Parameter(Mandatory)][PSCustomObject[]]$StatusRows,
        [Parameter(Mandatory)][PSCustomObject[]]$ContentObjects
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

    $headerRow = ($DataTable.Columns | ForEach-Object { "<th>$($_.ColumnName)</th>" }) -join ''
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
            "<td$cssClass>$val</td>"
        }
        "<tr>$($cells -join '')</tr>"
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $html = @(
        '<!DOCTYPE html>',
        '<html><head><meta charset="utf-8"><title>' + $ReportTitle + '</title>',
        $css,
        '</head><body>',
        "<h1>$ReportTitle</h1>",
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
