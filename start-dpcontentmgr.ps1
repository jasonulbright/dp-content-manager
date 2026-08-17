<#
.SYNOPSIS
    MahApps.Metro WPF shell for the MECM Distribution Point Content Manager.

.DESCRIPTION
    Sidebar navigation across three views (DPs, Content, Status Issues),
    inline action bar (Refresh, filter, status filter, exports), modal
    dialogs for Options, Redistribute, Remove from DP, Validate. Status
    Issues view is a triage list of every DP x Content row currently in
    a non-OK state.

    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.7.2+
      - MahApps.Metro DLLs in .\Lib\
      - DPContentMgrCommon module under .\Module\
      - ConfigurationManager console (provides Get-CMDistributionPoint, etc.)

.NOTES
    ScriptName : start-dpcontentmgr.ps1
    Version    : 1.1.0
    Updated    : 2026-05-02
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification='Per feedback_ps_wpf_handler_rules.md and PS51-WPF-001..003: $global: survives closure scope-strip.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification='WPF event handler scriptblocks bind positional sender/args ($s, $e).')]
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$__txDir = Join-Path $PSScriptRoot 'Logs'
try {
    if (-not (Test-Path -LiteralPath $__txDir)) { New-Item -ItemType Directory -Path $__txDir -Force | Out-Null }
    $__tx = Join-Path $__txDir ('DPCM-startup-{0}.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    Start-Transcript -LiteralPath $__tx -Force | Out-Null
} catch { $null = $_ }

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $psExe = (Get-Process -Id $PID).Path
    $fwd   = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',$PSCommandPath)
    Start-Process -FilePath $psExe -ArgumentList $fwd | Out-Null
    try { Stop-Transcript | Out-Null } catch { $null = $_ }
    exit 0
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$libDir = Join-Path $PSScriptRoot 'Lib'
if (-not (Test-Path -LiteralPath $libDir)) { throw "Lib/ directory not found at: $libDir." }
Get-ChildItem -LiteralPath $libDir -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'Microsoft.Xaml.Behaviors.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'ControlzEx.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'MahApps.Metro.dll'))

$__modulePath = Join-Path $PSScriptRoot 'Module\DPContentMgrCommon.psd1'
if (-not (Test-Path -LiteralPath $__modulePath)) { throw "Shared module not found at: $__modulePath" }
Import-Module -Name $__modulePath -Force -DisableNameChecking

$global:PrefsPath = Join-Path $PSScriptRoot 'DPContentMgr.prefs.json'
function Get-DpcmPreferences {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns the full preferences hashtable by design.')]
    param()
    return Read-SuiteSettings -Path $global:PrefsPath -Defaults @{ DarkMode = $true; SiteCode = ''; SMSProvider = '' }
}
function Save-DpcmPreferences {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Writes the full preferences hashtable by design.')]
    param([Parameter(Mandatory)][hashtable]$Prefs)
    $null = Save-SuiteSettings -Path $global:PrefsPath -Settings $Prefs
}
$global:Prefs = Get-DpcmPreferences

$script:ToolLogPath = Join-Path $__txDir ('DPCM-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $script:ToolLogPath

$xamlPath = Join-Path $PSScriptRoot 'MainWindow.xaml'
[xml]$xaml = Get-Content -LiteralPath $xamlPath -Raw
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$txtAppTitle        = $window.FindName('txtAppTitle')
$txtVersion         = $window.FindName('txtVersion')
$txtThemeLabel      = $window.FindName('txtThemeLabel')
$toggleTheme        = $window.FindName('toggleTheme')

$btnViewDPs          = $window.FindName('btnViewDPs')
$btnViewContent      = $window.FindName('btnViewContent')
$btnViewStatusIssues = $window.FindName('btnViewStatusIssues')
$btnOptions          = $window.FindName('btnOptions')

$txtModuleTitle    = $window.FindName('txtModuleTitle')
$txtModuleSubtitle = $window.FindName('txtModuleSubtitle')

$btnRefresh      = $window.FindName('btnRefresh')
$txtFilter       = $window.FindName('txtFilter')
$cboStatusFilter = $window.FindName('cboStatusFilter')
$btnExportCsv    = $window.FindName('btnExportCsv')
$btnExportHtml   = $window.FindName('btnExportHtml')

$viewDPs          = $window.FindName('viewDPs')
$viewContent      = $window.FindName('viewContent')
$viewStatusIssues = $window.FindName('viewStatusIssues')

$gridDPs          = $window.FindName('gridDPs')
$txtDPDetail      = $window.FindName('txtDPDetail')

$gridContent       = $window.FindName('gridContent')
$gridContentPerDP  = $window.FindName('gridContentPerDP')
$btnRedistribute   = $window.FindName('btnRedistribute')
$btnRemoveFromDP   = $window.FindName('btnRemoveFromDP')
$btnValidate       = $window.FindName('btnValidate')

$gridStatusIssues       = $window.FindName('gridStatusIssues')
$btnRedistributeIssues  = $window.FindName('btnRedistributeIssues')
$btnRemoveIssues        = $window.FindName('btnRemoveIssues')
$txtIssueCount          = $window.FindName('txtIssueCount')

$progressOverlay  = $window.FindName('progressOverlay')
$txtProgressTitle = $window.FindName('txtProgressTitle')
$txtProgressStep  = $window.FindName('txtProgressStep')

$lblLogOutput = $window.FindName('lblLogOutput')
$txtLog       = $window.FindName('txtLog')
$txtStatus    = $window.FindName('txtStatus')

$null = $txtAppTitle, $txtVersion

function Add-LogLine {
    param([Parameter(Mandatory)][string]$Message)
    $ts = (Get-Date).ToString('HH:mm:ss')
    $line = '{0}  {1}' -f $ts, $Message
    if ([string]::IsNullOrWhiteSpace($txtLog.Text)) { $txtLog.Text = $line }
    else { $txtLog.AppendText([Environment]::NewLine + $line) }
    $txtLog.ScrollToEnd()
}

function Set-StatusText {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates an in-window TextBlock only.')]
    param([Parameter(Mandatory)][string]$Text)
    $txtStatus.Text = $Text
}

# === Title-bar drag fallback (PS51-WPF-033): SuiteCommon owns the hook ===
Install-TitleBarDragFallback -Window $window

# === Theme (palette, sidebar, title bar, dialogs: SuiteCommon) ===
[void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Dark.Steel')

$script:ViewButtons = @(
    @{ Name = 'DPs';            Button = $btnViewDPs }
    @{ Name = 'Content';        Button = $btnViewContent }
    @{ Name = 'Status Issues';  Button = $btnViewStatusIssues }
)
$script:ActiveView = 'DPs'

Initialize-SuiteTheme -Window $window `
    -IsDarkGetter { [bool]$global:Prefs['DarkMode'] } `
    -ActiveViewGetter { $script:ActiveView } `
    -ViewButtons $script:ViewButtons `
    -OptionsButton $btnOptions `
    -LogLabel $lblLogOutput

$__startIsDark = [bool]$global:Prefs['DarkMode']
$toggleTheme.IsOn = $__startIsDark
$txtThemeLabel.Text = if ($__startIsDark) { 'Dark Theme' } else { 'Light Theme' }
Update-SidebarButtonTheme

$toggleTheme.Add_Toggled({
    $isDark = [bool]$toggleTheme.IsOn
    if ($isDark) { [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Dark.Steel'); $txtThemeLabel.Text = 'Dark Theme' }
    else         { [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Light.Blue'); $txtThemeLabel.Text = 'Light Theme' }
    $global:Prefs['DarkMode'] = $isDark
    Save-DpcmPreferences -Prefs $global:Prefs
    Update-SidebarButtonTheme
    Update-TitleBarBrushes
    Add-LogLine ('Theme: {0}' -f $(if ($isDark) { 'dark' } else { 'light' }))
})

# === View switching ===
$script:ViewMeta = @{
    'DPs'           = @{ Title = 'DPs';            Subtitle = 'Per-DP rollup of total content, installed, in progress, failed.' }
    'Content'       = @{ Title = 'Content';        Subtitle = 'Per-content-object rollup. Select a row to see its per-DP distribution status.' }
    'Status Issues' = @{ Title = 'Status Issues';  Subtitle = 'Every (DP, Content) row currently in a non-OK state. The triage view -- multi-select for bulk redistribute / remove.' }
}
function Set-ActiveView {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='In-window Visibility + header text only.')]
    param([Parameter(Mandatory)][ValidateSet('DPs','Content','Status Issues')][string]$View)
    $script:ActiveView = $View
    $viewDPs.Visibility          = if ($View -eq 'DPs')           { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $viewContent.Visibility      = if ($View -eq 'Content')       { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $viewStatusIssues.Visibility = if ($View -eq 'Status Issues') { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $meta = $script:ViewMeta[$View]
    if ($meta) {
        $txtModuleTitle.Text    = $meta.Title
        $txtModuleSubtitle.Text = $meta.Subtitle
    }
    Update-SidebarButtonTheme
    Update-ActionBarVisibility
    Update-Filter
    Update-StatusBarSummary
}
$btnViewDPs.Add_Click({          Set-ActiveView -View 'DPs'           })
$btnViewContent.Add_Click({      Set-ActiveView -View 'Content'       })
$btnViewStatusIssues.Add_Click({ Set-ActiveView -View 'Status Issues' })

# === Crash handlers ===
$global:__crashLog = Join-Path $__txDir ('DPCM-crash-{0}.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
$global:__writeCrash = {
    param($Source, $Exception)
    try {
        $lines = @(('=== ' + $Source + ' @ ' + (Get-Date -Format 'o') + ' ==='))
        $lines += ('Type   : ' + $Exception.GetType().FullName)
        $lines += ('Message: ' + $Exception.Message)
        $lines += ([string]$Exception.StackTrace).Split([Environment]::NewLine)
        [System.IO.File]::AppendAllText($global:__crashLog, (($lines -join [Environment]::NewLine) + [Environment]::NewLine))
    } catch { $null = $_ }
}
$window.Dispatcher.Add_UnhandledException({ param($s, $e) & $global:__writeCrash 'DispatcherUnhandledException' $e.Exception; $e.Handled = $false })
[AppDomain]::CurrentDomain.Add_UnhandledException({ param($s, $e) & $global:__writeCrash 'AppDomainUnhandledException' ([Exception]$e.ExceptionObject) })

# === State ===
$script:RawDPs            = @()
$script:RawContent        = @()
$script:RawStatus         = @()
$script:DPRows            = @()
$script:ContentRows       = @()
$script:StatusIssueRows   = @()
$script:LastRefreshTime   = $null
$script:IsConnectedFromBg = $false

# === Glyphs ===
function Get-RollupGlyph {
    param([int]$Failed, [int]$InProgress, [int]$Total)
    if ($Total -eq 0)        { return [char]0x22EF }
    if ($Failed -gt 0)       { return [char]0x2717 }
    if ($InProgress -gt 0)   { return [char]0x22EF }
    return [char]0x2713
}
function Get-StatusRowGlyph {
    param([int]$State)
    switch ($State) {
        0 { return [char]0x2713 }
        8 { return [char]0x2713 }
        3 { return [char]0x2717 }
        6 { return [char]0x2717 }
        default { return [char]0x22EF }
    }
}

# === Action bar visibility ===
function Update-ActionBarVisibility {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Toggles in-window Visibility.')]
    param()
    switch ($script:ActiveView) {
        'DPs' {
            $cboStatusFilter.Visibility = [System.Windows.Visibility]::Visible
            $txtFilter.Visibility       = [System.Windows.Visibility]::Visible
            $btnExportCsv.Visibility    = [System.Windows.Visibility]::Visible
            $btnExportHtml.Visibility   = [System.Windows.Visibility]::Visible
            $txtFilter.Tag              = 'Filter by DP name...'
            $cboStatusFilter.Items.Clear()
            foreach ($it in @('All','Healthy only','Has failures','Has in-progress')) {
                $cboItem = New-Object System.Windows.Controls.ComboBoxItem
                $cboItem.Content = $it
                if ($it -eq 'All') { $cboItem.IsSelected = $true }
                [void]$cboStatusFilter.Items.Add($cboItem)
            }
        }
        'Content' {
            $cboStatusFilter.Visibility = [System.Windows.Visibility]::Visible
            $txtFilter.Visibility       = [System.Windows.Visibility]::Visible
            $btnExportCsv.Visibility    = [System.Windows.Visibility]::Visible
            $btnExportHtml.Visibility   = [System.Windows.Visibility]::Visible
            $txtFilter.Tag              = 'Filter by content name or PackageID...'
            $cboStatusFilter.Items.Clear()
            foreach ($it in @('All','Application','Package','SUDP','BootImage','OSImage','DriverPackage','TaskSequence')) {
                $cboItem = New-Object System.Windows.Controls.ComboBoxItem
                $cboItem.Content = $it
                if ($it -eq 'All') { $cboItem.IsSelected = $true }
                [void]$cboStatusFilter.Items.Add($cboItem)
            }
        }
        'Status Issues' {
            $cboStatusFilter.Visibility = [System.Windows.Visibility]::Visible
            $txtFilter.Visibility       = [System.Windows.Visibility]::Visible
            $btnExportCsv.Visibility    = [System.Windows.Visibility]::Visible
            $btnExportHtml.Visibility   = [System.Windows.Visibility]::Visible
            $txtFilter.Tag              = 'Filter by DP or content name...'
            $cboStatusFilter.Items.Clear()
            foreach ($it in @('All','Failed only','In progress only')) {
                $cboItem = New-Object System.Windows.Controls.ComboBoxItem
                $cboItem.Content = $it
                if ($it -eq 'All') { $cboItem.IsSelected = $true }
                [void]$cboStatusFilter.Items.Add($cboItem)
            }
        }
    }
    [MahApps.Metro.Controls.TextBoxHelper]::SetWatermark($txtFilter, [string]$txtFilter.Tag)
}

function Update-StatusBarSummary {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='In-window TextBlock update.')]
    param()
    $parts = @()
    if ($script:IsConnectedFromBg -and $global:Prefs.SiteCode) { $parts += "Connected to $($global:Prefs.SiteCode)" }
    elseif (-not $global:Prefs.SiteCode -or -not $global:Prefs.SMSProvider) { $parts += 'Open Options to configure site code and SMS provider' }
    else { $parts += 'Ready. Click Refresh.' }
    if (@($script:RawDPs).Count -gt 0)          { $parts += ('{0} DPs' -f @($script:RawDPs).Count) }
    if (@($script:RawContent).Count -gt 0)      { $parts += ('{0} content' -f @($script:RawContent).Count) }
    if (@($script:StatusIssueRows).Count -gt 0) { $parts += ('{0} issues' -f @($script:StatusIssueRows).Count) }
    if ($script:LastRefreshTime) { $parts += ('last refresh {0}' -f $script:LastRefreshTime.ToString('HH:mm:ss')) }
    Set-StatusText ($parts -join '   |   ')
}

# === Filter ===
function Get-StatusFilterValue {
    if (-not $cboStatusFilter.SelectedItem) { return 'All' }
    $item = $cboStatusFilter.SelectedItem
    if ($item -is [System.Windows.Controls.ComboBoxItem]) { return [string]$item.Content }
    return [string]$item
}
function Update-Filter {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Recomputes ItemsSource on the active grid.')]
    param()
    $needle = ([string]$txtFilter.Text).Trim().ToLowerInvariant()
    $statusFilter = Get-StatusFilterValue
    switch ($script:ActiveView) {
        'DPs' {
            $rows = $script:DPRows
            if ($needle) { $rows = @($rows | Where-Object { ([string]$_.ServerName).ToLowerInvariant().Contains($needle) }) }
            switch ($statusFilter) {
                'Healthy only'    { $rows = @($rows | Where-Object { [int]$_.FailedCount -eq 0 -and [int]$_.InProgressCount -eq 0 }) }
                'Has failures'    { $rows = @($rows | Where-Object { [int]$_.FailedCount -gt 0 }) }
                'Has in-progress' { $rows = @($rows | Where-Object { [int]$_.InProgressCount -gt 0 }) }
            }
            $gridDPs.ItemsSource = $rows
        }
        'Content' {
            $rows = $script:ContentRows
            if ($needle) {
                $rows = @($rows | Where-Object {
                    ([string]$_.Name).ToLowerInvariant().Contains($needle) -or
                    ([string]$_.PackageID).ToLowerInvariant().Contains($needle)
                })
            }
            if ($statusFilter -ne 'All') {
                $rows = @($rows | Where-Object { ([string]$_.ContentType) -eq $statusFilter })
            }
            $gridContent.ItemsSource = $rows
        }
        'Status Issues' {
            $rows = $script:StatusIssueRows
            if ($needle) {
                $rows = @($rows | Where-Object {
                    ([string]$_.DPName).ToLowerInvariant().Contains($needle) -or
                    ([string]$_.ContentName).ToLowerInvariant().Contains($needle)
                })
            }
            switch ($statusFilter) {
                'Failed only'      { $rows = @($rows | Where-Object { [int]$_.State -in 3,6 }) }
                'In progress only' { $rows = @($rows | Where-Object { [int]$_.State -in 1,2,4,5,7 }) }
            }
            $gridStatusIssues.ItemsSource = $rows
            $txtIssueCount.Text = ('{0} issue rows' -f @($rows).Count)
        }
    }
}
$txtFilter.Add_TextChanged({ Update-Filter })
$cboStatusFilter.Add_SelectionChanged({ Update-Filter })

# === Detail panels ===
$gridDPs.Add_SelectionChanged({
    $row = $gridDPs.SelectedItem
    if (-not $row) { $txtDPDetail.Text = 'Select a DP to see content roll-up + storage analysis.'; return }
    $lines = @(
        ('DP:               {0}' -f $row.ServerName),
        ('Site:             {0}' -f $row.SiteCode),
        ('Pull DP:          {0}' -f $row.IsPullDP),
        ('Total Content:    {0}' -f $row.TotalContent),
        ('Installed:        {0}' -f $row.InstalledCount),
        ('In Progress:      {0}' -f $row.InProgressCount),
        ('Failed:           {0}' -f $row.FailedCount),
        ('Pct Complete:     {0}' -f $row.PctComplete)
    )
    if ($row.PSObject.Properties['TotalSizeGB']) {
        $lines += ('Total Storage:    {0} GB' -f $row.TotalSizeGB)
    }
    $txtDPDetail.Text = $lines -join [Environment]::NewLine
})

$gridContent.Add_SelectionChanged({
    $row = $gridContent.SelectedItem
    if (-not $row) { $gridContentPerDP.ItemsSource = $null; return }
    $perDp = @($script:RawStatus | Where-Object { $_.PackageID -eq $row.PackageID } | ForEach-Object {
        $glyph = Get-StatusRowGlyph -State ([int]$_.State)
        [PSCustomObject]@{
            StatusGlyph = $glyph
            DPName      = $_.DPName
            State       = $_.State
            StatusName  = $_.StatusName
            LastUpdate  = $_.LastUpdate
        }
    })
    $gridContentPerDP.ItemsSource = $perDp
})

# === Bg runspace + refresh ===
$script:BgRunspace     = $null
$script:BgPowerShell   = $null
$script:BgInvokeHandle = $null
$script:BgState        = $null
$script:BgTimer        = $null
$script:BgGraveyard    = @()

function Initialize-BgRunspace {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Lazy-init; idempotent.')]
    param()
    if ($script:BgRunspace -and $script:BgRunspace.RunspaceStateInfo.State -eq 'Opened') { return }
    $script:BgRunspace = New-SuiteBgRunspace -ModulePath (Join-Path $PSScriptRoot 'Module\DPContentMgrCommon.psd1') -LogPath $script:ToolLogPath
}

function Dispose-BgWork {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification='Dispose semantics intentional.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Tears down ephemeral runspace plumbing.')]
    param()
    $script:BgGraveyard = @(Stop-SuiteBgWork -PowerShell $script:BgPowerShell -Timer $script:BgTimer -Graveyard $script:BgGraveyard)
    $script:BgTimer = $null
    $script:BgPowerShell = $null
    $script:BgInvokeHandle = $null
}

function Invoke-Refresh {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Posts work to bg runspace.')]
    param()
    if (-not $global:Prefs.SiteCode -or -not $global:Prefs.SMSProvider) {
        Add-LogLine 'Refresh: site code and SMS provider must be set in Options first.'
        Set-StatusText 'Open Options to configure site code and SMS provider, then refresh.'
        return
    }
    Initialize-BgRunspace
    Dispose-BgWork
    $script:BgState = [hashtable]::Synchronized(@{ Step = 'Connecting...'; Done = $false; Result = $null; ErrorMsg = $null })
    $btnRefresh.IsEnabled = $false
    $txtProgressTitle.Text = 'Loading DP content status'
    $txtProgressStep.Text  = 'Connecting...'
    $progressOverlay.Visibility = [System.Windows.Visibility]::Visible
    Add-LogLine ('Refresh: site={0} provider={1}' -f $global:Prefs.SiteCode, $global:Prefs.SMSProvider)
    Set-StatusText 'Refreshing...'

    $siteCode    = [string]$global:Prefs.SiteCode
    $smsProvider = [string]$global:Prefs.SMSProvider

    $script:BgPowerShell = [powershell]::Create()
    $script:BgPowerShell.Runspace = $script:BgRunspace
    [void]$script:BgPowerShell.AddScript({
        param($SiteCode, $SMSProvider, $State)
        try {
            if (-not (Test-CMConnection)) {
                $State.Step = "Connecting to $SiteCode..."
                $ok = Connect-CMSite -SiteCode $SiteCode -SMSProvider $SMSProvider
                if (-not $ok) { $State.ErrorMsg = "Failed to connect to site $SiteCode (provider $SMSProvider)."; return }
            }
            $State.Step = 'Loading distribution points...'
            $dps = @(Get-AllDistributionPoints)

            $State.Step = 'Loading content objects (apps, packages, updates, images, drivers, TS refs)...'
            $content = @(Get-AllContentObjects)

            $State.Step = ('Querying bulk distribution status across {0} content x {1} DPs...' -f $content.Count, $dps.Count)
            $status = @(Get-BulkDistributionStatus -SMSProvider $SMSProvider -SiteCode $SiteCode)

            $State.Step = 'Aggregating per-content + per-DP rollups...'
            $byContent = @(ConvertTo-ContentStatusSummary -StatusRows $status)
            $byDP      = @(ConvertTo-DPStatusSummary      -StatusRows $status)
            $storage   = @(Get-DPStorageAnalysis          -StatusRows $status -ContentObjects $content)

            $State.Result = [PSCustomObject]@{
                DPs              = $dps
                Content          = $content
                Status           = $status
                ByContent        = $byContent
                ByDP             = $byDP
                StorageByDP      = $storage
            }
        }
        catch { $State.ErrorMsg = $_.Exception.Message }
        finally { $State.Done = $true }
    }).AddArgument($siteCode).AddArgument($smsProvider).AddArgument($script:BgState)

    $script:BgInvokeHandle = $script:BgPowerShell.BeginInvoke()
    $script:BgTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:BgTimer.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:BgTimer.Add_Tick({
        if ($script:BgState) { $current = [string]$script:BgState.Step; if ($txtProgressStep.Text -ne $current) { $txtProgressStep.Text = $current } }
        if ($script:BgState -and $script:BgState.Done) {
            $script:BgTimer.Stop()
            try { [void]$script:BgPowerShell.EndInvoke($script:BgInvokeHandle) } catch { $null = $_ }
            try { $script:BgPowerShell.Dispose() } catch { $null = $_ }
            $script:BgPowerShell   = $null
            $script:BgInvokeHandle = $null

            if ($script:BgState.ErrorMsg) {
                $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
                $btnRefresh.IsEnabled = $true
                $script:IsConnectedFromBg = $false
                Add-LogLine ('Refresh failed: {0}' -f $script:BgState.ErrorMsg)
                Set-StatusText 'Refresh failed.'
                return
            }

            $script:IsConnectedFromBg = $true
            $r = $script:BgState.Result
            $script:RawDPs     = @($r.DPs)
            $script:RawContent = @($r.Content)
            $script:RawStatus  = @($r.Status)
            $script:LastRefreshTime = Get-Date

            # DP rows: join byDP + raw DP info + storage
            $byDPLookup = @{}
            foreach ($d in @($r.ByDP)) { $byDPLookup[[string]$d.DPName] = $d }
            $storLookup = @{}
            foreach ($s in @($r.StorageByDP)) { $storLookup[[string]$s.DPName] = $s }

            $script:DPRows = @($script:RawDPs | ForEach-Object {
                $name = [string]$_.ServerName
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
                    SiteCode        = $_.SiteCode
                    IsPullDP        = $_.IsPullDP
                    TotalContent    = $total
                    InstalledCount  = $installed
                    InProgressCount = $inProgress
                    FailedCount     = $failed
                    PctComplete     = $pct
                    TotalSizeGB     = $sizeGB
                }
            })

            # Content rows: join byContent + raw content
            $byContentLookup = @{}
            foreach ($c in @($r.ByContent)) { $byContentLookup[[string]$c.PackageID] = $c }
            $script:ContentRows = @($script:RawContent | ForEach-Object {
                $packageId = [string]$_.PackageID
                $stat = $byContentLookup[$packageId]
                $total      = if ($stat) { [int]$stat.TotalDPs }        else { 0 }
                $installed  = if ($stat) { [int]$stat.InstalledCount }  else { 0 }
                $inProgress = if ($stat) { [int]$stat.InProgressCount } else { 0 }
                $failed     = if ($stat) { [int]$stat.FailedCount }     else { 0 }
                $pct        = if ($stat) { [string]$stat.PctComplete }  else { '0%' }
                $sizeKB = [int]$_.SourceSize
                $sizeStr = if ($sizeKB -ge 1048576) { ('{0:N2} GB' -f ($sizeKB / 1048576)) }
                          elseif ($sizeKB -ge 1024) { ('{0:N1} MB' -f ($sizeKB / 1024)) }
                          else { ('{0} KB' -f $sizeKB) }
                [PSCustomObject]@{
                    StatusGlyph     = Get-RollupGlyph -Failed $failed -InProgress $inProgress -Total $total
                    Name            = $_.ContentName
                    ContentType     = $_.ContentType
                    PackageID       = $_.PackageID
                    SizeDisplay     = $sizeStr
                    SourceSize      = $sizeKB
                    TotalDPs        = $total
                    InstalledCount  = $installed
                    InProgressCount = $inProgress
                    FailedCount     = $failed
                    PctComplete     = $pct
                }
            })

            # Status Issues: every status row in non-OK state
            $contentNameLookup = @{}
            foreach ($c in $script:RawContent) { $contentNameLookup[[string]$c.PackageID] = $c }
            $script:StatusIssueRows = @($script:RawStatus | Where-Object { [int]$_.State -notin 0,8 } | ForEach-Object {
                $packageId = [string]$_.PackageID
                $cInfo = $contentNameLookup[$packageId]
                [PSCustomObject]@{
                    StatusGlyph = Get-StatusRowGlyph -State ([int]$_.State)
                    DPName      = $_.DPName
                    ContentName = if ($cInfo) { $cInfo.ContentName } else { '(unknown / orphaned)' }
                    ContentType = if ($cInfo) { $cInfo.ContentType } else { '?' }
                    State       = $_.State
                    StatusName  = $_.StatusName
                    PackageID   = $packageId
                    LastUpdate  = $_.LastUpdate
                }
            })

            $gridDPs.ItemsSource = $script:DPRows
            $gridContent.ItemsSource = $script:ContentRows
            $gridStatusIssues.ItemsSource = $script:StatusIssueRows
            $txtIssueCount.Text = ('{0} issue rows' -f @($script:StatusIssueRows).Count)

            # UI thread CM connect mirror so per-action cmdlets resolve.
            try { [void](Connect-CMSite -SiteCode $global:Prefs.SiteCode -SMSProvider $global:Prefs.SMSProvider) }
            catch { Add-LogLine ('UI-thread CM connect: {0}' -f $_.Exception.Message) }

            Update-Filter
            Update-StatusBarSummary
            $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
            $btnRefresh.IsEnabled = $true
            Add-LogLine ('Refresh complete: {0} DPs, {1} content, {2} issues.' -f @($script:RawDPs).Count, @($script:RawContent).Count, @($script:StatusIssueRows).Count)
        }
    })
    $script:BgTimer.Start()
}
$btnRefresh.Add_Click({ Invoke-Refresh })

# =============================================================================
# Bulk operation confirmation dialog. Renders the full impact list, asks for
# explicit confirmation, then runs the per-op action scriptblock and logs
# success/failure per row. Every destructive bulk action funnels through here
# so the user always sees exactly what is about to happen before it happens.
# =============================================================================
function Show-BulkOpDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$ActionVerb,
        [Parameter(Mandatory)][array]$Operations,
        [Parameter(Mandatory)][scriptblock]$ActionScript
    )

    if (@($Operations).Count -eq 0) {
        Add-LogLine ('{0}: no matching operations to perform.' -f $ActionVerb)
        return $false
    }

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="" Width="780" Height="600" MinWidth="640" MinHeight="480"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1" ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="120"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="16,12,16,12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock x:Name="txtImpactHeader" Grid.Row="0" FontSize="13" Margin="0,4,0,8" TextWrapping="Wrap"/>
        <DataGrid x:Name="gridOps" Grid.Row="1" AutoGenerateColumns="False"
                  CanUserAddRows="False" CanUserDeleteRows="False" IsReadOnly="True"
                  SelectionMode="Single" SelectionUnit="FullRow"
                  Background="{DynamicResource MahApps.Brushes.ThemeBackground}"
                  Foreground="{DynamicResource MahApps.Brushes.ThemeForeground}"
                  AlternatingRowBackground="{DynamicResource MahApps.Brushes.Control.Background}"
                  RowHeaderWidth="0" BorderThickness="1"
                  BorderBrush="{DynamicResource MahApps.Brushes.Gray8}"
                  GridLinesVisibility="Horizontal" FontSize="11">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Result"   Width="80"  Binding="{Binding Result}"/>
                <DataGridTextColumn Header="DP"       Width="2*"  Binding="{Binding DPName}"/>
                <DataGridTextColumn Header="Content"  Width="3*"  Binding="{Binding ContentName}"/>
                <DataGridTextColumn Header="PackageID" Width="100" Binding="{Binding PackageID}"/>
                <DataGridTextColumn Header="Detail"   Width="2*"  Binding="{Binding Detail}"/>
            </DataGrid.Columns>
        </DataGrid>
        <TextBlock x:Name="txtBulkStatus" Grid.Row="2" FontSize="11" Margin="0,8,0,0"
                   Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="btnRun"   Content="" Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
            <Button x:Name="btnClose" Content="Close" Style="{StaticResource DialogButton}" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    $dlg.Title = $Title
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogTheme -Dialog $dlg

    $txtImpactHeader = $dlg.FindName('txtImpactHeader')
    $gridOps         = $dlg.FindName('gridOps')
    $txtBulkStatus   = $dlg.FindName('txtBulkStatus')
    $btnRun          = $dlg.FindName('btnRun')
    $btnClose        = $dlg.FindName('btnClose')

    $opCount = @($Operations).Count
    $txtImpactHeader.Text = ('You are about to {0} {1} (DP, Content) pair(s). Review the list below, then click "{0} {1}" to proceed.' -f $ActionVerb, $opCount)
    $btnRun.Content = ('{0} {1}' -f $ActionVerb, $opCount)

    # Decorate ops with mutable Result/Detail fields for live status updates
    $script:BulkOpRows = New-Object System.Collections.ObjectModel.ObservableCollection[psobject]
    foreach ($op in @($Operations)) {
        $row = [PSCustomObject]@{
            Result      = '-'
            DPName      = $op.DPName
            ContentName = $op.ContentName
            PackageID   = $op.PackageID
            Detail      = ''
            Op          = $op
        }
        $script:BulkOpRows.Add($row)
    }
    $gridOps.ItemsSource = $script:BulkOpRows
    $txtBulkStatus.Text = ('Ready. {0} operation(s) queued.' -f $opCount)

    $script:BulkOpDone = $false
    $btnRun.Add_Click({
        if ($script:BulkOpDone) { return }
        $btnRun.IsEnabled = $false
        $okCount = 0
        $failCount = 0
        $i = 0
        foreach ($row in $script:BulkOpRows) {
            $i++
            $txtBulkStatus.Text = ('Running {0} of {1}: {2} -> {3}' -f $i, $opCount, $row.ContentName, $row.DPName)
            $dlg.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
            try {
                & $ActionScript $row.Op
                $row.Result = 'OK'
                $row.Detail = ''
                $okCount++
            } catch {
                $row.Result = 'FAIL'
                $row.Detail = $_.Exception.Message
                $failCount++
            }
            $gridOps.Items.Refresh()
        }
        $txtBulkStatus.Text = ('Done. {0} succeeded, {1} failed. (Errors are in the Detail column.)' -f $okCount, $failCount)
        Add-LogLine ('{0}: {1} succeeded, {2} failed across {3} operations.' -f $ActionVerb, $okCount, $failCount, $opCount)
        $script:BulkOpDone = $true
    })
    $btnClose.Add_Click({ $dlg.DialogResult = $script:BulkOpDone; $dlg.Close() })

    [void]$dlg.ShowDialog()
    return $script:BulkOpDone
}

# =============================================================================
# Wired action button handlers.
# =============================================================================
function Get-FailedOpsForContent {
    param([Parameter(Mandatory)][array]$ContentSelection)
    $ops = @()
    foreach ($c in $ContentSelection) {
        $packageId = [string]$c.PackageID
        $name      = [string]$c.Name
        $failures = @($script:RawStatus | Where-Object { $_.PackageID -eq $packageId -and [int]$_.State -in 3,6 })
        foreach ($f in $failures) {
            $ops += [PSCustomObject]@{ DPName = $f.DPName; PackageID = $packageId; ContentName = $name }
        }
    }
    return $ops
}

function Get-OpsForContentToAllDPs {
    param([Parameter(Mandatory)][array]$ContentSelection)
    $ops = @()
    foreach ($c in $ContentSelection) {
        $packageId = [string]$c.PackageID
        $name      = [string]$c.Name
        $rows = @($script:RawStatus | Where-Object { $_.PackageID -eq $packageId })
        foreach ($r in $rows) {
            $ops += [PSCustomObject]@{ DPName = $r.DPName; PackageID = $packageId; ContentName = $name }
        }
    }
    return $ops
}

$btnRedistribute.Add_Click({
    $sel = @($gridContent.SelectedItems)
    if ($sel.Count -eq 0) { Add-LogLine 'Redistribute Failed: select one or more content rows first.'; return }
    if (-not $script:IsConnectedFromBg) { Add-LogLine 'Redistribute Failed: refresh first to establish a CM connection.'; return }
    $ops = Get-FailedOpsForContent -ContentSelection $sel
    if (Show-BulkOpDialog -Title 'Redistribute Failed Content' -ActionVerb 'Redistribute' -Operations $ops -ActionScript {
        param($op) Invoke-RedistributeContent -PackageID $op.PackageID -DPName $op.DPName
    }) { Invoke-Refresh }
})

$btnRemoveFromDP.Add_Click({
    $sel = @($gridContent.SelectedItems)
    if ($sel.Count -eq 0) { Add-LogLine 'Remove from DP: select one or more content rows first.'; return }
    if (-not $script:IsConnectedFromBg) { Add-LogLine 'Remove from DP: refresh first to establish a CM connection.'; return }
    $ops = Get-OpsForContentToAllDPs -ContentSelection $sel
    if (Show-BulkOpDialog -Title 'Remove Content from DPs' -ActionVerb 'Remove' -Operations $ops -ActionScript {
        param($op) Remove-ContentFromDP -PackageID $op.PackageID -DPName $op.DPName
    }) { Invoke-Refresh }
})

$btnValidate.Add_Click({
    $sel = @($gridContent.SelectedItems)
    if ($sel.Count -eq 0) { Add-LogLine 'Validate: select one or more content rows first.'; return }
    if (-not $script:IsConnectedFromBg) { Add-LogLine 'Validate: refresh first to establish a CM connection.'; return }
    $ops = Get-OpsForContentToAllDPs -ContentSelection $sel
    if (Show-BulkOpDialog -Title 'Validate Content on DPs' -ActionVerb 'Validate' -Operations $ops -ActionScript {
        param($op) Invoke-ContentValidation -PackageID $op.PackageID -DPName $op.DPName
    }) { Invoke-Refresh }
})

$btnRedistributeIssues.Add_Click({
    $sel = @($gridStatusIssues.SelectedItems)
    if ($sel.Count -eq 0) { Add-LogLine 'Redistribute Selected Issues: multi-select issue rows first.'; return }
    if (-not $script:IsConnectedFromBg) { Add-LogLine 'Redistribute: refresh first.'; return }
    $ops = @($sel | ForEach-Object { [PSCustomObject]@{ DPName = $_.DPName; PackageID = $_.PackageID; ContentName = $_.ContentName } })
    if (Show-BulkOpDialog -Title 'Redistribute Selected Issues' -ActionVerb 'Redistribute' -Operations $ops -ActionScript {
        param($op) Invoke-RedistributeContent -PackageID $op.PackageID -DPName $op.DPName
    }) { Invoke-Refresh }
})

$btnRemoveIssues.Add_Click({
    $sel = @($gridStatusIssues.SelectedItems)
    if ($sel.Count -eq 0) { Add-LogLine 'Remove Failures: multi-select issue rows first.'; return }
    if (-not $script:IsConnectedFromBg) { Add-LogLine 'Remove: refresh first.'; return }
    $ops = @($sel | ForEach-Object { [PSCustomObject]@{ DPName = $_.DPName; PackageID = $_.PackageID; ContentName = $_.ContentName } })
    if (Show-BulkOpDialog -Title 'Remove Failures' -ActionVerb 'Remove' -Operations $ops -ActionScript {
        param($op) Remove-ContentFromDP -PackageID $op.PackageID -DPName $op.DPName
    }) { Invoke-Refresh }
})

# =============================================================================
# Export buttons.
# =============================================================================
function Get-ActiveExportRows {
    switch ($script:ActiveView) {
        'DPs'           { return @{ Name = 'DPs'           ; Rows = $gridDPs.ItemsSource } }
        'Content'       { return @{ Name = 'Content'       ; Rows = $gridContent.ItemsSource } }
        'Status Issues' { return @{ Name = 'StatusIssues'  ; Rows = $gridStatusIssues.ItemsSource } }
    }
}

function Save-RowsAsCsv {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Writes to a user-chosen path; idempotent.')]
    param([Parameter(Mandatory)][array]$Rows, [Parameter(Mandatory)][string]$OutputPath)
    $Rows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
}

$btnExportCsv.Add_Click({
    $info = Get-ActiveExportRows
    if (-not $info -or @($info.Rows).Count -eq 0) { Add-LogLine 'Export CSV: nothing to export.'; return }
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Filter = 'CSV files (*.csv)|*.csv'
    $sfd.FileName = ('DPCM-{0}-{1}.csv' -f $info.Name, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $reportsDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq $true) {
        Save-RowsAsCsv -Rows @($info.Rows) -OutputPath $sfd.FileName
        Add-LogLine ('Exported CSV: {0}' -f $sfd.FileName)
    }
})

$btnExportHtml.Add_Click({
    $info = Get-ActiveExportRows
    if (-not $info -or @($info.Rows).Count -eq 0) { Add-LogLine 'Export HTML: nothing to export.'; return }
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Filter = 'HTML files (*.html)|*.html'
    $sfd.FileName = ('DPCM-{0}-{1}.html' -f $info.Name, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $reportsDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq $true) {
        @($info.Rows) | ConvertTo-Html -Title ('DPCM ' + $info.Name) | Set-Content -LiteralPath $sfd.FileName -Encoding UTF8
        Add-LogLine ('Exported HTML: {0}' -f $sfd.FileName)
    }
})

function Show-OptionsDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param()
    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Options" Width="640" Height="380"
    MinWidth="560" MinHeight="380"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1" ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="CategoryRowStyle" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="Height" Value="36"/>
                <Setter Property="HorizontalContentAlignment" Value="Left"/>
                <Setter Property="Padding" Value="14,0,14,0"/>
                <Setter Property="FontSize" Value="13"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
                <Setter Property="Margin" Value="0"/>
            </Style>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="180"/>
            <ColumnDefinition Width="1"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Border Grid.Column="0" Grid.Row="0" Padding="6,12,0,12">
            <StackPanel>
                <Button x:Name="btnCatConnection" Content="Connection" Style="{StaticResource CategoryRowStyle}"/>
                <Button x:Name="btnCatAbout"      Content="About"      Style="{StaticResource CategoryRowStyle}"/>
            </StackPanel>
        </Border>
        <Border Grid.Column="1" Grid.Row="0" Background="{DynamicResource MahApps.Brushes.Gray8}"/>
        <Grid Grid.Column="2" Grid.Row="0" Margin="20,16,20,16">
            <StackPanel x:Name="paneConnection" Visibility="Visible">
                <TextBlock Text="MECM Connection" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                <TextBlock Text="Site Code" FontSize="11" Margin="0,4,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
                <TextBox x:Name="txtSiteCode" FontSize="12" Padding="6,4,6,4"
                         Controls:TextBoxHelper.Watermark="e.g. MCM" Width="120" HorizontalAlignment="Left"/>
                <TextBlock Text="SMS Provider FQDN" FontSize="11" Margin="0,12,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
                <TextBox x:Name="txtSmsProvider" FontSize="12" Padding="6,4,6,4"
                         Controls:TextBoxHelper.Watermark="e.g. cm01.contoso.com"/>
                <TextBlock Text="Used for the CM PSDrive root + WMI distribution-status queries. The account running this app needs read access to the SMS provider for status retrieval and write access for redistribute / remove operations."
                           FontSize="11" TextWrapping="Wrap" Margin="0,16,0,0"
                           Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            </StackPanel>
            <StackPanel x:Name="paneAbout" Visibility="Collapsed">
                <TextBlock Text="About" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                <TextBlock Text="DP Content Manager v1.2.1" FontSize="13" FontWeight="SemiBold"/>
                <TextBlock Text="Audit DP content distribution at scale: per-DP and per-content rollups, a triage view for non-OK rows, bulk redistribute / remove / validate. WMI bulk-status pull keeps the round-trip count down even on large environments."
                           FontSize="12" TextWrapping="Wrap" Margin="0,8,0,0"/>
                <TextBlock Text="Author: Jason Ulbright. License: MIT."
                           FontSize="11" Margin="0,16,0,0" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            </StackPanel>
        </Grid>
        <Border Grid.Row="1" Grid.ColumnSpan="3" Padding="16,12,16,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnOk"     Content="OK"     Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
                <Button x:Name="btnCancel" Content="Cancel" Style="{StaticResource DialogButton}"        IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogTheme -Dialog $dlg
    $btnCatConnection = $dlg.FindName('btnCatConnection')
    $btnCatAbout      = $dlg.FindName('btnCatAbout')
    $paneConnection   = $dlg.FindName('paneConnection')
    $paneAbout        = $dlg.FindName('paneAbout')
    $txtSiteCode      = $dlg.FindName('txtSiteCode')
    $txtSmsProvider   = $dlg.FindName('txtSmsProvider')
    $btnOk            = $dlg.FindName('btnOk')
    $btnCancel        = $dlg.FindName('btnCancel')
    $txtSiteCode.Text    = [string]$global:Prefs.SiteCode
    $txtSmsProvider.Text = [string]$global:Prefs.SMSProvider
    $btnCatConnection.Add_Click({ $paneConnection.Visibility = [System.Windows.Visibility]::Visible; $paneAbout.Visibility = [System.Windows.Visibility]::Collapsed })
    $btnCatAbout.Add_Click({      $paneConnection.Visibility = [System.Windows.Visibility]::Collapsed; $paneAbout.Visibility = [System.Windows.Visibility]::Visible })
    $btnOk.Add_Click({
        $newSite     = ([string]$txtSiteCode.Text).Trim()
        $newProvider = ([string]$txtSmsProvider.Text).Trim()
        $changed = ($newSite -ne [string]$global:Prefs.SiteCode) -or ($newProvider -ne [string]$global:Prefs.SMSProvider)
        $global:Prefs.SiteCode    = $newSite
        $global:Prefs.SMSProvider = $newProvider
        Save-DpcmPreferences -Prefs $global:Prefs
        if ($changed) {
            Dispose-BgWork
            Close-SuiteBgRunspace -Runspace $script:BgRunspace
            $script:BgRunspace = $null
            $script:BgState = $null
            $script:IsConnectedFromBg = $false
            $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
            $btnRefresh.IsEnabled = $true
        }
        $dlg.DialogResult = $true; $dlg.Close()
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })
    [void]$dlg.ShowDialog()
    Update-StatusBarSummary
}
$btnOptions.Add_Click({ Show-OptionsDialog })

# === Window state (geometry logic: SuiteCommon) ===
$global:WindowStatePath = Join-Path $PSScriptRoot 'DPContentMgr.windowstate.json'

$window.Add_Closing({
    Save-WindowState -Window $window -Path $global:WindowStatePath -ExtraState @{ ActiveView = $script:ActiveView }
    Dispose-BgWork
    Close-SuiteBgRunspace -Runspace $script:BgRunspace
})

$window.Add_Loaded({
    Restore-WindowState -Window $window -Path $global:WindowStatePath -OnStateLoaded {
        param($s)
        if ($s.ActiveView -in @('DPs','Content','Status Issues')) { Set-ActiveView -View ([string]$s.ActiveView) }
    }
    $isDark = [bool]$global:Prefs['DarkMode']
    if (-not $isDark) { [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Light.Blue') }
    Update-TitleBarBrushes
    Update-ActionBarVisibility
    Update-StatusBarSummary
    Add-LogLine 'DP Content Manager ready. Configure Site / Provider in Options, then click Refresh.'
})

[void]$window.ShowDialog()
try { Stop-Transcript | Out-Null } catch { $null = $_ }
