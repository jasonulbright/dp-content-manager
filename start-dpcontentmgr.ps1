<#
.SYNOPSIS
    WinForms front-end for DP Content Manager - MECM distribution point content status and management.

.DESCRIPTION
    Provides a GUI for viewing and managing MECM distribution point content.
    On launch, performs LOCAL-ONLY operations (no network access).
    Network operations occur when the user clicks Load/Refresh.

    Features:
      - View content status across all DPs (content-centric and DP-centric views)
      - Redistribute failed content
      - Remove orphaned content
      - Validate content integrity
      - Storage analysis per DP
      - Export results to CSV or HTML

.EXAMPLE
    .\start-dpcontentmgr.ps1

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.8+
      - Windows Forms (System.Windows.Forms)
      - Configuration Manager console installed

    ScriptName : start-dpcontentmgr.ps1
    Purpose    : WinForms front-end for MECM DP content management
    Version    : 1.0.0
    Updated    : 2026-02-26
#>

param()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }

$moduleRoot = Join-Path $PSScriptRoot "Module"
Import-Module (Join-Path $moduleRoot "DPContentMgrCommon.psd1") -Force

# Initialize tool logging
$toolLogFolder = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $toolLogFolder)) {
    New-Item -ItemType Directory -Path $toolLogFolder -Force | Out-Null
}
$toolLogPath = Join-Path $toolLogFolder ("DPContentMgr-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $toolLogPath

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Set-ModernButtonStyle {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Button]$Button,
        [Parameter(Mandatory)][System.Drawing.Color]$BackColor
    )

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $BackColor
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.UseVisualStyleBackColor = $false
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand

    $hover = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $BackColor.R - 18),
        [Math]::Max(0, $BackColor.G - 18),
        [Math]::Max(0, $BackColor.B - 18)
    )
    $down = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $BackColor.R - 36),
        [Math]::Max(0, $BackColor.G - 36),
        [Math]::Max(0, $BackColor.B - 36)
    )

    $Button.FlatAppearance.MouseOverBackColor = $hover
    $Button.FlatAppearance.MouseDownBackColor = $down
}

function Enable-DoubleBuffer {
    param([Parameter(Mandatory)][System.Windows.Forms.Control]$Control)
    $prop = $Control.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance,NonPublic")
    if ($prop) { $prop.SetValue($Control, $true, $null) | Out-Null }
}

function Add-LogLine {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TextBox,
        [Parameter(Mandatory)][string]$Message
    )
    $ts = (Get-Date).ToString("HH:mm:ss")
    $line = "{0}  {1}" -f $ts, $Message

    if ([string]::IsNullOrWhiteSpace($TextBox.Text)) {
        $TextBox.Text = $line
    }
    else {
        $TextBox.AppendText([Environment]::NewLine + $line)
    }

    $TextBox.SelectionStart = $TextBox.TextLength
    $TextBox.ScrollToCaret()
}

function Save-WindowState {
    $statePath = Join-Path $PSScriptRoot "DPContentMgr.windowstate.json"
    $state = @{
        X                = $form.Location.X
        Y                = $form.Location.Y
        Width            = $form.Size.Width
        Height           = $form.Size.Height
        Maximized        = ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Maximized)
        SplitterDistance = $splitContent.SplitterDistance
        ActiveTab        = $tabMain.SelectedIndex
    }
    $state | ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding UTF8
}

function Restore-WindowState {
    $statePath = Join-Path $PSScriptRoot "DPContentMgr.windowstate.json"
    if (-not (Test-Path -LiteralPath $statePath)) { return }

    try {
        $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
        if ($state.Maximized) {
            $form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
        } else {
            $form.Location = New-Object System.Drawing.Point($state.X, $state.Y)
            $form.Size = New-Object System.Drawing.Size($state.Width, $state.Height)
        }
        if ($state.SplitterDistance) {
            $splitContent.SplitterDistance = [int]$state.SplitterDistance
        }
        if ($null -ne $state.ActiveTab) {
            $tabMain.SelectedIndex = [int]$state.ActiveTab
        }
    } catch { }
}

# ---------------------------------------------------------------------------
# Preferences
# ---------------------------------------------------------------------------

function Get-DcmPreferences {
    $prefsPath = Join-Path $PSScriptRoot "DPContentMgr.prefs.json"
    $defaults = @{ DarkMode = $false; SiteCode = ''; SMSProvider = '' }

    if (Test-Path -LiteralPath $prefsPath) {
        try {
            $loaded = Get-Content -LiteralPath $prefsPath -Raw | ConvertFrom-Json
            if ($null -ne $loaded.DarkMode)    { $defaults.DarkMode    = [bool]$loaded.DarkMode }
            if ($loaded.SiteCode)              { $defaults.SiteCode    = $loaded.SiteCode }
            if ($loaded.SMSProvider)           { $defaults.SMSProvider = $loaded.SMSProvider }
        } catch { }
    }

    return $defaults
}

function Save-DcmPreferences {
    param([hashtable]$Prefs)
    $prefsPath = Join-Path $PSScriptRoot "DPContentMgr.prefs.json"
    $Prefs | ConvertTo-Json | Set-Content -LiteralPath $prefsPath -Encoding UTF8
}

$script:Prefs = Get-DcmPreferences

# ---------------------------------------------------------------------------
# Colors (theme-aware)
# ---------------------------------------------------------------------------

$clrAccent = [System.Drawing.Color]::FromArgb(0, 120, 212)

if ($script:Prefs.DarkMode) {
    $clrFormBg     = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $clrPanelBg    = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $clrHint       = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrSubtitle   = [System.Drawing.Color]::FromArgb(180, 200, 220)
    $clrGridAlt    = [System.Drawing.Color]::FromArgb(48, 48, 48)
    $clrGridLine   = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $clrDetailBg   = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $clrSepLine    = [System.Drawing.Color]::FromArgb(55, 55, 55)
    $clrInputBdr   = [System.Drawing.Color]::FromArgb(70, 70, 70)
    $clrLogBg      = [System.Drawing.Color]::FromArgb(35, 35, 35)
    $clrLogFg      = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $clrText       = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrGridText   = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrErrText    = [System.Drawing.Color]::FromArgb(255, 100, 100)
    $clrWarnText   = [System.Drawing.Color]::FromArgb(255, 200, 80)
    $clrOkText     = [System.Drawing.Color]::FromArgb(80, 200, 80)
} else {
    $clrFormBg     = [System.Drawing.Color]::FromArgb(245, 246, 248)
    $clrPanelBg    = [System.Drawing.Color]::White
    $clrHint       = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrSubtitle   = [System.Drawing.Color]::FromArgb(220, 230, 245)
    $clrGridAlt    = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $clrGridLine   = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $clrDetailBg   = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $clrSepLine    = [System.Drawing.Color]::FromArgb(218, 220, 224)
    $clrInputBdr   = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $clrLogBg      = [System.Drawing.Color]::White
    $clrLogFg      = [System.Drawing.Color]::Black
    $clrText       = [System.Drawing.Color]::Black
    $clrGridText   = [System.Drawing.Color]::Black
    $clrErrText    = [System.Drawing.Color]::FromArgb(180, 0, 0)
    $clrWarnText   = [System.Drawing.Color]::FromArgb(180, 120, 0)
    $clrOkText     = [System.Drawing.Color]::FromArgb(34, 139, 34)
}

# Custom dark mode ToolStrip renderer
if ($script:Prefs.DarkMode) {
    if (-not ('DarkToolStripRenderer' -as [type])) {
        $rendererCs = (
            'using System.Drawing;',
            'using System.Windows.Forms;',
            'public class DarkToolStripRenderer : ToolStripProfessionalRenderer {',
            '    private Color _bg;',
            '    public DarkToolStripRenderer(Color bg) : base() { _bg = bg; }',
            '    protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) { }',
            '    protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) {',
            '        using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); }',
            '    }',
            '    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e) {',
            '        if (e.Item.Selected || e.Item.Pressed) {',
            '            using (var b = new SolidBrush(Color.FromArgb(60, 60, 60)))',
            '            { e.Graphics.FillRectangle(b, new Rectangle(Point.Empty, e.Item.Size)); }',
            '        }',
            '    }',
            '    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e) {',
            '        int y = e.Item.Height / 2;',
            '        using (var p = new Pen(Color.FromArgb(70, 70, 70)))',
            '        { e.Graphics.DrawLine(p, 0, y, e.Item.Width, y); }',
            '    }',
            '    protected override void OnRenderImageMargin(ToolStripRenderEventArgs e) {',
            '        using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); }',
            '    }',
            '}'
        ) -join "`r`n"
        Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition $rendererCs
    }
    $script:DarkRenderer = New-Object DarkToolStripRenderer($clrPanelBg)
}

# ---------------------------------------------------------------------------
# Dialogs
# ---------------------------------------------------------------------------

function Show-PreferencesDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Preferences"
    $dlg.Size = New-Object System.Drawing.Size(420, 300)
    $dlg.MinimumSize = $dlg.Size
    $dlg.MaximumSize = $dlg.Size
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $dlg.BackColor = $clrFormBg

    # Appearance
    $grpAppearance = New-Object System.Windows.Forms.GroupBox
    $grpAppearance.Text = "Appearance"
    $grpAppearance.SetBounds(16, 12, 372, 60)
    $grpAppearance.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpAppearance.ForeColor = $clrText
    $grpAppearance.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpAppearance.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpAppearance.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpAppearance)

    $chkDark = New-Object System.Windows.Forms.CheckBox
    $chkDark.Text = "Enable dark mode (requires restart)"
    $chkDark.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkDark.AutoSize = $true
    $chkDark.Location = New-Object System.Drawing.Point(14, 24)
    $chkDark.Checked = $script:Prefs.DarkMode
    $chkDark.ForeColor = $clrText
    $chkDark.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $chkDark.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkDark.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
    $grpAppearance.Controls.Add($chkDark)

    # MECM Connection
    $grpConn = New-Object System.Windows.Forms.GroupBox
    $grpConn.Text = "MECM Connection"
    $grpConn.SetBounds(16, 82, 372, 110)
    $grpConn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpConn.ForeColor = $clrText
    $grpConn.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpConn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpConn.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpConn)

    $lblSiteCode = New-Object System.Windows.Forms.Label
    $lblSiteCode.Text = "Site Code:"
    $lblSiteCode.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSiteCode.Location = New-Object System.Drawing.Point(14, 30)
    $lblSiteCode.AutoSize = $true
    $lblSiteCode.ForeColor = $clrText
    $grpConn.Controls.Add($lblSiteCode)

    $txtSiteCodePref = New-Object System.Windows.Forms.TextBox
    $txtSiteCodePref.SetBounds(130, 27, 80, 24)
    $txtSiteCodePref.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtSiteCodePref.MaxLength = 3
    $txtSiteCodePref.Text = $script:Prefs.SiteCode
    $txtSiteCodePref.BackColor = $clrDetailBg
    $txtSiteCodePref.ForeColor = $clrText
    $txtSiteCodePref.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
    $grpConn.Controls.Add($txtSiteCodePref)

    $lblServer = New-Object System.Windows.Forms.Label
    $lblServer.Text = "SMS Provider:"
    $lblServer.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblServer.Location = New-Object System.Drawing.Point(14, 64)
    $lblServer.AutoSize = $true
    $lblServer.ForeColor = $clrText
    $grpConn.Controls.Add($lblServer)

    $txtServerPref = New-Object System.Windows.Forms.TextBox
    $txtServerPref.SetBounds(130, 61, 220, 24)
    $txtServerPref.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtServerPref.Text = $script:Prefs.SMSProvider
    $txtServerPref.BackColor = $clrDetailBg
    $txtServerPref.ForeColor = $clrText
    $txtServerPref.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
    $grpConn.Controls.Add($txtServerPref)

    # OK / Cancel
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Size = New-Object System.Drawing.Size(90, 32)
    $btnOK.Location = New-Object System.Drawing.Point(208, 210)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    Set-ModernButtonStyle -Button $btnOK -BackColor $clrAccent
    $dlg.Controls.Add($btnOK)
    $dlg.AcceptButton = $btnOK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(90, 32)
    $btnCancel.Location = New-Object System.Drawing.Point(306, 210)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.FlatAppearance.BorderColor = $clrSepLine
    $btnCancel.ForeColor = $clrText
    $btnCancel.BackColor = $clrFormBg
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel

    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $darkChanged = ($chkDark.Checked -ne $script:Prefs.DarkMode)
        $script:Prefs.DarkMode    = $chkDark.Checked
        $script:Prefs.SiteCode    = $txtSiteCodePref.Text.Trim().ToUpper()
        $script:Prefs.SMSProvider = $txtServerPref.Text.Trim()
        Save-DcmPreferences -Prefs $script:Prefs

        # Update connection bar labels
        $lblSiteVal.Text   = if ($script:Prefs.SiteCode)    { $script:Prefs.SiteCode }    else { '(not set)' }
        $lblServerVal.Text = if ($script:Prefs.SMSProvider)  { $script:Prefs.SMSProvider }  else { '(not set)' }

        if ($darkChanged) {
            $restart = [System.Windows.Forms.MessageBox]::Show(
                "Theme change requires a restart. Restart now?",
                "Restart Required",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($restart -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process powershell -ArgumentList @('-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
                $form.Close()
            }
        }
    }

    $dlg.Dispose()
}

function Show-AboutDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "About DP Content Manager"
    $dlg.Size = New-Object System.Drawing.Size(460, 320)
    $dlg.MinimumSize = $dlg.Size
    $dlg.MaximumSize = $dlg.Size
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $dlg.BackColor = $clrFormBg

    $lblAboutTitle = New-Object System.Windows.Forms.Label
    $lblAboutTitle.Text = "DP Content Manager"
    $lblAboutTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblAboutTitle.ForeColor = $clrAccent
    $lblAboutTitle.AutoSize = $true
    $lblAboutTitle.BackColor = $clrFormBg
    $lblAboutTitle.Location = New-Object System.Drawing.Point(110, 30)
    $dlg.Controls.Add($lblAboutTitle)

    $lblVersion = New-Object System.Windows.Forms.Label
    $lblVersion.Text = "DP Content Manager v1.0.0"
    $lblVersion.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblVersion.ForeColor = $clrText
    $lblVersion.AutoSize = $true
    $lblVersion.BackColor = $clrFormBg
    $lblVersion.Location = New-Object System.Drawing.Point(120, 60)
    $dlg.Controls.Add($lblVersion)

    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = ("Unified MECM distribution point content management." +
        " View content status across all DPs, redistribute failed content," +
        " remove orphaned content, validate integrity, and analyze storage.")
    $lblDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblDesc.ForeColor = $clrText
    $lblDesc.SetBounds(30, 100, 390, 80)
    $lblDesc.BackColor = $clrFormBg
    $lblDesc.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
    $dlg.Controls.Add($lblDesc)

    $lblCopyright = New-Object System.Windows.Forms.Label
    $lblCopyright.Text = "(c) 2026 - All rights reserved"
    $lblCopyright.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)
    $lblCopyright.ForeColor = $clrHint
    $lblCopyright.AutoSize = $true
    $lblCopyright.BackColor = $clrFormBg
    $lblCopyright.Location = New-Object System.Drawing.Point(142, 200)
    $dlg.Controls.Add($lblCopyright)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "OK"
    $btnClose.Size = New-Object System.Drawing.Size(90, 32)
    $btnClose.Location = New-Object System.Drawing.Point(175, 240)
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    Set-ModernButtonStyle -Button $btnClose -BackColor $clrAccent
    $dlg.Controls.Add($btnClose)
    $dlg.AcceptButton = $btnClose

    [void]$dlg.ShowDialog($form)
    $dlg.Dispose()
}

function Show-StorageAnalysisDialog {
    if (-not $script:AllStatus -or -not $script:AllContent) {
        [System.Windows.Forms.MessageBox]::Show("Load data first by clicking Refresh All.", "No Data", "OK", "Information") | Out-Null
        return
    }

    $analysis = Get-DPStorageAnalysis -StatusRows $script:AllStatus -ContentObjects $script:AllContent

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Storage Analysis"
    $dlg.Size = New-Object System.Drawing.Size(800, 600)
    $dlg.MinimumSize = New-Object System.Drawing.Size(600, 400)
    $dlg.StartPosition = "CenterParent"
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $dlg.BackColor = $clrFormBg

    $totalGB = ($analysis | Measure-Object -Property TotalSizeGB -Sum).Sum
    $lblSummary = New-Object System.Windows.Forms.Label
    $lblSummary.Text = "Total distributed content: {0:N1} GB across {1} DPs" -f $totalGB, $analysis.Count
    $lblSummary.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblSummary.ForeColor = $clrAccent
    $lblSummary.Dock = [System.Windows.Forms.DockStyle]::Top
    $lblSummary.Height = 36
    $lblSummary.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 0)
    $lblSummary.BackColor = $clrFormBg
    $dlg.Controls.Add($lblSummary)

    $storageGrid = New-Object System.Windows.Forms.DataGridView
    $storageGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $storageGrid.ReadOnly = $true
    $storageGrid.AllowUserToAddRows = $false
    $storageGrid.AllowUserToDeleteRows = $false
    $storageGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $storageGrid.AutoGenerateColumns = $false
    $storageGrid.RowHeadersVisible = $false
    $storageGrid.BackgroundColor = $clrPanelBg
    $storageGrid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $storageGrid.GridColor = $clrGridLine
    $storageGrid.ColumnHeadersDefaultCellStyle.BackColor = $clrAccent
    $storageGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $storageGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $storageGrid.ColumnHeadersHeight = 32
    $storageGrid.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    $storageGrid.EnableHeadersVisualStyles = $false
    $storageGrid.DefaultCellStyle.BackColor = $clrPanelBg
    $storageGrid.DefaultCellStyle.ForeColor = $clrGridText
    $storageGrid.DefaultCellStyle.SelectionBackColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(38, 79, 120) } else { [System.Drawing.Color]::FromArgb(0, 120, 215) }
    $storageGrid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $storageGrid.AlternatingRowsDefaultCellStyle.BackColor = $clrGridAlt
    $storageGrid.RowTemplate.Height = 26
    Enable-DoubleBuffer -Control $storageGrid

    $colSDp      = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSDp.HeaderText = "DP Name"; $colSDp.DataPropertyName = "DPName"; $colSDp.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $colSSize    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSSize.HeaderText = "Total Size (GB)"; $colSSize.DataPropertyName = "TotalSizeGB"; $colSSize.Width = 110
    $colSCount   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSCount.HeaderText = "Content Count"; $colSCount.DataPropertyName = "ContentCount"; $colSCount.Width = 100
    $colSFailed  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSFailed.HeaderText = "Failed"; $colSFailed.DataPropertyName = "FailedCount"; $colSFailed.Width = 70
    $storageGrid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colSDp, $colSSize, $colSCount, $colSFailed))

    $dtStorage = New-Object System.Data.DataTable
    [void]$dtStorage.Columns.Add("DPName", [string])
    [void]$dtStorage.Columns.Add("TotalSizeGB", [double])
    [void]$dtStorage.Columns.Add("ContentCount", [int])
    [void]$dtStorage.Columns.Add("FailedCount", [int])

    foreach ($item in $analysis) {
        [void]$dtStorage.Rows.Add($item.DPName, $item.TotalSizeGB, $item.ContentCount, $item.FailedCount)
    }
    $storageGrid.DataSource = $dtStorage
    $dlg.Controls.Add($storageGrid)
    $storageGrid.BringToFront()

    $pnlStorageBtn = New-Object System.Windows.Forms.Panel
    $pnlStorageBtn.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $pnlStorageBtn.Height = 48
    $pnlStorageBtn.BackColor = $clrFormBg
    $pnlStorageBtn.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 8)
    $dlg.Controls.Add($pnlStorageBtn)

    $btnStorageExport = New-Object System.Windows.Forms.Button
    $btnStorageExport.Text = "Export CSV"
    $btnStorageExport.Size = New-Object System.Drawing.Size(120, 32)
    $btnStorageExport.Location = New-Object System.Drawing.Point(12, 8)
    Set-ModernButtonStyle -Button $btnStorageExport -BackColor ([System.Drawing.Color]::FromArgb(34, 139, 34))
    $btnStorageExport.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV Files (*.csv)|*.csv"
        $sfd.FileName = "StorageAnalysis-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
        $sfd.InitialDirectory = Join-Path $PSScriptRoot "Reports"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Export-ContentStatusCsv -DataTable $dtStorage -OutputPath $sfd.FileName
            Add-LogLine -TextBox $txtLog -Message "Storage analysis exported to $($sfd.FileName)"
        }
    })
    $pnlStorageBtn.Controls.Add($btnStorageExport)

    [void]$dlg.ShowDialog($form)
    $dlg.Dispose()
}

# ---------------------------------------------------------------------------
# Form
# ---------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "DP Content Manager"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(1440, 900)
$form.MinimumSize = New-Object System.Drawing.Size(1100, 750)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$form.BackColor = $clrFormBg
$form.Icon = [System.Drawing.SystemIcons]::Application

# ---------------------------------------------------------------------------
# Menu bar
# ---------------------------------------------------------------------------

$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.Dock = [System.Windows.Forms.DockStyle]::Top
$menuStrip.BackColor = $clrPanelBg
$menuStrip.ForeColor = $clrText
if ($script:DarkRenderer) {
    $menuStrip.Renderer = $script:DarkRenderer
} else {
    $menuStrip.RenderMode = [System.Windows.Forms.ToolStripRenderMode]::System
}
$menuStrip.Padding = New-Object System.Windows.Forms.Padding(4, 2, 0, 0)

# File menu
$mnuFile = New-Object System.Windows.Forms.ToolStripMenuItem("&File")
$mnuFilePrefs = New-Object System.Windows.Forms.ToolStripMenuItem("&Preferences...")
$mnuFilePrefs.Add_Click({ Show-PreferencesDialog })
$mnuFileSep = New-Object System.Windows.Forms.ToolStripSeparator
$mnuFileExit = New-Object System.Windows.Forms.ToolStripMenuItem("E&xit")
$mnuFileExit.Add_Click({ $form.Close() })
[void]$mnuFile.DropDownItems.Add($mnuFilePrefs)
[void]$mnuFile.DropDownItems.Add($mnuFileSep)
[void]$mnuFile.DropDownItems.Add($mnuFileExit)

# Actions menu
$mnuActions = New-Object System.Windows.Forms.ToolStripMenuItem("&Actions")
$mnuActRefresh = New-Object System.Windows.Forms.ToolStripMenuItem("&Refresh All")
$mnuActRefresh.Add_Click({ Invoke-RefreshAll })
$mnuActRedist = New-Object System.Windows.Forms.ToolStripMenuItem("Re&distribute Selected")
$mnuActRedist.Add_Click({ Invoke-RedistributeSelected })
$mnuActRemoveOrphan = New-Object System.Windows.Forms.ToolStripMenuItem("Remove &Orphaned Content")
$mnuActRemoveOrphan.Add_Click({ Invoke-RemoveOrphaned })
$mnuActValidate = New-Object System.Windows.Forms.ToolStripMenuItem("&Validate Selected")
$mnuActValidate.Add_Click({ Invoke-ValidateSelected })
[void]$mnuActions.DropDownItems.Add($mnuActRefresh)
[void]$mnuActions.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$mnuActions.DropDownItems.Add($mnuActRedist)
[void]$mnuActions.DropDownItems.Add($mnuActValidate)
[void]$mnuActions.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$mnuActions.DropDownItems.Add($mnuActRemoveOrphan)

# View menu
$mnuView = New-Object System.Windows.Forms.ToolStripMenuItem("&View")
$mnuViewContent = New-Object System.Windows.Forms.ToolStripMenuItem("&Content View")
$mnuViewContent.Add_Click({ $tabMain.SelectedIndex = 0 })
$mnuViewDP = New-Object System.Windows.Forms.ToolStripMenuItem("&DP View")
$mnuViewDP.Add_Click({ $tabMain.SelectedIndex = 1 })
$mnuViewSep = New-Object System.Windows.Forms.ToolStripSeparator
$mnuViewStorage = New-Object System.Windows.Forms.ToolStripMenuItem("&Storage Analysis...")
$mnuViewStorage.Add_Click({ Show-StorageAnalysisDialog })
[void]$mnuView.DropDownItems.Add($mnuViewContent)
[void]$mnuView.DropDownItems.Add($mnuViewDP)
[void]$mnuView.DropDownItems.Add($mnuViewSep)
[void]$mnuView.DropDownItems.Add($mnuViewStorage)

# Help menu
$mnuHelp = New-Object System.Windows.Forms.ToolStripMenuItem("&Help")
$mnuHelpAbout = New-Object System.Windows.Forms.ToolStripMenuItem("&About...")
$mnuHelpAbout.Add_Click({ Show-AboutDialog })
[void]$mnuHelp.DropDownItems.Add($mnuHelpAbout)

[void]$menuStrip.Items.Add($mnuFile)
[void]$menuStrip.Items.Add($mnuActions)
[void]$menuStrip.Items.Add($mnuView)
[void]$menuStrip.Items.Add($mnuHelp)
$form.MainMenuStrip = $menuStrip

# ---------------------------------------------------------------------------
# StatusStrip (Dock:Bottom - add FIRST so it stays at very bottom)
# ---------------------------------------------------------------------------

$status = New-Object System.Windows.Forms.StatusStrip
$status.BackColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(45, 45, 45) } else { [System.Drawing.Color]::FromArgb(240, 240, 240) }
$status.ForeColor = $clrText
$status.Dock = [System.Windows.Forms.DockStyle]::Bottom
if ($script:DarkRenderer) {
    $status.Renderer = $script:DarkRenderer
} else {
    $status.RenderMode = [System.Windows.Forms.ToolStripRenderMode]::System
}
$status.SizingGrip = $false
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Disconnected. Configure site in File > Preferences, then click Refresh All."
$statusLabel.ForeColor = $clrText
$status.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($status)

# ---------------------------------------------------------------------------
# Log console panel (Dock:Bottom)
# ---------------------------------------------------------------------------

$pnlLog = New-Object System.Windows.Forms.Panel
$pnlLog.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlLog.Height = 95
$pnlLog.Padding = New-Object System.Windows.Forms.Padding(12, 4, 12, 6)
$pnlLog.BackColor = $clrFormBg
$form.Controls.Add($pnlLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = if ($script:Prefs.DarkMode) { [System.Windows.Forms.ScrollBars]::None } else { [System.Windows.Forms.ScrollBars]::Vertical }
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.BackColor = $clrLogBg
$txtLog.ForeColor = $clrLogFg
$txtLog.WordWrap = $true
$txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$pnlLog.Controls.Add($txtLog)

# ---------------------------------------------------------------------------
# Button panel (Dock:Bottom)
# ---------------------------------------------------------------------------

$pnlButtons = New-Object System.Windows.Forms.Panel
$pnlButtons.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlButtons.Height = 56
$pnlButtons.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 4)
$pnlButtons.BackColor = $clrFormBg
$form.Controls.Add($pnlButtons)

$pnlSepButtons = New-Object System.Windows.Forms.Panel
$pnlSepButtons.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSepButtons.Height = 1
$pnlSepButtons.BackColor = $clrSepLine
$pnlButtons.Controls.Add($pnlSepButtons)

$flowButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$flowButtons.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowButtons.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$flowButtons.WrapContents = $false
$flowButtons.BackColor = $clrFormBg
$pnlButtons.Controls.Add($flowButtons)

$btnRedistribute = New-Object System.Windows.Forms.Button
$btnRedistribute.Text = "Redistribute Selected"
$btnRedistribute.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$btnRedistribute.Size = New-Object System.Drawing.Size(180, 38)
$btnRedistribute.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
Set-ModernButtonStyle -Button $btnRedistribute -BackColor ([System.Drawing.Color]::FromArgb(217, 95, 2))
$flowButtons.Controls.Add($btnRedistribute)

$btnValidate = New-Object System.Windows.Forms.Button
$btnValidate.Text = "Validate Selected"
$btnValidate.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$btnValidate.Size = New-Object System.Drawing.Size(160, 38)
$btnValidate.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
Set-ModernButtonStyle -Button $btnValidate -BackColor ([System.Drawing.Color]::FromArgb(100, 60, 160))
$flowButtons.Controls.Add($btnValidate)

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Remove Selected"
$btnRemove.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$btnRemove.Size = New-Object System.Drawing.Size(160, 38)
$btnRemove.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
Set-ModernButtonStyle -Button $btnRemove -BackColor ([System.Drawing.Color]::FromArgb(180, 30, 30))
$flowButtons.Controls.Add($btnRemove)

$btnExportCsv = New-Object System.Windows.Forms.Button
$btnExportCsv.Text = "Export CSV"
$btnExportCsv.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$btnExportCsv.Size = New-Object System.Drawing.Size(120, 38)
$btnExportCsv.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
Set-ModernButtonStyle -Button $btnExportCsv -BackColor ([System.Drawing.Color]::FromArgb(34, 139, 34))
$flowButtons.Controls.Add($btnExportCsv)

$btnExportHtml = New-Object System.Windows.Forms.Button
$btnExportHtml.Text = "Export HTML"
$btnExportHtml.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$btnExportHtml.Size = New-Object System.Drawing.Size(120, 38)
$btnExportHtml.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
Set-ModernButtonStyle -Button $btnExportHtml -BackColor ([System.Drawing.Color]::FromArgb(34, 139, 34))
$flowButtons.Controls.Add($btnExportHtml)

# ---------------------------------------------------------------------------
# Header panel (Dock:Top)
# ---------------------------------------------------------------------------

$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlHeader.Height = 60
$pnlHeader.BackColor = $clrAccent
$pnlHeader.Padding = New-Object System.Windows.Forms.Padding(16, 0, 16, 0)
$form.Controls.Add($pnlHeader)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "DP Content Manager"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 17, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.AutoSize = $true
$lblTitle.BackColor = [System.Drawing.Color]::Transparent
$lblTitle.Location = New-Object System.Drawing.Point(16, 8)
$pnlHeader.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = "MECM Distribution Point Content Status"
$lblSubtitle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblSubtitle.ForeColor = $clrSubtitle
$lblSubtitle.AutoSize = $true
$lblSubtitle.BackColor = [System.Drawing.Color]::Transparent
$lblSubtitle.Location = New-Object System.Drawing.Point(18, 36)
$pnlHeader.Controls.Add($lblSubtitle)

# ---------------------------------------------------------------------------
# Connection bar (Dock:Top)
# ---------------------------------------------------------------------------

$pnlConnBar = New-Object System.Windows.Forms.Panel
$pnlConnBar.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlConnBar.Height = 36
$pnlConnBar.BackColor = $clrPanelBg
$pnlConnBar.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)
$form.Controls.Add($pnlConnBar)

$flowConn = New-Object System.Windows.Forms.FlowLayoutPanel
$flowConn.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowConn.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$flowConn.WrapContents = $false
$flowConn.BackColor = $clrPanelBg
$pnlConnBar.Controls.Add($flowConn)

$lblSiteLabel = New-Object System.Windows.Forms.Label
$lblSiteLabel.Text = "Site:"
$lblSiteLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblSiteLabel.AutoSize = $true
$lblSiteLabel.Margin = New-Object System.Windows.Forms.Padding(0, 3, 2, 0)
$lblSiteLabel.ForeColor = $clrText
$lblSiteLabel.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblSiteLabel)

$lblSiteVal = New-Object System.Windows.Forms.Label
$lblSiteVal.Text = if ($script:Prefs.SiteCode) { $script:Prefs.SiteCode } else { '(not set)' }
$lblSiteVal.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblSiteVal.AutoSize = $true
$lblSiteVal.Margin = New-Object System.Windows.Forms.Padding(0, 3, 16, 0)
$lblSiteVal.ForeColor = if ($script:Prefs.SiteCode) { $clrAccent } else { $clrHint }
$lblSiteVal.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblSiteVal)

$lblServerLabel = New-Object System.Windows.Forms.Label
$lblServerLabel.Text = "Server:"
$lblServerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblServerLabel.AutoSize = $true
$lblServerLabel.Margin = New-Object System.Windows.Forms.Padding(0, 3, 2, 0)
$lblServerLabel.ForeColor = $clrText
$lblServerLabel.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblServerLabel)

$lblServerVal = New-Object System.Windows.Forms.Label
$lblServerVal.Text = if ($script:Prefs.SMSProvider) { $script:Prefs.SMSProvider } else { '(not set)' }
$lblServerVal.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblServerVal.AutoSize = $true
$lblServerVal.Margin = New-Object System.Windows.Forms.Padding(0, 3, 16, 0)
$lblServerVal.ForeColor = if ($script:Prefs.SMSProvider) { $clrAccent } else { $clrHint }
$lblServerVal.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblServerVal)

$lblConnStatus = New-Object System.Windows.Forms.Label
$lblConnStatus.Text = "Disconnected"
$lblConnStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblConnStatus.AutoSize = $true
$lblConnStatus.Margin = New-Object System.Windows.Forms.Padding(0, 3, 20, 0)
$lblConnStatus.ForeColor = $clrHint
$lblConnStatus.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblConnStatus)

$btnRefreshAll = New-Object System.Windows.Forms.Button
$btnRefreshAll.Text = "Refresh All"
$btnRefreshAll.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnRefreshAll.Size = New-Object System.Drawing.Size(100, 24)
$btnRefreshAll.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
Set-ModernButtonStyle -Button $btnRefreshAll -BackColor $clrAccent
$flowConn.Controls.Add($btnRefreshAll)

# Separator below connection bar
$pnlSep1 = New-Object System.Windows.Forms.Panel
$pnlSep1.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSep1.Height = 1
$pnlSep1.BackColor = $clrSepLine
$form.Controls.Add($pnlSep1)

# ---------------------------------------------------------------------------
# Filter bar (Dock:Top)
# ---------------------------------------------------------------------------

$pnlFilter = New-Object System.Windows.Forms.Panel
$pnlFilter.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlFilter.Height = 44
$pnlFilter.BackColor = $clrPanelBg
$pnlFilter.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)
$form.Controls.Add($pnlFilter)

$flowFilter = New-Object System.Windows.Forms.FlowLayoutPanel
$flowFilter.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowFilter.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$flowFilter.WrapContents = $false
$flowFilter.BackColor = $clrPanelBg
$pnlFilter.Controls.Add($flowFilter)

$lblTypeFilter = New-Object System.Windows.Forms.Label
$lblTypeFilter.Text = "Type:"
$lblTypeFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblTypeFilter.AutoSize = $true
$lblTypeFilter.Margin = New-Object System.Windows.Forms.Padding(0, 6, 4, 0)
$lblTypeFilter.ForeColor = $clrText
$lblTypeFilter.BackColor = $clrPanelBg
$flowFilter.Controls.Add($lblTypeFilter)

$cboType = New-Object System.Windows.Forms.ComboBox
$cboType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboType.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cboType.Width = 150
$cboType.Margin = New-Object System.Windows.Forms.Padding(0, 2, 16, 0)
$cboType.BackColor = $clrDetailBg
$cboType.ForeColor = $clrText
$cboType.FlatStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.FlatStyle]::Flat } else { [System.Windows.Forms.FlatStyle]::Standard }
[void]$cboType.Items.AddRange(@('All', 'Application', 'Package', 'SU Deployment Pkg', 'Boot Image', 'OS Image', 'Driver Package', 'Task Sequence'))
$cboType.SelectedIndex = 0
$flowFilter.Controls.Add($cboType)

$lblStatusFilter = New-Object System.Windows.Forms.Label
$lblStatusFilter.Text = "Status:"
$lblStatusFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblStatusFilter.AutoSize = $true
$lblStatusFilter.Margin = New-Object System.Windows.Forms.Padding(0, 6, 4, 0)
$lblStatusFilter.ForeColor = $clrText
$lblStatusFilter.BackColor = $clrPanelBg
$flowFilter.Controls.Add($lblStatusFilter)

$cboStatus = New-Object System.Windows.Forms.ComboBox
$cboStatus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cboStatus.Width = 120
$cboStatus.Margin = New-Object System.Windows.Forms.Padding(0, 2, 16, 0)
$cboStatus.BackColor = $clrDetailBg
$cboStatus.ForeColor = $clrText
$cboStatus.FlatStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.FlatStyle]::Flat } else { [System.Windows.Forms.FlatStyle]::Standard }
[void]$cboStatus.Items.AddRange(@('All', 'Failed Only', 'In Progress', 'Installed'))
$cboStatus.SelectedIndex = 0
$flowFilter.Controls.Add($cboStatus)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Filter:"
$lblFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblFilter.AutoSize = $true
$lblFilter.Margin = New-Object System.Windows.Forms.Padding(0, 6, 4, 0)
$lblFilter.ForeColor = $clrText
$lblFilter.BackColor = $clrPanelBg
$flowFilter.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$txtFilter.Width = 260
$txtFilter.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
$txtFilter.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
$txtFilter.BackColor = $clrDetailBg
$txtFilter.ForeColor = $clrText
$flowFilter.Controls.Add($txtFilter)

# Separator below filter bar
$pnlSep2 = New-Object System.Windows.Forms.Panel
$pnlSep2.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSep2.Height = 1
$pnlSep2.BackColor = $clrSepLine
$form.Controls.Add($pnlSep2)

# ---------------------------------------------------------------------------
# Helper: Create a themed DataGridView
# ---------------------------------------------------------------------------

function New-ThemedGrid {
    param([switch]$MultiSelect)

    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = [System.Windows.Forms.DockStyle]::Fill
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.AllowUserToResizeRows = $false
    $g.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $g.MultiSelect = [bool]$MultiSelect
    $g.AutoGenerateColumns = $false
    $g.RowHeadersVisible = $false
    $g.BackgroundColor = $clrPanelBg
    $g.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $g.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $g.GridColor = $clrGridLine
    $g.ColumnHeadersDefaultCellStyle.BackColor = $clrAccent
    $g.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $g.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $g.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(4)
    $g.ColumnHeadersDefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
    $g.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $g.ColumnHeadersHeight = 32
    $g.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    $g.EnableHeadersVisualStyles = $false
    $g.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $g.DefaultCellStyle.ForeColor = $clrGridText
    $g.DefaultCellStyle.BackColor = $clrPanelBg
    $g.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(2)
    $g.DefaultCellStyle.SelectionBackColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(38, 79, 120) } else { [System.Drawing.Color]::FromArgb(0, 120, 215) }
    $g.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $g.RowTemplate.Height = 26
    $g.AlternatingRowsDefaultCellStyle.BackColor = $clrGridAlt

    Enable-DoubleBuffer -Control $g
    return $g
}

# ---------------------------------------------------------------------------
# TabControl (Fill)
# ---------------------------------------------------------------------------

$tabMain = New-Object System.Windows.Forms.TabControl
$tabMain.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabMain.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

if ($script:Prefs.DarkMode) {
    $tabMain.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
    $tabMain.Add_DrawItem({
        param($s, $e)
        $tab = $s.TabPages[$e.Index]
        $isSelected = ($s.SelectedIndex -eq $e.Index)
        $bgColor = if ($isSelected) { $clrAccent } else { $clrPanelBg }
        $fgColor = if ($isSelected) { [System.Drawing.Color]::White } else { $clrText }
        $brush = New-Object System.Drawing.SolidBrush($bgColor)
        $e.Graphics.FillRectangle($brush, $e.Bounds)
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = [System.Drawing.StringAlignment]::Center
        $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
        $textBrush = New-Object System.Drawing.SolidBrush($fgColor)
        $e.Graphics.DrawString($tab.Text, $s.Font, $textBrush, [System.Drawing.RectangleF]$e.Bounds, $sf)
        $brush.Dispose(); $textBrush.Dispose(); $sf.Dispose()
    })
}

$form.Controls.Add($tabMain)

# ===================== TAB 0: Content View =====================

$tabContent = New-Object System.Windows.Forms.TabPage
$tabContent.Text = "Content View"
$tabContent.BackColor = $clrFormBg
$tabMain.TabPages.Add($tabContent)

$splitContent = New-Object System.Windows.Forms.SplitContainer
$splitContent.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContent.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitContent.SplitterDistance = 400
$splitContent.SplitterWidth = 6
$splitContent.BackColor = $clrSepLine
$splitContent.Panel1.BackColor = $clrPanelBg
$splitContent.Panel2.BackColor = $clrPanelBg
$splitContent.Panel1MinSize = 100
$splitContent.Panel2MinSize = 80
$tabContent.Controls.Add($splitContent)

# -- Content main grid (Panel1)
$gridContent = New-ThemedGrid -MultiSelect

$colCName     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCName.HeaderText = "Content Name";    $colCName.DataPropertyName = "ContentName";     $colCName.Width = 250
$colCType     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCType.HeaderText = "Type";            $colCType.DataPropertyName = "ContentType";     $colCType.Width = 120
$colCPkgId    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCPkgId.HeaderText = "Package ID";     $colCPkgId.DataPropertyName = "PackageID";      $colCPkgId.Width = 90
$colCSizeMB   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCSizeMB.HeaderText = "Size (MB)";     $colCSizeMB.DataPropertyName = "SourceSizeMB";  $colCSizeMB.Width = 80;  $colCSizeMB.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleRight
$colCTotalDP  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCTotalDP.HeaderText = "Total DPs";    $colCTotalDP.DataPropertyName = "TotalDPs";     $colCTotalDP.Width = 70
$colCInstall  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCInstall.HeaderText = "Installed";    $colCInstall.DataPropertyName = "InstalledCount"; $colCInstall.Width = 70
$colCInProg   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCInProg.HeaderText = "In Progress";   $colCInProg.DataPropertyName = "InProgressCount"; $colCInProg.Width = 80
$colCFailed   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCFailed.HeaderText = "Failed";        $colCFailed.DataPropertyName = "FailedCount";   $colCFailed.Width = 60
$colCPct      = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCPct.HeaderText = "% Complete";       $colCPct.DataPropertyName = "PctComplete";      $colCPct.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

$gridContent.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colCName, $colCType, $colCPkgId, $colCSizeMB, $colCTotalDP, $colCInstall, $colCInProg, $colCFailed, $colCPct))
$splitContent.Panel1.Controls.Add($gridContent)

# Content DataTable
$dtContent = New-Object System.Data.DataTable
[void]$dtContent.Columns.Add("ContentName", [string])
[void]$dtContent.Columns.Add("ContentType", [string])
[void]$dtContent.Columns.Add("PackageID", [string])
[void]$dtContent.Columns.Add("SourceSizeMB", [double])
[void]$dtContent.Columns.Add("TotalDPs", [int])
[void]$dtContent.Columns.Add("InstalledCount", [int])
[void]$dtContent.Columns.Add("InProgressCount", [int])
[void]$dtContent.Columns.Add("FailedCount", [int])
[void]$dtContent.Columns.Add("PctComplete", [string])

$gridContent.DataSource = $dtContent

# Content grid row color coding
$gridContent.Add_RowPrePaint({
    param($s, $e)
    try {
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $dtContent.DefaultView.Count) {
            $rowView = $dtContent.DefaultView[$e.RowIndex]
            $failed = [int]$rowView["FailedCount"]
            $inProg = [int]$rowView["InProgressCount"]
            if ($failed -gt 0) {
                $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrErrText
            }
            elseif ($inProg -gt 0) {
                $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrWarnText
            }
            else {
                $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrGridText
            }
        }
    } catch {}
})

# -- Content detail panel (Panel2)
$pnlContentDetail = New-Object System.Windows.Forms.Panel
$pnlContentDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlContentDetail.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 8)
$pnlContentDetail.BackColor = $clrPanelBg
$splitContent.Panel2.Controls.Add($pnlContentDetail)

$lblContentDetailTitle = New-Object System.Windows.Forms.Label
$lblContentDetailTitle.Text = "Content Detail"
$lblContentDetailTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblContentDetailTitle.Dock = [System.Windows.Forms.DockStyle]::Top
$lblContentDetailTitle.Height = 20
$lblContentDetailTitle.ForeColor = $clrHint
$lblContentDetailTitle.BackColor = $clrPanelBg
$pnlContentDetail.Controls.Add($lblContentDetailTitle)

# Split detail into info text (top) and per-DP sub-grid (bottom)
$splitContentDetail = New-Object System.Windows.Forms.SplitContainer
$splitContentDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitContentDetail.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitContentDetail.SplitterDistance = 80
$splitContentDetail.SplitterWidth = 4
$splitContentDetail.BackColor = $clrSepLine
$splitContentDetail.Panel1.BackColor = $clrPanelBg
$splitContentDetail.Panel2.BackColor = $clrPanelBg
$pnlContentDetail.Controls.Add($splitContentDetail)
$splitContentDetail.BringToFront()

$txtContentInfo = New-Object System.Windows.Forms.RichTextBox
$txtContentInfo.ReadOnly = $true
$txtContentInfo.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtContentInfo.BackColor = $clrDetailBg
$txtContentInfo.ForeColor = $clrText
$txtContentInfo.WordWrap = $true
$txtContentInfo.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtContentInfo.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$txtContentInfo.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$splitContentDetail.Panel1.Controls.Add($txtContentInfo)

# Per-DP sub-grid for selected content
$gridContentDPs = New-ThemedGrid
$gridContentDPs.ColumnHeadersHeight = 28
$gridContentDPs.RowTemplate.Height = 24

$colCDDP     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCDDP.HeaderText = "DP Name";          $colCDDP.DataPropertyName = "DPName";        $colCDDP.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$colCDSite   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCDSite.HeaderText = "Site";           $colCDSite.DataPropertyName = "SiteCode";    $colCDSite.Width = 60
$colCDStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCDStatus.HeaderText = "Status";       $colCDStatus.DataPropertyName = "Status";    $colCDStatus.Width = 110
$colCDTime   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCDTime.HeaderText = "Last Updated";   $colCDTime.DataPropertyName = "SummaryDate"; $colCDTime.Width = 140
$colCDVer    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colCDVer.HeaderText = "Source Ver";      $colCDVer.DataPropertyName = "SourceVersion"; $colCDVer.Width = 80
$gridContentDPs.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colCDDP, $colCDSite, $colCDStatus, $colCDTime, $colCDVer))

$dtContentDPs = New-Object System.Data.DataTable
[void]$dtContentDPs.Columns.Add("DPName", [string])
[void]$dtContentDPs.Columns.Add("SiteCode", [string])
[void]$dtContentDPs.Columns.Add("Status", [string])
[void]$dtContentDPs.Columns.Add("SummaryDate", [string])
[void]$dtContentDPs.Columns.Add("SourceVersion", [string])
$gridContentDPs.DataSource = $dtContentDPs
$splitContentDetail.Panel2.Controls.Add($gridContentDPs)

# ===================== TAB 1: DP View =====================

$tabDP = New-Object System.Windows.Forms.TabPage
$tabDP.Text = "DP View"
$tabDP.BackColor = $clrFormBg
$tabMain.TabPages.Add($tabDP)

$splitDP = New-Object System.Windows.Forms.SplitContainer
$splitDP.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitDP.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitDP.SplitterDistance = 400
$splitDP.SplitterWidth = 6
$splitDP.BackColor = $clrSepLine
$splitDP.Panel1.BackColor = $clrPanelBg
$splitDP.Panel2.BackColor = $clrPanelBg
$splitDP.Panel1MinSize = 100
$splitDP.Panel2MinSize = 80
$tabDP.Controls.Add($splitDP)

# -- DP main grid (Panel1)
$gridDP = New-ThemedGrid -MultiSelect

$colDName    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDName.HeaderText = "DP Name";        $colDName.DataPropertyName = "DPName";          $colDName.Width = 280
$colDSite    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDSite.HeaderText = "Site";           $colDSite.DataPropertyName = "SiteCode";        $colDSite.Width = 60
$colDPull    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDPull.HeaderText = "Pull DP";        $colDPull.DataPropertyName = "IsPullDP";        $colDPull.Width = 60
$colDTotal   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDTotal.HeaderText = "Total Content"; $colDTotal.DataPropertyName = "TotalContent";   $colDTotal.Width = 90
$colDInstall = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDInstall.HeaderText = "Installed";   $colDInstall.DataPropertyName = "InstalledCount"; $colDInstall.Width = 70
$colDInProg  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDInProg.HeaderText = "In Progress";  $colDInProg.DataPropertyName = "InProgressCount"; $colDInProg.Width = 80
$colDFailed  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDFailed.HeaderText = "Failed";       $colDFailed.DataPropertyName = "FailedCount";   $colDFailed.Width = 60
$colDSize    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDSize.HeaderText = "Size (GB)";      $colDSize.DataPropertyName = "TotalSizeGB";     $colDSize.Width = 80; $colDSize.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleRight
$colDPct     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDPct.HeaderText = "% Complete";      $colDPct.DataPropertyName = "PctComplete";      $colDPct.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

$gridDP.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colDName, $colDSite, $colDPull, $colDTotal, $colDInstall, $colDInProg, $colDFailed, $colDSize, $colDPct))
$splitDP.Panel1.Controls.Add($gridDP)

# DP DataTable
$dtDP = New-Object System.Data.DataTable
[void]$dtDP.Columns.Add("DPName", [string])
[void]$dtDP.Columns.Add("SiteCode", [string])
[void]$dtDP.Columns.Add("IsPullDP", [string])
[void]$dtDP.Columns.Add("TotalContent", [int])
[void]$dtDP.Columns.Add("InstalledCount", [int])
[void]$dtDP.Columns.Add("InProgressCount", [int])
[void]$dtDP.Columns.Add("FailedCount", [int])
[void]$dtDP.Columns.Add("TotalSizeGB", [double])
[void]$dtDP.Columns.Add("PctComplete", [string])

$gridDP.DataSource = $dtDP

# DP grid row color coding
$gridDP.Add_RowPrePaint({
    param($s, $e)
    try {
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $dtDP.DefaultView.Count) {
            $rowView = $dtDP.DefaultView[$e.RowIndex]
            $failed = [int]$rowView["FailedCount"]
            $inProg = [int]$rowView["InProgressCount"]
            if ($failed -gt 0) {
                $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrErrText
            }
            elseif ($inProg -gt 0) {
                $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrWarnText
            }
            else {
                $s.Rows[$e.RowIndex].DefaultCellStyle.ForeColor = $clrGridText
            }
        }
    } catch {}
})

# -- DP detail panel (Panel2)
$pnlDPDetail = New-Object System.Windows.Forms.Panel
$pnlDPDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlDPDetail.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 8)
$pnlDPDetail.BackColor = $clrPanelBg
$splitDP.Panel2.Controls.Add($pnlDPDetail)

$lblDPDetailTitle = New-Object System.Windows.Forms.Label
$lblDPDetailTitle.Text = "DP Detail"
$lblDPDetailTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblDPDetailTitle.Dock = [System.Windows.Forms.DockStyle]::Top
$lblDPDetailTitle.Height = 20
$lblDPDetailTitle.ForeColor = $clrHint
$lblDPDetailTitle.BackColor = $clrPanelBg
$pnlDPDetail.Controls.Add($lblDPDetailTitle)

$splitDPDetail = New-Object System.Windows.Forms.SplitContainer
$splitDPDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitDPDetail.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitDPDetail.SplitterDistance = 80
$splitDPDetail.SplitterWidth = 4
$splitDPDetail.BackColor = $clrSepLine
$splitDPDetail.Panel1.BackColor = $clrPanelBg
$splitDPDetail.Panel2.BackColor = $clrPanelBg
$pnlDPDetail.Controls.Add($splitDPDetail)
$splitDPDetail.BringToFront()

$txtDPInfo = New-Object System.Windows.Forms.RichTextBox
$txtDPInfo.ReadOnly = $true
$txtDPInfo.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtDPInfo.BackColor = $clrDetailBg
$txtDPInfo.ForeColor = $clrText
$txtDPInfo.WordWrap = $true
$txtDPInfo.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtDPInfo.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$txtDPInfo.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$splitDPDetail.Panel1.Controls.Add($txtDPInfo)

# Content sub-grid for selected DP
$gridDPContent = New-ThemedGrid
$gridDPContent.ColumnHeadersHeight = 28
$gridDPContent.RowTemplate.Height = 24

$colDCName   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDCName.HeaderText = "Content Name";   $colDCName.DataPropertyName = "ContentName";   $colDCName.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$colDCType   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDCType.HeaderText = "Type";           $colDCType.DataPropertyName = "ContentType";   $colDCType.Width = 110
$colDCPkgId  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDCPkgId.HeaderText = "Package ID";    $colDCPkgId.DataPropertyName = "PackageID";    $colDCPkgId.Width = 90
$colDCStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDCStatus.HeaderText = "Status";       $colDCStatus.DataPropertyName = "Status";      $colDCStatus.Width = 110
$colDCSizeMB = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDCSizeMB.HeaderText = "Size (MB)";    $colDCSizeMB.DataPropertyName = "SizeMB";      $colDCSizeMB.Width = 80; $colDCSizeMB.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleRight
$colDCTime   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDCTime.HeaderText = "Last Updated";   $colDCTime.DataPropertyName = "SummaryDate";   $colDCTime.Width = 140
$gridDPContent.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colDCName, $colDCType, $colDCPkgId, $colDCStatus, $colDCSizeMB, $colDCTime))

$dtDPContent = New-Object System.Data.DataTable
[void]$dtDPContent.Columns.Add("ContentName", [string])
[void]$dtDPContent.Columns.Add("ContentType", [string])
[void]$dtDPContent.Columns.Add("PackageID", [string])
[void]$dtDPContent.Columns.Add("Status", [string])
[void]$dtDPContent.Columns.Add("SizeMB", [double])
[void]$dtDPContent.Columns.Add("SummaryDate", [string])
$gridDPContent.DataSource = $dtDPContent
$splitDPDetail.Panel2.Controls.Add($gridDPContent)

# ---------------------------------------------------------------------------
# Finalize dock Z-order
# ---------------------------------------------------------------------------

$form.Controls.Add($menuStrip)
$menuStrip.SendToBack()

$pnlSep2.BringToFront()
$pnlFilter.BringToFront()
$pnlSep1.BringToFront()
$pnlConnBar.BringToFront()
$pnlHeader.BringToFront()

$tabMain.BringToFront()

# ---------------------------------------------------------------------------
# Module-scoped data (populated by Refresh All)
# ---------------------------------------------------------------------------

$script:AllDPs      = @()
$script:AllContent  = @()
$script:AllStatus   = @()
$script:ContentLookup = @{}
$script:DPLookup      = @{}

# ---------------------------------------------------------------------------
# Core operations
# ---------------------------------------------------------------------------

function Invoke-RefreshAll {
    if (-not $script:Prefs.SiteCode -or -not $script:Prefs.SMSProvider) {
        [System.Windows.Forms.MessageBox]::Show(
            "Site Code and SMS Provider must be configured in File > Preferences.",
            "Configuration Required", "OK", "Warning") | Out-Null
        return
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnRefreshAll.Enabled = $false

    try {
        # Connect if needed
        if (-not (Test-CMConnection)) {
            Add-LogLine -TextBox $txtLog -Message "Connecting to site $($script:Prefs.SiteCode)..."
            [System.Windows.Forms.Application]::DoEvents()

            $connected = Connect-CMSite -SiteCode $script:Prefs.SiteCode -SMSProvider $script:Prefs.SMSProvider
            if (-not $connected) {
                Add-LogLine -TextBox $txtLog -Message "Connection failed. Check preferences and try again."
                $statusLabel.Text = "Connection failed."
                return
            }
            $lblConnStatus.Text = "Connected"
            $lblConnStatus.ForeColor = $clrOkText
        }

        # Step 1: Load DPs
        Add-LogLine -TextBox $txtLog -Message "Loading distribution points..."
        [System.Windows.Forms.Application]::DoEvents()
        $script:AllDPs = @(Get-AllDistributionPoints)
        Add-LogLine -TextBox $txtLog -Message "Found $($script:AllDPs.Count) DPs"
        [System.Windows.Forms.Application]::DoEvents()

        # Build DP lookup
        $script:DPLookup = @{}
        foreach ($dp in $script:AllDPs) { $script:DPLookup[$dp.Name] = $dp }

        # Step 2: Load content objects
        Add-LogLine -TextBox $txtLog -Message "Loading content objects (all types)..."
        [System.Windows.Forms.Application]::DoEvents()
        $script:AllContent = @(Get-AllContentObjects)
        Add-LogLine -TextBox $txtLog -Message "Loaded $($script:AllContent.Count) content objects"
        [System.Windows.Forms.Application]::DoEvents()

        # Build content lookup
        $script:ContentLookup = @{}
        foreach ($obj in $script:AllContent) { $script:ContentLookup[$obj.PackageID] = $obj }

        # Step 3: Bulk status query
        Add-LogLine -TextBox $txtLog -Message "Running bulk distribution status query (this may take a minute)..."
        [System.Windows.Forms.Application]::DoEvents()
        $script:AllStatus = @(Get-BulkDistributionStatus -SMSProvider $script:Prefs.SMSProvider -SiteCode $script:Prefs.SiteCode)
        Add-LogLine -TextBox $txtLog -Message "Retrieved $($script:AllStatus.Count) status records"
        [System.Windows.Forms.Application]::DoEvents()

        # Step 4: Aggregate and populate Content View
        Add-LogLine -TextBox $txtLog -Message "Building content summary..."
        $contentSummary = ConvertTo-ContentStatusSummary -StatusRows $script:AllStatus

        $dtContent.Clear()
        foreach ($cs in $contentSummary) {
            $obj = $script:ContentLookup[$cs.PackageID]
            $name = if ($obj) { $obj.ContentName } else { $cs.PackageID }
            $type = if ($obj) { $obj.ContentType } else { 'Unknown' }
            $sizeMB = if ($obj -and $obj.SourceSize) { [math]::Round($obj.SourceSize / 1024, 1) } else { 0 }

            [void]$dtContent.Rows.Add(
                $name,
                $type,
                $cs.PackageID,
                $sizeMB,
                $cs.TotalDPs,
                $cs.InstalledCount,
                $cs.InProgressCount,
                $cs.FailedCount,
                $cs.PctComplete
            )
        }
        [System.Windows.Forms.Application]::DoEvents()

        # Step 5: Populate DP View
        Add-LogLine -TextBox $txtLog -Message "Building DP summary..."
        $dpSummary = ConvertTo-DPStatusSummary -StatusRows $script:AllStatus

        # Storage analysis for size column
        $storageData = Get-DPStorageAnalysis -StatusRows $script:AllStatus -ContentObjects $script:AllContent
        $storageLookup = @{}
        foreach ($sd in $storageData) { $storageLookup[$sd.DPName] = $sd }

        $dtDP.Clear()
        foreach ($ds in $dpSummary) {
            $dpObj = $script:DPLookup[$ds.DPName]
            $siteCode = if ($dpObj) { $dpObj.SiteCode } else { '' }
            $isPullDP = if ($dpObj) { if ($dpObj.IsPullDP) { 'Yes' } else { 'No' } } else { '' }
            $sizeGB = if ($storageLookup[$ds.DPName]) { $storageLookup[$ds.DPName].TotalSizeGB } else { 0 }

            [void]$dtDP.Rows.Add(
                $ds.DPName,
                $siteCode,
                $isPullDP,
                $ds.TotalContent,
                $ds.InstalledCount,
                $ds.InProgressCount,
                $ds.FailedCount,
                $sizeGB,
                $ds.PctComplete
            )
        }
        [System.Windows.Forms.Application]::DoEvents()

        # Update status bar
        $totalFailed = ($contentSummary | Measure-Object -Property FailedCount -Sum).Sum
        $statusLabel.Text = "Connected to {0} | {1} DPs | {2} content objects | {3} failed | Last refresh: {4}" -f `
            $script:Prefs.SiteCode, $script:AllDPs.Count, $script:AllContent.Count, $totalFailed, (Get-Date -Format 'HH:mm:ss')

        Add-LogLine -TextBox $txtLog -Message "Refresh complete."

    }
    catch {
        Add-LogLine -TextBox $txtLog -Message "ERROR: $_"
        $statusLabel.Text = "Error during refresh."
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnRefreshAll.Enabled = $true
    }
}

# ---------------------------------------------------------------------------
# Content drill-down
# ---------------------------------------------------------------------------

$gridContent.Add_SelectionChanged({
    if ($gridContent.SelectedRows.Count -eq 0) { $txtContentInfo.Text = ''; $dtContentDPs.Clear(); return }
    if (-not $script:AllStatus) { return }

    $rowIdx = $gridContent.SelectedRows[0].Index
    if ($rowIdx -lt 0 -or $rowIdx -ge $dtContent.DefaultView.Count) { return }

    $row = $dtContent.DefaultView[$rowIdx]
    $pkgId = [string]$row["PackageID"]

    # Info text
    $lines = @()
    $lines += "Content:      $($row['ContentName'])"
    $lines += "Type:         $($row['ContentType'])"
    $lines += "Package ID:   $pkgId"
    $lines += "Size:         $($row['SourceSizeMB']) MB"
    $lines += "Status:       $($row['InstalledCount'])/$($row['TotalDPs']) DPs installed ($($row['PctComplete']))"
    if ([int]$row['FailedCount'] -gt 0) {
        $lines += "FAILED:       $($row['FailedCount']) DPs"
    }
    $txtContentInfo.Text = ($lines -join "`r`n")

    # Per-DP sub-grid
    $dtContentDPs.Clear()
    $matchingRows = $script:AllStatus | Where-Object { $_.PackageID -eq $pkgId }
    foreach ($sr in $matchingRows) {
        $dpObj = $script:DPLookup[$sr.DPName]
        $site = if ($dpObj) { $dpObj.SiteCode } else { '' }
        $dateStr = if ($sr.SummaryDate) { $sr.SummaryDate.ToString('yyyy-MM-dd HH:mm') } else { '' }
        [void]$dtContentDPs.Rows.Add($sr.DPName, $site, $sr.StatusName, $dateStr, $sr.SourceVersion)
    }
})

# ---------------------------------------------------------------------------
# DP drill-down
# ---------------------------------------------------------------------------

$gridDP.Add_SelectionChanged({
    if ($gridDP.SelectedRows.Count -eq 0) { $txtDPInfo.Text = ''; $dtDPContent.Clear(); return }
    if (-not $script:AllStatus) { return }

    $rowIdx = $gridDP.SelectedRows[0].Index
    if ($rowIdx -lt 0 -or $rowIdx -ge $dtDP.DefaultView.Count) { return }

    $row = $dtDP.DefaultView[$rowIdx]
    $dpName = [string]$row["DPName"]

    # Info text
    $lines = @()
    $lines += "DP:           $dpName"
    $lines += "Site:         $($row['SiteCode'])"
    $lines += "Pull DP:      $($row['IsPullDP'])"
    $lines += "Content:      $($row['TotalContent']) objects"
    $lines += "Installed:    $($row['InstalledCount']) ($($row['PctComplete']))"
    $lines += "Size:         $($row['TotalSizeGB']) GB"
    if ([int]$row['FailedCount'] -gt 0) {
        $lines += "FAILED:       $($row['FailedCount']) items"
    }
    $txtDPInfo.Text = ($lines -join "`r`n")

    # Content sub-grid
    $dtDPContent.Clear()
    $matchingRows = $script:AllStatus | Where-Object { $_.DPName -eq $dpName }
    foreach ($sr in $matchingRows) {
        $obj = $script:ContentLookup[$sr.PackageID]
        $name = if ($obj) { $obj.ContentName } else { $sr.PackageID }
        $type = if ($obj) { $obj.ContentType } else { 'Unknown' }
        $sizeMB = if ($obj -and $obj.SourceSize) { [math]::Round($obj.SourceSize / 1024, 1) } else { 0 }
        $dateStr = if ($sr.SummaryDate) { $sr.SummaryDate.ToString('yyyy-MM-dd HH:mm') } else { '' }
        [void]$dtDPContent.Rows.Add($name, $type, $sr.PackageID, $sr.StatusName, $sizeMB, $dateStr)
    }
})

# ---------------------------------------------------------------------------
# Filter handlers
# ---------------------------------------------------------------------------

function Apply-ContentFilter {
    $parts = @()

    $typeFilter = $cboType.SelectedItem
    if ($typeFilter -and $typeFilter -ne 'All') {
        $parts += "ContentType = '$typeFilter'"
    }

    $statusFilter = $cboStatus.SelectedItem
    switch ($statusFilter) {
        'Failed Only' { $parts += "FailedCount > 0" }
        'In Progress' { $parts += "InProgressCount > 0" }
        'Installed'   { $parts += "FailedCount = 0 AND InProgressCount = 0" }
    }

    $text = $txtFilter.Text.Trim()
    if ($text) {
        $escaped = $text.Replace("'", "''")
        $parts += "(ContentName LIKE '*$escaped*' OR PackageID LIKE '*$escaped*')"
    }

    $dtContent.DefaultView.RowFilter = ($parts -join ' AND ')
}

function Apply-DPFilter {
    $text = $txtFilter.Text.Trim()
    if ($text) {
        $escaped = $text.Replace("'", "''")
        $dtDP.DefaultView.RowFilter = "DPName LIKE '*$escaped*'"
    } else {
        $dtDP.DefaultView.RowFilter = ''
    }
}

$cboType.Add_SelectedIndexChanged({ Apply-ContentFilter })
$cboStatus.Add_SelectedIndexChanged({ Apply-ContentFilter })
$txtFilter.Add_TextChanged({
    if ($tabMain.SelectedIndex -eq 0) { Apply-ContentFilter }
    else { Apply-DPFilter }
})

# ---------------------------------------------------------------------------
# Action handlers
# ---------------------------------------------------------------------------

function Get-SelectedContentPkgIds {
    $ids = @()
    foreach ($row in $gridContent.SelectedRows) {
        $idx = $row.Index
        if ($idx -ge 0 -and $idx -lt $dtContent.DefaultView.Count) {
            $ids += [string]$dtContent.DefaultView[$idx]["PackageID"]
        }
    }
    return $ids
}

function Invoke-RedistributeSelected {
    if (-not $script:AllStatus) {
        [System.Windows.Forms.MessageBox]::Show("Load data first.", "No Data", "OK", "Information") | Out-Null
        return
    }

    $pkgIds = Get-SelectedContentPkgIds
    if ($pkgIds.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select one or more content items in the Content View.", "No Selection", "OK", "Information") | Out-Null
        return
    }

    # Find failed DPs for each selected package
    $pairs = @()
    foreach ($pkgId in $pkgIds) {
        $failedDPs = @($script:AllStatus | Where-Object { $_.PackageID -eq $pkgId -and $_.State -in 3, 6 } | ForEach-Object { $_.DPName })
        foreach ($dp in $failedDPs) {
            $pairs += [PSCustomObject]@{ PackageID = $pkgId; DPName = $dp }
        }
    }

    if ($pairs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No failed DPs found for the selected content.", "Nothing to Redistribute", "OK", "Information") | Out-Null
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Redistribute $($pkgIds.Count) content item(s) to $($pairs.Count) failed DP(s)?",
        "Confirm Redistribute",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $successCount = 0
    $failCount = 0

    foreach ($pair in $pairs) {
        Add-LogLine -TextBox $txtLog -Message "Redistributing $($pair.PackageID) to $($pair.DPName)..."
        [System.Windows.Forms.Application]::DoEvents()

        try {
            Start-CMContentDistribution -PackageId $pair.PackageID -DistributionPointName $pair.DPName -ErrorAction Stop
            $successCount++
        }
        catch {
            Add-LogLine -TextBox $txtLog -Message "  FAILED: $_"
            $failCount++
        }
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    Add-LogLine -TextBox $txtLog -Message "Redistribution complete. $successCount succeeded, $failCount failed."

    $refresh = [System.Windows.Forms.MessageBox]::Show("Refresh status now?", "Refresh", "YesNo", "Question")
    if ($refresh -eq [System.Windows.Forms.DialogResult]::Yes) { Invoke-RefreshAll }
}

function Invoke-ValidateSelected {
    if (-not $script:AllStatus) {
        [System.Windows.Forms.MessageBox]::Show("Load data first.", "No Data", "OK", "Information") | Out-Null
        return
    }

    $pkgIds = Get-SelectedContentPkgIds
    if ($pkgIds.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select one or more content items in the Content View.", "No Selection", "OK", "Information") | Out-Null
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Validate content integrity for $($pkgIds.Count) item(s)?`nValidation is server-side and results appear on next refresh.",
        "Confirm Validate",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $count = 0

    foreach ($pkgId in $pkgIds) {
        $dpNames = @($script:AllStatus | Where-Object { $_.PackageID -eq $pkgId } | ForEach-Object { $_.DPName } | Select-Object -Unique)
        foreach ($dp in $dpNames) {
            try {
                Invoke-CMContentValidation -PackageId $pkgId -DistributionPointName $dp -ErrorAction Stop
                $count++
            } catch { }
        }
        Add-LogLine -TextBox $txtLog -Message "Validation initiated for $pkgId on $($dpNames.Count) DPs"
        [System.Windows.Forms.Application]::DoEvents()
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    Add-LogLine -TextBox $txtLog -Message "Validation initiated for $count total content-DP pairs. Results appear on next refresh."
}

function Invoke-RemoveOrphaned {
    if (-not $script:AllStatus -or -not $script:AllContent) {
        [System.Windows.Forms.MessageBox]::Show("Load data first.", "No Data", "OK", "Information") | Out-Null
        return
    }

    Add-LogLine -TextBox $txtLog -Message "Scanning for orphaned content..."
    [System.Windows.Forms.Application]::DoEvents()

    $orphans = Find-OrphanedContent -StatusRows $script:AllStatus -ContentObjects $script:AllContent

    if ($orphans.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No orphaned content found.", "Clean", "OK", "Information") | Out-Null
        Add-LogLine -TextBox $txtLog -Message "No orphaned content found."
        return
    }

    $totalDPs = ($orphans | Measure-Object -Property DPCount -Sum).Sum
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Found $($orphans.Count) orphaned content item(s) across $totalDPs DP(s).`n`nRemove all orphaned content?",
        "Confirm Remove Orphaned",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    foreach ($orphan in $orphans) {
        Add-LogLine -TextBox $txtLog -Message "Removing orphan $($orphan.PackageID) from $($orphan.DPCount) DPs..."
        [System.Windows.Forms.Application]::DoEvents()

        foreach ($dp in $orphan.DPNames) {
            try {
                Remove-CMContentDistribution -PackageId $orphan.PackageID -DistributionPointName $dp -Force -ErrorAction Stop
            } catch {
                Add-LogLine -TextBox $txtLog -Message "  Failed on $dp : $_"
            }
        }
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    Add-LogLine -TextBox $txtLog -Message "Orphan removal complete."

    $refresh = [System.Windows.Forms.MessageBox]::Show("Refresh status now?", "Refresh", "YesNo", "Question")
    if ($refresh -eq [System.Windows.Forms.DialogResult]::Yes) { Invoke-RefreshAll }
}

# Wire action buttons
$btnRefreshAll.Add_Click({ Invoke-RefreshAll })
$btnRedistribute.Add_Click({ Invoke-RedistributeSelected })
$btnValidate.Add_Click({ Invoke-ValidateSelected })
$btnRemove.Add_Click({
    if (-not $script:AllStatus) { return }
    $pkgIds = Get-SelectedContentPkgIds
    if ($pkgIds.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select content items to remove.", "No Selection", "OK", "Information") | Out-Null
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Remove $($pkgIds.Count) selected content item(s) from ALL their distribution points?`n`nThis is a destructive action.",
        "Confirm Remove",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    foreach ($pkgId in $pkgIds) {
        $dpNames = @($script:AllStatus | Where-Object { $_.PackageID -eq $pkgId } | ForEach-Object { $_.DPName } | Select-Object -Unique)
        Add-LogLine -TextBox $txtLog -Message "Removing $pkgId from $($dpNames.Count) DPs..."
        [System.Windows.Forms.Application]::DoEvents()

        foreach ($dp in $dpNames) {
            try {
                Remove-CMContentDistribution -PackageId $pkgId -DistributionPointName $dp -Force -ErrorAction Stop
            } catch {
                Add-LogLine -TextBox $txtLog -Message "  Failed on $dp : $_"
            }
        }
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    Add-LogLine -TextBox $txtLog -Message "Removal complete."

    $refresh = [System.Windows.Forms.MessageBox]::Show("Refresh status now?", "Refresh", "YesNo", "Question")
    if ($refresh -eq [System.Windows.Forms.DialogResult]::Yes) { Invoke-RefreshAll }
})

# Export buttons
$btnExportCsv.Add_Click({
    $activeTable = if ($tabMain.SelectedIndex -eq 0) { $dtContent } else { $dtDP }
    $tabName = if ($tabMain.SelectedIndex -eq 0) { 'Content' } else { 'DP' }

    $reportsDir = Join-Path $PSScriptRoot "Reports"
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV Files (*.csv)|*.csv"
    $sfd.FileName = "DPContentMgr-$tabName-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-ContentStatusCsv -DataTable $activeTable -OutputPath $sfd.FileName
        Add-LogLine -TextBox $txtLog -Message "Exported CSV: $($sfd.FileName)"
    }
})

$btnExportHtml.Add_Click({
    $activeTable = if ($tabMain.SelectedIndex -eq 0) { $dtContent } else { $dtDP }
    $tabName = if ($tabMain.SelectedIndex -eq 0) { 'Content Status' } else { 'DP Status' }

    $reportsDir = Join-Path $PSScriptRoot "Reports"
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "HTML Files (*.html)|*.html"
    $sfd.FileName = "DPContentMgr-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-ContentStatusHtml -DataTable $activeTable -OutputPath $sfd.FileName -ReportTitle "DP Content Manager - $tabName"
        Add-LogLine -TextBox $txtLog -Message "Exported HTML: $($sfd.FileName)"
    }
})

# ---------------------------------------------------------------------------
# Window state persistence
# ---------------------------------------------------------------------------

$form.Add_Shown({ Restore-WindowState })
$form.Add_FormClosing({
    Save-WindowState
    Disconnect-CMSite
})

# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------

Add-LogLine -TextBox $txtLog -Message "DP Content Manager started. Configure site in File > Preferences, then click Refresh All."

[void]$form.ShowDialog()
$form.Dispose()
