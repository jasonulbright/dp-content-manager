#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }

<#
.SYNOPSIS
    Pester 5.x window test for the compact 1080p density. The main window is opened off screen, in both themes, at
    the work area of the session (1920x1032 on a 1080p screen at 100 percent) and at 1024x720, and measured:
    sidebar buttons, inline buttons, grid rows and headers, the width that the content uses, and the vacant area
    under and beside the content of each view (at most 120 px high and 160 px wide at the work area).
    The run prints the work area of the session (SystemParameters.WorkArea).
#>

BeforeDiscovery {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed -and -not $env:SUITE_WINDOW_GUARD_REPORTED) { Write-Host ('Window tests skipped: {0}.' -f $script:WindowGuard.Reason) }
    $script:Cases = @(foreach ($dark in $true, $false) {
            $design = Get-DesignSize
            foreach ($size in @(@(1024, 720), @($design.W, $design.H))) {
                @{ Dark = $dark; Theme = $(if ($dark) { 'dark' } else { 'light' }); W = $size[0]; H = $size[1]; Size = ('{0}x{1}' -f $size[0], $size[1]) }
            }
        })
    $script:ViewCases = @(@{ View = 'DPs' }, @{ View = 'Content' }, @{ View = 'Status Issues' })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed) { return }
    . "$PSScriptRoot\WindowConfig.ps1"
    $script:Design = Get-DesignSize
    $wa = [System.Windows.SystemParameters]::WorkArea
    Write-Host ('Work area of this session: {0}x{1}' -f $wa.Width, $wa.Height)
    $script:Measured = @{}
    foreach ($case in @(
            @{ Dark = $true; W = 1024; H = 720 }, @{ Dark = $true; W = $script:Design.W; H = $script:Design.H },
            @{ Dark = $false; W = 1024; H = 720 }, @{ Dark = $false; W = $script:Design.W; H = $script:Design.H })) {
        $key = '{0}|{1}x{2}' -f $(if ($case.Dark) { 'dark' } else { 'light' }), $case.W, $case.H
        $case.Config = $script:AppConfig
        $script:Measured[$key] = Invoke-WpfSta -Arguments $case -Script {
            $dsCfg = $Arguments.Config
            $dsW = $null
            try {
                $dsW = Open-TestWindow -XamlPath (Join-Path $Root $dsCfg.Xaml) -Dark $Arguments.Dark -Width $Arguments.W -Height $Arguments.H
                Wait-Layout $dsW
                $dsSidebar = @(foreach ($dsV in $dsCfg.Views) { $dsB = $dsW.FindName($dsV.Button); $dsP = $dsB.TransformToAncestor($dsW).Transform([System.Windows.Point]::new(0, 0)); [pscustomobject]@{ Name = $dsV.Button; Top = $dsP.Y; Bottom = $dsP.Y + $dsB.ActualHeight; Height = $dsB.ActualHeight } })
                $dsViews = @(foreach ($dsView in $dsCfg.Views) {
                        foreach ($dsV in $dsCfg.Views) { $dsW.FindName($dsV.Container).Visibility = if ($dsV.Name -eq $dsView.Name) { 'Visible' } else { 'Collapsed' } }
                        Wait-Layout $dsW
                        $dsC = $dsW.FindName($dsView.Container)
                        $dsInline = @(foreach ($dsEl in (Get-VisualDescendants $dsC)) { if ($dsEl -is [System.Windows.Controls.Button] -and $dsEl.Style -and (Test-VisibleInTree $dsEl) -and $dsEl.Name) { [pscustomobject]@{ Name = $dsEl.Name; Height = $dsEl.ActualHeight } } })
                        $dsGrids = @(foreach ($dsEl in (Get-VisualDescendants $dsC)) {
                                if ($dsEl -is [System.Windows.Controls.DataGrid] -and (Test-VisibleInTree $dsEl)) {
                                    $dsHdr = $null; foreach ($dsX in (Get-VisualDescendants $dsEl)) { if ($dsX -is [System.Windows.Controls.Primitives.DataGridColumnHeadersPresenter]) { $dsHdr = $dsX; break } }
                                    $dsRow = $null; foreach ($dsX in (Get-VisualDescendants $dsEl)) { if ($dsX -is [System.Windows.Controls.DataGridRow]) { $dsRow = $dsX; break } }
                                    [pscustomobject]@{ Name = $dsEl.Name; Width = $dsEl.ActualWidth; HeaderHeight = $(if ($dsHdr) { $dsHdr.ActualHeight } else { $null }); RowHeight = $(if ($dsRow) { $dsRow.ActualHeight } else { $null }) }
                                }
                            })
                        # The vacant area: the free height under, and the free width beside, the lowest and the right-most visible control of the view.
                        $dsBottom = 0.0; $dsRight = 0.0
                        foreach ($dsEl in (Get-VisualDescendants $dsC)) {
                            if ($dsEl -is [System.Windows.Controls.Primitives.ButtonBase] -or $dsEl -is [System.Windows.Controls.TextBox] -or $dsEl -is [System.Windows.Controls.DataGrid] -or $dsEl -is [System.Windows.Controls.ComboBox]) {
                                if ((Test-VisibleInTree $dsEl) -and $dsEl.ActualWidth -gt 0 -and $dsEl.ActualHeight -gt 0) {
                                    $dsP = $dsEl.TransformToAncestor($dsC).Transform([System.Windows.Point]::new(0, 0))
                                    $dsBottom = [Math]::Max($dsBottom, $dsP.Y + $dsEl.ActualHeight); $dsRight = [Math]::Max($dsRight, $dsP.X + $dsEl.ActualWidth)
                                }
                            }
                        }
                        [pscustomobject]@{ View = $dsView.Name; Width = $dsC.ActualWidth; Height = $dsC.ActualHeight; Buttons = $dsInline; Grids = $dsGrids; VacantHeight = ($dsC.ActualHeight - $dsBottom); VacantWidth = ($dsC.ActualWidth - $dsRight) }
                    })
                [pscustomobject]@{ Width = $dsW.ActualWidth; Height = $dsW.ActualHeight; Sidebar = $dsSidebar; Views = $dsViews }
            }
            finally { if ($dsW) { $dsW.Close() } }
        }
    }
}

AfterAll {
    if ($script:WindowGuard.Allowed) { Close-WpfSession }
}

Describe 'Density <Theme> <Size>' -ForEach $script:Cases -Skip:(-not $script:WindowGuard.Allowed) {
    It 'sizes every sidebar button between 34 and 36 px and fits all of them without scrolling' {
        $m = $script:Measured['{0}|{1}' -f $Theme, $Size]
        foreach ($b in $m.Sidebar) { $b.Height | Should -BeIn 34, 35, 36 -Because $b.Name }
        ($m.Sidebar | Select-Object -Last 1).Bottom | Should -BeLessThan $m.Height
    }

    It 'sizes the inline buttons at 28 px or less in the <View> view' -ForEach $script:ViewCases {
        $m = $script:Measured['{0}|{1}' -f $Theme, $Size].Views | Where-Object View -eq $View
        foreach ($b in $m.Buttons | Where-Object { $_.Name -notmatch '^btnView|^btnOptions' }) { $b.Height | Should -BeLessOrEqual 28 -Because $b.Name }
    }

    It 'gives grid rows 24 to 26 px and headers 26 px in the <View> view' -ForEach $script:ViewCases {
        $m = $script:Measured['{0}|{1}' -f $Theme, $Size].Views | Where-Object View -eq $View
        foreach ($g in $m.Grids) {
            if ($null -ne $g.HeaderHeight) { $g.HeaderHeight | Should -BeLessOrEqual 26.5 -Because $g.Name }
            if ($null -ne $g.RowHeight) { [math]::Round($g.RowHeight) | Should -BeIn 24, 25, 26 -Because $g.Name }
        }
    }

    It 'lets the content fill the width in the <View> view' -ForEach $script:ViewCases {
        $mm = $script:Measured['{0}|{1}' -f $Theme, $Size]
        $m = $mm.Views | Where-Object View -eq $View
        $m.Width | Should -BeGreaterOrEqual ($mm.Width - 300)
        foreach ($g in $m.Grids) { $g.Width | Should -BeGreaterOrEqual ($m.Width * 0.3) -Because $g.Name }
    }
}

Describe 'Vacant area <Theme>' -ForEach @(@{ Theme = 'dark' }, @{ Theme = 'light' }) -Skip:(-not $script:WindowGuard.Allowed) {
    It 'leaves no vacant area larger than 120 px high or 160 px wide in any view at the work area' {
        $m = $script:Measured['{0}|{1}' -f $Theme, $script:Design.Key]
        $bad = @($m.Views | Where-Object { $_.VacantHeight -gt 120 -or $_.VacantWidth -gt 160 } | ForEach-Object { '{0}: {1} px high, {2} px wide vacant' -f $_.View, [math]::Round($_.VacantHeight), [math]::Round($_.VacantWidth) })
        $bad | Should -BeNullOrEmpty -Because ('vacant area: ' + ($bad -join '; '))
    }
}
