#Requires -Modules Pester
# Window tests. The main window and the two dialogs the script builds from
# XAML text are opened off screen on an STA thread, in both themes and at
# 1024x720 and 1440x900, and measured.

BeforeDiscovery {
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:Matrix = foreach ($dark in $true, $false) {
        foreach ($size in @(@(1024, 720), @(1440, 900))) {
            @{ Dark = $dark; Theme = $(if ($dark) { 'dark' } else { 'light' }); W = $size[0]; H = $size[1]; Size = ('{0}x{1}' -f $size[0], $size[1]) }
        }
    }
}

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    . "$PSScriptRoot\WindowHarness.ps1"
    . "$PSScriptRoot\KnownContrastDefects.ps1"
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:ShellPath = Join-Path $script:Root 'start-dpcontentmgr.ps1'
    $script:XamlPath = Join-Path $script:Root 'MainWindow.xaml'

    # Controls the script reads with $window.FindName, taken from the script
    # itself so the two cannot drift apart.
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)
    $script:WindowNames = @($ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'FindName' -and
                $n.Expression.Extent.Text -eq '$window' -and $n.Arguments.Count -eq 1
            }, $true) | ForEach-Object { $_.Arguments[0].Value } | Sort-Object -Unique)

    # Dialog XAML kept as here-strings in the script.
    $script:DialogXaml = @($ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -match '^\s*<Controls:MetroWindow'
            }, $true) | ForEach-Object { $_.Value })

    $script:ViewNames = 'DPs', 'Content', 'Status Issues'
    $script:Measured = @{}
    foreach ($case in @(
            @{ Dark = $true; W = 1024; H = 720 }, @{ Dark = $true; W = 1440; H = 900 },
            @{ Dark = $false; W = 1024; H = 720 }, @{ Dark = $false; W = 1440; H = 900 })) {
        $key = '{0}|{1}x{2}' -f $(if ($case.Dark) { 'dark' } else { 'light' }), $case.W, $case.H
        $script:Measured[$key] = Invoke-WpfSta -Arguments $case -Script {
            $w = Open-TestWindow -XamlPath (Join-Path $Root 'MainWindow.xaml') -Dark $Arguments.Dark -Width $Arguments.W -Height $Arguments.H
            $out = @()
            foreach ($view in 'DPs', 'Content', 'Status Issues') {
                $w.FindName('viewDPs').Visibility = if ($view -eq 'DPs') { 'Visible' } else { 'Collapsed' }
                $w.FindName('viewContent').Visibility = if ($view -eq 'Content') { 'Visible' } else { 'Collapsed' }
                $w.FindName('viewStatusIssues').Visibility = if ($view -eq 'Status Issues') { 'Visible' } else { 'Collapsed' }
                Initialize-ShellTheme -Window $w -Dark $Arguments.Dark -ViewButtonNames 'btnViewDPs', 'btnViewContent', 'btnViewStatusIssues' `
                    -OptionsButtonName 'btnOptions' -LogLabelName 'lblLogOutput' -ActiveView $view
                $out += [PSCustomObject]@{
                    View        = $view
                    Text        = @(Get-TextContrast $w)
                    Clipped     = @(Get-ClippedText $w)
                    Offscreen   = @(Get-OffscreenControls $w)
                    Interactive = @(Get-InteractiveContrast $w)
                }
            }
            $w.Close()
            , $out
        }
    }
}

AfterAll { Close-WpfSession }

Describe 'Main window XAML' {
    It 'loads as a MetroWindow titled DP Content Manager' {
        $r = Invoke-WpfSta -Script {
            $w = Open-TestWindow -XamlPath (Join-Path $Root 'MainWindow.xaml')
            $o = [PSCustomObject]@{ Type = $w.GetType().FullName; Title = $w.Title; MinW = $w.MinWidth; MinH = $w.MinHeight }
            $w.Close(); $o
        }
        $r.Type | Should -Be 'MahApps.Metro.Controls.MetroWindow'
        $r.Title | Should -Be 'DP Content Manager'
        $r.MinW | Should -BeLessOrEqual 1024
        $r.MinH | Should -BeLessOrEqual 720
    }

    It 'declares each x:Name once' {
        $names = [regex]::Matches((Get-Content -LiteralPath $script:XamlPath -Raw), 'x:Name="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
        ($names | Group-Object | Where-Object Count -gt 1 | ForEach-Object Name) | Should -BeNullOrEmpty
    }

    It 'has every control the script looks up (<_>)' -ForEach @(
        'txtAppTitle', 'txtVersion', 'txtThemeLabel', 'toggleTheme', 'btnViewDPs', 'btnViewContent', 'btnViewStatusIssues', 'btnOptions',
        'txtModuleTitle', 'txtModuleSubtitle', 'btnRefresh', 'txtFilter', 'cboStatusFilter', 'btnExportCsv', 'btnExportHtml',
        'viewDPs', 'viewContent', 'viewStatusIssues', 'gridDPs', 'txtDPDetail', 'gridContent', 'gridContentPerDP',
        'btnRedistribute', 'btnRemoveFromDP', 'btnValidate', 'gridStatusIssues', 'btnRedistributeIssues', 'btnRemoveIssues',
        'txtIssueCount', 'progressOverlay', 'txtProgressTitle', 'txtProgressStep', 'lblLogOutput', 'txtLog', 'txtStatus'
    ) {
        $script:WindowNames | Should -Contain $_
    }

    It 'resolves every FindName in the script to a control in the XAML' {
        $xaml = Get-Content -LiteralPath $script:XamlPath -Raw
        $missing = @($script:WindowNames | Where-Object { $xaml -notmatch ('x:Name="{0}"' -f [regex]::Escape($_)) })
        $missing | Should -BeNullOrEmpty
    }

    It 'shows the DPs view first and keeps the other views hidden' {
        $r = Invoke-WpfSta -Script {
            $w = Open-TestWindow -XamlPath (Join-Path $Root 'MainWindow.xaml')
            $o = [PSCustomObject]@{ D = [string]$w.FindName('viewDPs').Visibility; C = [string]$w.FindName('viewContent').Visibility; S = [string]$w.FindName('viewStatusIssues').Visibility; P = [string]$w.FindName('progressOverlay').Visibility }
            $w.Close(); $o
        }
        $r.D | Should -Be 'Visible'
        $r.C | Should -Be 'Collapsed'
        $r.S | Should -Be 'Collapsed'
        $r.P | Should -Be 'Collapsed'
    }

    It 'starts with Refresh, the filter and the exports enabled' {
        $r = Invoke-WpfSta -Script {
            $w = Open-TestWindow -XamlPath (Join-Path $Root 'MainWindow.xaml')
            $o = foreach ($n in 'btnRefresh', 'txtFilter', 'cboStatusFilter', 'btnExportCsv', 'btnExportHtml', 'btnRedistribute', 'btnRemoveFromDP', 'btnValidate', 'btnRedistributeIssues', 'btnRemoveIssues') { [PSCustomObject]@{ N = $n; E = $w.FindName($n).IsEnabled } }
            $w.Close(); , @($o)
        }
        ($r | Where-Object { -not $_.E }).N | Should -BeNullOrEmpty
    }

    It 'makes the log read only' {
        $r = Invoke-WpfSta -Script {
            $w = Open-TestWindow -XamlPath (Join-Path $Root 'MainWindow.xaml')
            $o = [PSCustomObject]@{ Log = $w.FindName('txtLog').IsReadOnly; Detail = $w.FindName('txtDPDetail').IsReadOnly }
            $w.Close(); $o
        }
        $r.Log | Should -BeTrue
        $r.Detail | Should -BeTrue
    }
}

Describe 'Dialog windows' {
    It 'finds the Options and bulk operation dialog XAML in the script' {
        $script:DialogXaml.Count | Should -Be 2
    }

    It 'parses dialog <Index> as a MetroWindow with named controls' -ForEach @(@{ Index = 0 }, @{ Index = 1 }) {
        $r = Invoke-WpfSta -Arguments @{ Xaml = $script:DialogXaml[$Index] } -Script {
            [xml]$x = $Arguments.Xaml
            $w = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
            $names = [regex]::Matches($Arguments.Xaml, 'x:Name="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
            [PSCustomObject]@{ Type = $w.GetType().Name; Missing = @($names | Where-Object { -not $w.FindName($_) }); Count = @($names).Count }
        }
        $r.Type | Should -Be 'MetroWindow'
        $r.Missing | Should -BeNullOrEmpty
        $r.Count | Should -BeGreaterThan 3
    }

    It 'gives every dialog a title that differs from the main window title' {
        $titles = @([regex]::Matches((Get-Content -LiteralPath $script:ShellPath -Raw), "-Title\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value })
        $titles += @($script:DialogXaml | ForEach-Object { if ($_ -match '\sTitle="([^"]*)"') { $Matches[1] } } | Where-Object { $_ })
        $titles.Count | Should -BeGreaterThan 3
        $titles | Should -Not -Contain 'DP Content Manager'
        $titles | ForEach-Object { $_.Trim() | Should -Not -BeNullOrEmpty }
    }

    It 'gives the four bulk dialogs different titles' {
        $titles = @([regex]::Matches((Get-Content -LiteralPath $script:ShellPath -Raw), "Show-BulkOpDialog -Title '([^']+)'") | ForEach-Object { $_.Groups[1].Value })
        $titles.Count | Should -Be 5
        # Redistribute and Remove each appear on the Content and the Issues view, with different titles.
        ($titles | Select-Object -Unique).Count | Should -Be 5
    }
}

Describe 'Layout <Theme> <Size>' -ForEach $script:Matrix {
    It 'has no clipped text in the <View> view' -ForEach @(@{ View = 'DPs' }, @{ View = 'Content' }, @{ View = 'Status Issues' }) {
        $m = ($script:Measured['{0}|{1}' -f $Theme, $Size] | Where-Object View -eq $View)
        $m.Clipped | ForEach-Object { '{0}: needs {1}px, has {2}px' -f $_.Text, $_.Needed, $_.Actual } | Should -BeNullOrEmpty
    }

    It 'keeps every button, box and list inside the window in the <View> view' -ForEach @(@{ View = 'DPs' }, @{ View = 'Content' }, @{ View = 'Status Issues' }) {
        $m = ($script:Measured['{0}|{1}' -f $Theme, $Size] | Where-Object View -eq $View)
        $m.Offscreen | ForEach-Object { '{0} at {1},{2} size {3}x{4}' -f $_.Name, $_.X, $_.Y, $_.W, $_.H } | Should -BeNullOrEmpty
    }
}

Describe 'Contrast <Theme>' -ForEach @(@{ Theme = 'dark' }, @{ Theme = 'light' }) {
    It 'shows every text at 4.5:1 (3:1 for large text) in the <View> view' -ForEach @(@{ View = 'DPs' }, @{ View = 'Content' }, @{ View = 'Status Issues' }) {
        $m = ($script:Measured['{0}|1440x900' -f $Theme] | Where-Object View -eq $View)
        $m.Text.Count | Should -BeGreaterThan 5
        $bad = $m.Text | Where-Object { $min = if ($_.Size -ge 24 -or ($_.Bold -and $_.Size -ge 18.66)) { 3.0 } else { 4.5 }; $_.Ratio -lt $min }
        $bad | ForEach-Object { '{0} "{1}" {2} on {3} = {4}' -f $_.Name, $_.Text, $_.Foreground, $_.Backdrop, $_.Ratio } | Should -BeNullOrEmpty
    }

    It 'meets the standard on every control, except the listed known defects' {
        $known = Get-KnownContrastDefect -Theme $Theme
        $violations = New-Object System.Collections.Generic.List[string]
        foreach ($v in ($script:Measured['{0}|1440x900' -f $Theme])) {
            foreach ($c in $v.Interactive) {
                foreach ($metric in 'EnabledText', 'DisabledText', 'EnabledBorder', 'DisabledBorder') {
                    $value = $c.$metric
                    if ($null -eq $value) { continue }
                    $min = if ($metric -like '*Text') { 4.5 } else { 3.0 }
                    if ($value -lt $min) { [void]$violations.Add(('{0}:{1}' -f $c.Name, $metric)) }
                }
            }
        }
        $found = @($violations | Sort-Object -Unique)
        $new = @($found | Where-Object { $known -notcontains $_ })
        $fixed = @($known | Where-Object { $found -notcontains $_ })
        $new | Should -BeNullOrEmpty -Because 'a new contrast violation appeared'
        $fixed | Should -BeNullOrEmpty -Because 'a known defect is fixed: remove it from Get-KnownContrastDefect'
    }
}
