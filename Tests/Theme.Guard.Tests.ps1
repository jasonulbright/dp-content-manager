#Requires -Modules Pester
# Guard tests for the theme rules. They parse the files and create no WPF control, so they run on
# Windows PowerShell 5.1 and PowerShell 7 and open no window.
#   (a) Every blue background in XAML and PowerShell has a white foreground.
#   (b) No font size, weight, style or text decoration depends on the theme.
#   (c) No banned word is in a tooltip, a UI string, the README or the user guide.

BeforeDiscovery {
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
    $script:RepoRootD = Split-Path -Parent $PSScriptRoot
    $script:PsFilesD = @(
        'start-dpcontentmgr.ps1', 'Module\DPContentMgrCommon.psm1', 'Lib\SuiteCommon\SuiteCommon.psm1'
    ) | ForEach-Object { @{ Name = $_; Path = (Join-Path $script:RepoRootD $_) } }
}

BeforeAll {
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
    $script:XamlDocs = @(Get-XamlDocument)
}

Describe 'Blue backgrounds have white text (a)' {
    It 'finds XAML to check' {
        $script:XamlDocs.Count | Should -BeGreaterThan 1
    }

    It 'sets a white foreground on every blue background in <_.Source>' -ForEach @(Get-XamlDocument) {
        $bad = @(Get-BlueWithoutWhite -Text $_.Text -Source $_.Source)
        $bad | Should -BeNullOrEmpty
    }

    It 'recognizes blue and white brushes' {
        Test-BlueBrush '{DynamicResource MahApps.Brushes.Accent}' | Should -BeTrue
        Test-BlueBrush '#0078D4' | Should -BeTrue
        Test-BlueBrush '#1E1E1E' | Should -BeFalse
        Test-BlueBrush '#3A3A3A' | Should -BeFalse
        Test-WhiteBrush '{DynamicResource MahApps.Brushes.IdealForeground}' | Should -BeTrue
        Test-WhiteBrush 'Black' | Should -BeFalse
    }

    It 'flags a blue background that has no white foreground' {
        $xaml = '<Style xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" TargetType="Button"><Setter Property="Background" Value="#0078D4"/></Style>'
        @(Get-BlueWithoutWhite -Text $xaml -Source 'sample').Count | Should -Be 1
        $ok = '<Style xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" TargetType="Button"><Setter Property="Background" Value="#0078D4"/><Setter Property="Foreground" Value="White"/></Style>'
        @(Get-BlueWithoutWhite -Text $ok -Source 'sample').Count | Should -Be 0
    }

    It 'has no blue brush literal in <Name>, and each CSS rule with a blue background has white text' -ForEach @(
        @{ Name = 'start-dpcontentmgr.ps1'; Path = (Join-Path (Split-Path -Parent $PSScriptRoot) 'start-dpcontentmgr.ps1') }
        @{ Name = 'DPContentMgrCommon.psm1'; Path = (Join-Path (Split-Path -Parent $PSScriptRoot) 'Module\DPContentMgrCommon.psm1') }
    ) {
        $bad = foreach ($s in @(Get-ScriptString -Path $Path)) {
            if ($s -notmatch '<' -and (Test-BlueBrush $s)) { 'literal ' + $s }
            foreach ($m in [regex]::Matches($s, '\{([^}]*)\}')) {
                $rule = $m.Groups[1].Value
                $bg = [regex]::Match($rule, '(?i)(?<![-\w])background(?:-color)?\s*:\s*(#[0-9A-F]{6,8})\b')
                if ($bg.Success -and (Test-BlueBrush $bg.Groups[1].Value) -and $rule -notmatch '(?i)(?<![-\w])color\s*:\s*(#fff(fff)?|white)\b') { 'css ' + $rule }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'paints the light theme buttons of SuiteCommon with white text' {
        $path = Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1'
        $text = Get-Content -LiteralPath $path -Raw
        $text | Should -Match "OnBlue\s*=\s*\`$bc\.ConvertFrom\('#FFFFFF'\)"
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
        foreach ($name in 'Update-SidebarButtonTheme', 'Set-ButtonTheme') {
            $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
            $fn | Should -Not -BeNullOrEmpty -Because $name
            $fn.Extent.Text | Should -Match 'Set-ThemeForeground' -Because ('{0} sets the foreground with each background' -f $name)
        }
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Set-ThemeForeground' }, $true)
        $fn.Extent.Text | Should -Match '\.Foreground\s*=\s*\$OnBlue'
    }
}

Describe 'Fonts do not depend on the theme (b)' {
    It 'has no font property in a theme branch of <Name>' -ForEach $script:PsFilesD {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
        $assigns = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] }, $true)
        $bad = foreach ($a in $assigns) {
            $member = [string]$a.Left.Member.Value
            if ($script:ThemeProperty -notcontains $member) { continue }
            $p = $a.Parent
            while ($p) {
                $condText = ''
                if ($p -is [System.Management.Automation.Language.IfStatementAst]) { $condText = (@($p.Clauses | ForEach-Object { $_.Item1.Extent.Text }) -join ' ') }
                elseif ($p -is [System.Management.Automation.Language.SwitchStatementAst]) { $condText = $p.Condition.Extent.Text }
                if ($condText -match '(?i)IsDark|DarkMode|Light|Theme|IsOn') { '{0}:{1} sets {2} in a theme branch' -f $Name, $a.Extent.StartLineNumber, $member; break }
                $p = $p.Parent
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'sets no font property in the theme functions of SuiteCommon except the connection boxes' {
        $path = Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1'
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
        foreach ($name in 'Set-ThemeForeground', 'Update-SidebarButtonTheme', 'Set-ButtonTheme', 'Set-DialogTheme', 'Update-TitleBarBrushes', 'Initialize-SuiteTheme') {
            $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
            $fn | Should -Not -BeNullOrEmpty -Because $name
            $fn.Extent.Text | Should -Not -Match '\.(FontSize|FontWeight|FontStyle|TextDecorations)\s*=' -Because $name
        }
    }

    It 'makes the connection boxes bold and 14 pt or larger the same way in both themes' {
        $path = Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1'
        $text = Get-Content -LiteralPath $path -Raw
        $text | Should -Match 'ConnectionBoxFontSize\s*=\s*14'
        $text | Should -Match '\.FontWeight\s*=\s*\[System\.Windows\.FontWeights\]::Bold'
    }

    It 'uses no theme trigger on a font property in <_.Source>' -ForEach @(Get-XamlDocument) {
        [xml]$x = $_.Text
        $bad = foreach ($s in @($x.SelectNodes("//*[local-name()='Setter']"))) {
            if ($script:ThemeProperty -contains $s.GetAttribute('Property')) {
                $p = $s.ParentNode
                if ($p.LocalName -match 'Trigger') { $p.OuterXml.Substring(0, [Math]::Min(80, $p.OuterXml.Length)) }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'sets the site code and SMS Provider boxes to bold 14 pt or more in <_.Source>' -ForEach @(Get-XamlDocument) {
        [xml]$x = $_.Text
        foreach ($name in 'txtSiteCode', 'txtSmsProvider', 'txtProvider', 'txtSiteServer') {
            $el = $x.SelectSingleNode(("//*[@*[local-name()='Name']='{0}']" -f $name))
            if (-not $el) { continue }
            [double]$el.GetAttribute('FontSize') | Should -BeGreaterOrEqual 14 -Because $name
            $el.GetAttribute('FontWeight') | Should -Be 'Bold' -Because $name
        }
    }
}

Describe 'No banned word in the UI text and the documents (c)' {
    It 'finds the banned word <Word>' -ForEach @(@{ Word = 'simply' }, @{ Word = 'Helps you' }, @{ Word = 'streamlines' }) {
        $Word | Should -Match $script:BannedPattern
    }

    It 'has no banned word in the XAML text of <_.Source>' -ForEach @(Get-XamlDocument) {
        $bad = @(Get-XamlUiText -Text $_.Text | Where-Object { $_ -match $script:BannedPattern })
        $bad | Should -BeNullOrEmpty
    }

    It 'has no banned word in the strings of <Name>' -ForEach $script:PsFilesD {
        $bad = @(Get-ScriptString -Path $Path | Where-Object { $_ -match $script:BannedPattern })
        $bad | Should -BeNullOrEmpty
    }

    It 'has no banned word in the README or the user guide' {
        $files = @(Get-Item -LiteralPath (Join-Path $script:RepoRoot 'README.md')) + @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'docs') -Filter '*.md' -File -Recurse)
        $bad = foreach ($f in $files) {
            $n = 0
            foreach ($line in (Get-Content -LiteralPath $f.FullName)) {
                $n++
                if ($line -match $script:BannedPattern) { '{0}:{1}: {2}' -f $f.Name, $n, $Matches[0] }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }
}
