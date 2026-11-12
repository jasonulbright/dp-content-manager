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

    It 'makes the field titles bold and 14 pt or larger the same way in both themes' {
        $path = Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1'
        $text = Get-Content -LiteralPath $path -Raw
        $text | Should -Match 'FieldTitleFontSize\s*=\s*14'
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

Describe 'Theme helpers of SuiteCommon behave and are used (a, b, d)' {
    BeforeAll {
        $script:SuitePath = Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1'
        $script:SuiteAst = [System.Management.Automation.Language.Parser]::ParseFile($script:SuitePath, [ref]$null, [ref]$null)
        function Get-SuiteFunction {
            param([string]$Name)
            $script:SuiteAst.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
        }
        function Get-CommandCount {
            param($Ast, [string]$Name)
            @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq $Name }, $true)).Count
        }
        . ([scriptblock]::Create((Get-SuiteFunction 'Set-ThemeForeground').Extent.Text))
    }

    It 'Set-ThemeForeground sets the on-blue brush in light mode and clears it in dark mode' {
        $c = [pscustomobject]@{ Foreground = $null }
        Set-ThemeForeground -Control $c -IsDark $false -OnBlue 'White'
        $c.Foreground | Should -Be 'White'
        Set-ThemeForeground -Control $c -IsDark $true -OnBlue 'White'
        $c.Foreground | Should -BeNullOrEmpty
    }

    It '<Name> calls Set-ThemeForeground' -ForEach @(@{ Name = 'Set-ButtonTheme' }, @{ Name = 'Update-SidebarButtonTheme' }) {
        (Get-CommandCount (Get-SuiteFunction $Name) 'Set-ThemeForeground') | Should -BeGreaterThan 0
    }

    It '<Name> calls Set-ConnectionTextTheme' -ForEach @(@{ Name = 'Set-ButtonTheme' }, @{ Name = 'Update-SidebarButtonTheme' }) {
        (Get-CommandCount (Get-SuiteFunction $Name) 'Set-ConnectionTextTheme') | Should -BeGreaterThan 0
    }

    It 'the script starts the theme and every dialog calls Set-DialogTheme' {
        $shell = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-dpcontentmgr.ps1'), [ref]$null, [ref]$null)
        (Get-CommandCount $shell 'Initialize-SuiteTheme') | Should -BeGreaterThan 0
        (Get-CommandCount $shell 'Update-SidebarButtonTheme') | Should -BeGreaterThan 0
        $dialogs = @($script:XamlDocs | Where-Object { $_.Source -like 'start-dpcontentmgr*here-string*' }).Count
        (Get-CommandCount $shell 'Set-DialogTheme') | Should -BeGreaterOrEqual $dialogs
        (Get-SuiteFunction 'Set-DialogTheme') | Should -Not -BeNullOrEmpty
    }

    It 'sets a white foreground in every style based on an accent style, in <_.Source>' -ForEach @(Get-XamlDocument | Where-Object { $_.Source -notlike 'SuiteCommon.psm1*' }) {
        [xml]$x = $_.Text
        $bad = foreach ($st in @($x.SelectNodes("//*[local-name()='Style']"))) {
            if ($st.GetAttribute('BasedOn') -match 'Accent') {
                $fg = $st.ChildNodes | Where-Object { $_.LocalName -eq 'Setter' -and $_.GetAttribute('Property') -eq 'Foreground' } | Select-Object -First 1
                if (-not $fg -or -not (Test-WhiteBrush (Get-SetterValue $fg))) { 'Style {0} has no white Foreground' -f $st.GetAttribute('Key') }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'has the site code and SMS Provider boxes with a bold 14 pt title and a normal value' {
        foreach ($name in 'txtSiteCode', 'txtSmsProvider') {
            $hits = @()
            foreach ($d in $script:XamlDocs) {
                [xml]$x = $d.Text
                $el = $x.SelectSingleNode(("//*[@*[local-name()='Name']='{0}']" -f $name))
                if ($el) { $hits += [pscustomobject]@{ Value = $el; Title = (Get-FieldTitleElement $el) } }
            }
            $hits.Count | Should -BeGreaterThan 0 -Because ('{0} must exist in a document' -f $name)
            foreach ($h in $hits) {
                $h.Title | Should -Not -BeNullOrEmpty -Because ('{0} needs a title TextBlock before it' -f $name)
                [double]$h.Title.GetAttribute('FontSize') | Should -BeGreaterOrEqual 14 -Because ('title of {0}' -f $name)
                $h.Title.GetAttribute('FontWeight') | Should -Be 'Bold' -Because ('title of {0}' -f $name)
                $h.Value.GetAttribute('FontWeight') | Should -BeIn @('', 'Normal') -Because ('value {0}' -f $name)
                if ($h.Value.GetAttribute('FontSize')) { [double]$h.Value.GetAttribute('FontSize') | Should -BeLessThan 14 -Because ('value {0}' -f $name) }
            }
        }
    }

    It 'SuiteCommon bolds the titles with 14 pt and not the values' {
        $text = Get-Content -LiteralPath $script:SuitePath -Raw
        $text | Should -Match 'FieldTitleFontSize\s*=\s*14'
        $text | Should -Match 'Find-FieldTitle'
        $text | Should -Not -Match '\$box\.FontWeight'
    }
}

Describe 'Every button takes the shared themed style' {
    BeforeAll {
        function Get-StyleChain {
            param($Doc, [string]$Key)
            $guard = 0
            while ($Key -and $guard -lt 10) {
                $guard++
                $st = $Doc.SelectSingleNode(("//*[local-name()='Style'][@*[local-name()='Key']='{0}']" -f $Key))
                if (-not $st) { return $false }
                $based = $st.GetAttribute('BasedOn')
                if ($based -match 'MahApps\.Styles\.Button') { return $true }
                $Key = if ($based -match 'StaticResource\s+([^\}\s]+)') { $Matches[1] } else { '' }
            }
            return $false
        }
        $script:ShellText = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'start-dpcontentmgr.ps1') -Raw
    }

    # The dialogs of the vendored SuiteCommon are outside this scope: the app does not open them, and the fix belongs upstream.
    It 'gives each Button in <_.Source> a style that is based on a MahApps button style' -ForEach @(Get-XamlDocument | Where-Object { $_.Source -notlike 'SuiteCommon.psm1*' }) {
        [xml]$x = $_.Text
        $bad = foreach ($b in @($x.SelectNodes("//*[local-name()='Button']"))) {
            $style = $b.GetAttribute('Style')
            $key = if ($style -match '(?:Static|Dynamic)Resource\s+([^\}\s]+)') { $Matches[1] } else { '' }
            if (-not $key -or -not (Get-StyleChain -Doc $x -Key $key)) { 'Button {0} has no shared style' -f $b.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml') }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'creates no button in code' {
        foreach ($f in 'start-dpcontentmgr.ps1', 'Module\DPContentMgrCommon.psm1') {
            (Get-Content -LiteralPath (Join-Path $script:RepoRoot $f) -Raw) | Should -Not -Match 'Controls\.Button\]|New-Object\s+(-TypeName\s+)?System\.Windows\.Controls\.Button' -Because $f
        }
    }

    It 'themes every button of the main window through SuiteCommon' {
        [xml]$x = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'MainWindow.xaml') -Raw
        $names = @($x.SelectNodes("//*[local-name()='Button']") | ForEach-Object { $_.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml') })
        $names.Count | Should -BeGreaterThan 5
        $wf = [regex]::Match($script:ShellText, '\$script:WorkflowButtons\s*=\s*@\(([^)]*)\)').Groups[1].Value
        $views = ([regex]::Matches($script:ShellText, 'Button\s*=\s*\$(\w+)') | ForEach-Object { $_.Groups[1].Value }) -join ' '
        $opt = [regex]::Match($script:ShellText, '-OptionsButton\s+\$(\w+)').Groups[1].Value
        $themed = ($wf + ' ' + $views + ' ' + $opt)
        $unthemed = $names | Where-Object { $themed -notmatch ('\b{0}\b' -f $_) }
        @($unthemed) | Should -BeNullOrEmpty
        $script:ShellText | Should -Match '-WorkflowButtons\s+\$script:WorkflowButtons'
        $script:ShellText | Should -Match '(?m)^Set-ButtonTheme'
    }

    It 'themes every button of each dialog: an accent style or Set-DialogButtonTheme' {
        $docs = @($script:XamlDocs | Where-Object { $_.Source -like 'start-dpcontentmgr*here-string*' })
        $docs.Count | Should -BeGreaterThan 0
        $called = ([regex]::Matches($script:ShellText, 'Set-DialogButtonTheme\s+-Dialog\s+\$dlg\s+-Name\s+([^\r\n]+)') | ForEach-Object { $_.Groups[1].Value }) -join ' '
        $bad = foreach ($d in $docs) {
            [xml]$x = $d.Text
            foreach ($b in @($x.SelectNodes("//*[local-name()='Button']"))) {
                $name = $b.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
                $style = $b.GetAttribute('Style')
                $key = if ($style -match 'StaticResource\s+([^\}\s]+)') { $Matches[1] } else { '' }
                $st = $x.SelectSingleNode(("//*[local-name()='Style'][@*[local-name()='Key']='{0}']" -f $key))
                $accent = $st -and $st.GetAttribute('BasedOn') -match 'Accent'
                if (-not $accent -and $called -notmatch ("'{0}'" -f $name)) { '{0}: {1} is not themed' -f $d.Source, $name }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }
}

Describe 'No raw exception text in the UI' {
    It 'flags raw exception text in a UI sink' {
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { Add-LogLine ("x {0}" -f $_.Exception.Message) }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { $row.Detail = $_.Exception.Message }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { Write-Log ("x {0}" -f $_.Exception.Message) -Level ERROR -Quiet; Add-LogLine "Failed. See the log file." }').Count | Should -Be 0
    }

    It 'sends no raw exception text to a UI sink in <Name>' -ForEach $script:PsFilesD {
        @(Get-RawExceptionToUi -Path $Path) | Should -BeNullOrEmpty
    }
}
