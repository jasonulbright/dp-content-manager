#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
# Headless guards. No window opens here.
#   - Invoke-WpfSta closes every window a call left open, in its finally.
#   - A window test opens a window only inside a try that has a finally.
#   - Every button style based on the MahApps Square or Square.Accent style starts with the shared Suite template setter.

BeforeDiscovery {
    $script:TestFiles = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Tests.ps1' -File |
            Where-Object { $_.Name -ne 'WindowCleanup.Guard.Tests.ps1' -and (Get-Content -LiteralPath $_.FullName -Raw) -match '\.Show\(\)|ShowDialog|Open-TestWindow' } |
            ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } })
    $script:StyleFiles = @(
        @{ Name = 'MainWindow.xaml'; Path = (Join-Path (Split-Path -Parent $PSScriptRoot) 'MainWindow.xaml') }
        @{ Name = 'start-dpcontentmgr.ps1'; Path = (Join-Path (Split-Path -Parent $PSScriptRoot) 'start-dpcontentmgr.ps1') }
    )
}

Describe 'Invoke-WpfSta closes stray windows' {
    BeforeAll {
        $script:HarnessPath = Join-Path $PSScriptRoot 'WindowHarness.ps1'
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:HarnessPath, [ref]$null, [ref]$null)
        $script:Fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-WpfSta' }, $true)
    }

    It 'calls Close-SuiteOpenWindow inside the finally of Invoke-WpfSta' {
        $script:Fn | Should -Not -BeNullOrEmpty
        $try = $script:Fn.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] -and $null -ne $n.Finally }, $true) | Select-Object -First 1
        $try | Should -Not -BeNullOrEmpty
        $try.Finally.Extent.Text | Should -Match 'Close-SuiteOpenWindow'
    }
}

Describe 'Window tests close what they open' {
    It '<Name> opens a window only inside a try that has a finally' -ForEach $script:TestFiles {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
        $opens = @($ast.FindAll({
                    param($n)
                    ($n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Extent.Text -in 'Show', 'ShowDialog') -or
                    ($n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Open-TestWindow')
                }, $true))
        $bad = @()
        foreach ($o in $opens) {
            $p = $o.Parent
            $guarded = $false
            # Open-TestWindow in a statement that a try with a finally follows at once also counts: the finally closes the window.
            $stmt = $o
            while ($stmt -and $stmt.Parent -and $stmt.Parent -isnot [System.Management.Automation.Language.StatementBlockAst] -and $stmt.Parent -isnot [System.Management.Automation.Language.NamedBlockAst]) { $stmt = $stmt.Parent }
            if ($stmt -and $stmt.Parent) {
                $siblings = @($stmt.Parent.Statements)
                $at = [array]::IndexOf($siblings, $stmt)
                if ($at -ge 0 -and $at + 1 -lt $siblings.Count -and $siblings[$at + 1] -is [System.Management.Automation.Language.TryStatementAst] -and $null -ne $siblings[$at + 1].Finally) { $guarded = $true }
            }
            while ($p -and -not $guarded) {
                if ($p -is [System.Management.Automation.Language.TryStatementAst] -and $null -ne $p.Finally -and $p.Body.Extent.StartOffset -le $o.Extent.StartOffset -and $p.Body.Extent.EndOffset -ge $o.Extent.EndOffset) { $guarded = $true; break }
                $p = $p.Parent
            }
            if (-not $guarded) { $bad += ('line {0}: {1}' -f $o.Extent.StartLineNumber, $o.Extent.Text) }
        }
        $bad | Should -BeNullOrEmpty
    }
}

Describe 'Button styles take the shared Suite template' {
    It '<Name>: a style based on MahApps Square or Square.Accent that is not a sidebar style starts with the Suite template setter' -ForEach $script:StyleFiles {
        $text = Get-Content -LiteralPath $Path -Raw
        $styles = [regex]::Matches($text, '(?s)<Style\s+x:Key="(?<key>[^"]+)"\s+TargetType="Button"\s+BasedOn="\{StaticResource MahApps\.Styles\.Button\.Square(?<accent>\.Accent)?\}">\s*(?<first><Setter[^>]*>)')
        $bad = @()
        foreach ($m in $styles) {
            if ($m.Groups['first'].Value -match 'SidebarButtonTemplate') { continue }
            # A Square style is a normal button (neutral template) or a toolbar button that Set-ButtonTheme paints blue (accent template).
            $want = if ($m.Groups['accent'].Success) { 'SuiteAccentButtonTemplate' } else { 'Suite(Neutral|Accent)ButtonTemplate' }
            if ($m.Groups['first'].Value -notmatch ('Property="Template" Value="\{DynamicResource ' + $want + '\}"')) { $bad += ('{0}: first setter {1}' -f $m.Groups['key'].Value, $m.Groups['first'].Value) }
        }
        $bad | Should -BeNullOrEmpty
    }
}
