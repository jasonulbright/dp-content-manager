#Requires -Modules Pester
# Logic of start-dpcontentmgr.ps1. The script cannot be dot-sourced (it opens
# the window), so each function is cut out of its syntax tree and defined here,
# with plain objects in place of the window controls.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-DpcmModule
    $script:ShellPath = Join-Path $script:DpcmRoot 'start-dpcontentmgr.ps1'
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)

    $wanted = 'Get-FailedOpsForContent', 'Get-OpsForContentToAllDPs', 'Get-StatusFilterValue', 'Update-Filter',
        'Set-StatusText', 'Update-StatusBarSummary', 'Get-ActiveExportRows', 'Save-RowsAsCsv', 'Get-DpcmPreferences', 'Save-DpcmPreferences'
    foreach ($n in $wanted) {
        $def = $script:Ast.Find({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true)
        if (-not $def) { throw "function $n not found in the script" }
        . ([scriptblock]::Create($def.Extent.Text))
    }
}

Describe 'Get-FailedOpsForContent' {
    BeforeEach {
        $script:RawStatus = @((New-StatusRow 'P1' 'dp1' 3), (New-StatusRow 'P1' 'dp2' 0), (New-StatusRow 'P1' 'dp3' 6), (New-StatusRow 'P2' 'dp1' 3), (New-StatusRow 'P3' 'dp9' 1))
    }

    It 'returns one operation per failed DP (states 3 and 6)' {
        $ops = @(Get-FailedOpsForContent -ContentSelection @([PSCustomObject]@{ PackageID = 'P1'; Name = 'Office' }))
        ($ops.DPName -join ',') | Should -Be 'dp1,dp3'
        $ops[0].ContentName | Should -Be 'Office'
        $ops[0].PackageID | Should -Be 'P1'
    }

    It 'returns nothing for content with no failure' {
        @(Get-FailedOpsForContent -ContentSelection @([PSCustomObject]@{ PackageID = 'P3'; Name = 'x' })).Count | Should -Be 0
    }

    It 'covers several selected rows' {
        @(Get-FailedOpsForContent -ContentSelection @([PSCustomObject]@{ PackageID = 'P1'; Name = 'a' }, [PSCustomObject]@{ PackageID = 'P2'; Name = 'b' })).Count | Should -Be 3
    }

    It 'returns nothing when the status list is empty' {
        $script:RawStatus = @()
        @(Get-FailedOpsForContent -ContentSelection @([PSCustomObject]@{ PackageID = 'P1'; Name = 'a' })).Count | Should -Be 0
    }
}

Describe 'Get-OpsForContentToAllDPs' {
    It 'returns one operation for every DP that holds the content, whatever its state' {
        $script:RawStatus = @((New-StatusRow 'P1' 'dp1' 0), (New-StatusRow 'P1' 'dp2' 3), (New-StatusRow 'P2' 'dp1' 0))
        $ops = @(Get-OpsForContentToAllDPs -ContentSelection @([PSCustomObject]@{ PackageID = 'P1'; Name = 'Office' }))
        ($ops.DPName -join ',') | Should -Be 'dp1,dp2'
    }

    It 'returns nothing for content on no DP' {
        $script:RawStatus = @((New-StatusRow 'P2' 'dp1' 0))
        @(Get-OpsForContentToAllDPs -ContentSelection @([PSCustomObject]@{ PackageID = 'P1'; Name = 'x' })).Count | Should -Be 0
    }
}

Describe 'Get-StatusFilterValue' {
    It 'returns All when nothing is selected' {
        $cboStatusFilter = [PSCustomObject]@{ SelectedItem = $null }
        Get-StatusFilterValue | Should -Be 'All'
    }

    It 'returns the content of a combo box item' {
        Add-Type -AssemblyName PresentationFramework
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = 'Has failures'
        $cboStatusFilter = [PSCustomObject]@{ SelectedItem = $item }
        Get-StatusFilterValue | Should -Be 'Has failures'
    }

    It 'returns a plain string item as it is' {
        $cboStatusFilter = [PSCustomObject]@{ SelectedItem = 'Failed only' }
        Get-StatusFilterValue | Should -Be 'Failed only'
    }
}

Describe 'Update-Filter' {
    BeforeEach {
        function New-Dp { param($Name, $Failed = 0, $Progress = 0) [PSCustomObject]@{ ServerName = $Name; FailedCount = $Failed; InProgressCount = $Progress } }
        $script:ActiveView = 'DPs'
        $txtFilter = [PSCustomObject]@{ Text = '' }
        $cboStatusFilter = [PSCustomObject]@{ SelectedItem = 'All' }
        $gridDPs = [PSCustomObject]@{ ItemsSource = $null }
        $gridContent = [PSCustomObject]@{ ItemsSource = $null }
        $gridStatusIssues = [PSCustomObject]@{ ItemsSource = $null }
        $txtIssueCount = [PSCustomObject]@{ Text = '' }
        $script:DPRows = @((New-Dp 'DP-Alpha'), (New-Dp 'dp-beta' -Failed 2), (New-Dp 'dp-gamma' -Progress 1))
        $script:ContentRows = @(
            [PSCustomObject]@{ Name = 'Office 365'; PackageID = 'MCM00001'; ContentType = 'Application' },
            [PSCustomObject]@{ Name = 'Drivers'; PackageID = 'MCM00002'; ContentType = 'Driver Package' },
            [PSCustomObject]@{ Name = $null; PackageID = 'MCM00003'; ContentType = 'Package' })
        $script:StatusIssueRows = @(
            [PSCustomObject]@{ DPName = 'dp1'; ContentName = 'Office'; State = 3 },
            [PSCustomObject]@{ DPName = 'dp2'; ContentName = 'Drivers'; State = 1 },
            [PSCustomObject]@{ DPName = 'dp3'; ContentName = $null; State = 7 },
            [PSCustomObject]@{ DPName = 'dp4'; ContentName = 'Boot'; State = 6 })
    }

    It 'DPs view: text "<Text>" and status "<Status>" show <Expected>' -ForEach @(
        @{ Text = '';          Status = 'All';             Expected = 'DP-Alpha,dp-beta,dp-gamma' }
        @{ Text = 'ALPHA';     Status = 'All';             Expected = 'DP-Alpha' }
        @{ Text = '  beta  ';  Status = 'All';             Expected = 'dp-beta' }
        @{ Text = 'zzz';       Status = 'All';             Expected = '' }
        @{ Text = '';          Status = 'Healthy only';    Expected = 'DP-Alpha' }
        @{ Text = '';          Status = 'Has failures';    Expected = 'dp-beta' }
        @{ Text = '';          Status = 'Has in-progress'; Expected = 'dp-gamma' }
        @{ Text = 'dp-';       Status = 'Has failures';    Expected = 'dp-beta' }
        @{ Text = '';          Status = 'Unknown value';   Expected = 'DP-Alpha,dp-beta,dp-gamma' }
    ) {
        $txtFilter.Text = $Text
        $cboStatusFilter.SelectedItem = $Status
        Update-Filter
        (@($gridDPs.ItemsSource) | ForEach-Object ServerName) -join ',' | Should -Be $Expected
    }

    It 'Content view: text "<Text>" and type "<Type>" show <Expected>' -ForEach @(
        @{ Text = '';         Type = 'All';            Expected = 'MCM00001,MCM00002,MCM00003' }
        @{ Text = 'office';   Type = 'All';            Expected = 'MCM00001' }
        @{ Text = 'MCM00002'; Type = 'All';            Expected = 'MCM00002' }
        @{ Text = 'mcm0000';  Type = 'Application';    Expected = 'MCM00001' }
        @{ Text = '';         Type = 'Driver Package'; Expected = 'MCM00002' }
        @{ Text = '';         Type = 'Nope';           Expected = '' }
        @{ Text = '00003';    Type = 'All';            Expected = 'MCM00003' }
    ) {
        $script:ActiveView = 'Content'
        $txtFilter.Text = $Text
        $cboStatusFilter.SelectedItem = $Type
        Update-Filter
        (@($gridContent.ItemsSource) | ForEach-Object PackageID) -join ',' | Should -Be $Expected
    }

    It 'Status Issues view: text "<Text>" and filter "<Filter>" show <Expected> and count text <Count>' -ForEach @(
        @{ Text = '';       Filter = 'All';              Expected = 'dp1,dp2,dp3,dp4'; Count = '4 issue rows' }
        @{ Text = '';       Filter = 'Failed only';      Expected = 'dp1,dp4';         Count = '2 issue rows' }
        @{ Text = '';       Filter = 'In progress only'; Expected = 'dp2,dp3';         Count = '2 issue rows' }
        @{ Text = 'office'; Filter = 'All';              Expected = 'dp1';             Count = '1 issue rows' }
        @{ Text = 'dp3';    Filter = 'All';              Expected = 'dp3';             Count = '1 issue rows' }
        @{ Text = 'none';   Filter = 'All';              Expected = '';                Count = '0 issue rows' }
    ) {
        $script:ActiveView = 'Status Issues'
        $txtFilter.Text = $Text
        $cboStatusFilter.SelectedItem = $Filter
        Update-Filter
        (@($gridStatusIssues.ItemsSource) | ForEach-Object DPName) -join ',' | Should -Be $Expected
        $txtIssueCount.Text | Should -Be $Count
    }

    It 'treats a filter with regex and wildcard characters as plain text' -ForEach @(
        @{ Text = '.*' }
        @{ Text = '[' }
        @{ Text = '(' }
        @{ Text = '%' }
        @{ Text = "'" }
    ) {
        $txtFilter.Text = $Text
        { Update-Filter } | Should -Not -Throw
        @($gridDPs.ItemsSource).Count | Should -Be 0
    }

    It 'handles a row whose name is null' {
        $script:ActiveView = 'Content'
        $txtFilter.Text = 'x'
        { Update-Filter } | Should -Not -Throw
    }
}

Describe 'Update-StatusBarSummary' {
    BeforeEach {
        function Set-StatusText { param([string]$Text) $script:StatusText = $Text }
        $global:Prefs = @{ SiteCode = 'MCM'; SMSProvider = 'cm01' }
        $script:IsConnectedFromBg = $false
        $script:RawDPs = @(); $script:RawContent = @(); $script:StatusIssueRows = @(); $script:LastRefreshTime = $null
        $script:StatusText = $null
    }

    AfterAll { Remove-Variable -Name Prefs -Scope Global -ErrorAction SilentlyContinue }

    It 'says what to do when <Case>' -ForEach @(
        @{ Case = 'the site code is missing';   Site = '';    Provider = 'cm01'; Expected = 'Open Options to configure site code and SMS provider' }
        @{ Case = 'the provider is missing';    Site = 'MCM'; Provider = '';     Expected = 'Open Options to configure site code and SMS provider' }
        @{ Case = 'both are missing';           Site = '';    Provider = '';     Expected = 'Open Options to configure site code and SMS provider' }
        @{ Case = 'both are set, not connected'; Site = 'MCM'; Provider = 'cm01'; Expected = 'Ready. Click Refresh.' }
    ) {
        $global:Prefs = @{ SiteCode = $Site; SMSProvider = $Provider }
        Update-StatusBarSummary
        $script:StatusText | Should -Be $Expected
    }

    It 'names the site when connected' {
        $script:IsConnectedFromBg = $true
        Update-StatusBarSummary
        $script:StatusText | Should -Be 'Connected to MCM'
    }

    It 'shows the counts and the refresh time' {
        $script:IsConnectedFromBg = $true
        $script:RawDPs = @(1, 2); $script:RawContent = @(1, 2, 3); $script:StatusIssueRows = @(1)
        $script:LastRefreshTime = [datetime]'2026-11-06T14:05:09'
        Update-StatusBarSummary
        $script:StatusText | Should -Be 'Connected to MCM   |   2 DPs   |   3 content   |   1 issues   |   last refresh 14:05:09'
    }

    It 'leaves a count out when it is zero' {
        $script:RawDPs = @(1)
        Update-StatusBarSummary
        $script:StatusText | Should -Be 'Ready. Click Refresh.   |   1 DPs'
    }
}

Describe 'Get-ActiveExportRows' {
    BeforeEach {
        $gridDPs = [PSCustomObject]@{ ItemsSource = 'dp rows' }
        $gridContent = [PSCustomObject]@{ ItemsSource = 'content rows' }
        $gridStatusIssues = [PSCustomObject]@{ ItemsSource = 'issue rows' }
    }

    It 'exports the <View> grid as <Name>' -ForEach @(
        @{ View = 'DPs';           Name = 'DPs';          Rows = 'dp rows' }
        @{ View = 'Content';       Name = 'Content';      Rows = 'content rows' }
        @{ View = 'Status Issues'; Name = 'StatusIssues'; Rows = 'issue rows' }
    ) {
        $script:ActiveView = $View
        $r = Get-ActiveExportRows
        $r.Name | Should -Be $Name
        $r.Rows | Should -Be $Rows
    }

    It 'returns nothing for an unknown view' {
        $script:ActiveView = 'Nope'
        Get-ActiveExportRows | Should -BeNullOrEmpty
    }
}

Describe 'Save-RowsAsCsv' {
    It 'writes the rows with their properties' {
        $path = Join-Path $TestDrive 'rows.csv'
        Save-RowsAsCsv -Rows @([PSCustomObject]@{ A = 1; B = 'x,y' }, [PSCustomObject]@{ A = 2; B = 'z' }) -OutputPath $path
        $r = @(Import-Csv -LiteralPath $path)
        $r.Count | Should -Be 2
        $r[0].B | Should -Be 'x,y'
    }
}

Describe 'Preferences file' {
    BeforeEach { $global:PrefsPath = Join-Path $TestDrive ('prefs-{0}.json' -f [guid]::NewGuid().ToString('N')) }
    AfterAll { Remove-Variable -Name PrefsPath -Scope Global -ErrorAction SilentlyContinue }

    It 'gives the defaults on first run, when the file does not exist' {
        $p = Get-DpcmPreferences
        $p.DarkMode | Should -BeTrue
        $p.SiteCode | Should -Be ''
        $p.SMSProvider | Should -Be ''
        Test-Path -LiteralPath $global:PrefsPath | Should -BeFalse
    }

    It 'gives the defaults for a <Case> file' -ForEach @(
        @{ Case = 'corrupt';        Content = '{ not json' }
        @{ Case = 'empty';          Content = '' }
        @{ Case = 'array';          Content = '[1,2,3]' }
        @{ Case = 'binary';         Content = ([string][char]0 * 10) }
    ) {
        Set-Content -LiteralPath $global:PrefsPath -Value $Content
        $p = Get-DpcmPreferences
        $p.DarkMode | Should -BeTrue
        $p.SiteCode | Should -Be ''
    }

    It 'keeps a valid value and defaults the rest' {
        Set-Content -LiteralPath $global:PrefsPath -Value '{ "SiteCode": "MCM" }'
        $p = Get-DpcmPreferences
        $p.SiteCode | Should -Be 'MCM'
        $p.SMSProvider | Should -Be ''
        $p.DarkMode | Should -BeTrue
    }

    It 'ignores a key it does not know' {
        Set-Content -LiteralPath $global:PrefsPath -Value '{ "SiteCode": "MCM", "Evil": "x" }'
        (Get-DpcmPreferences).ContainsKey('Evil') | Should -BeFalse
    }

    It 'saves and reads back a round trip' {
        Save-DpcmPreferences -Prefs @{ DarkMode = $false; SiteCode = 'MCM'; SMSProvider = 'cm01.contoso.com' }
        $p = Get-DpcmPreferences
        $p.DarkMode | Should -BeFalse
        $p.SiteCode | Should -Be 'MCM'
        $p.SMSProvider | Should -Be 'cm01.contoso.com'
    }

    It 'round-trips the value <Value>' -ForEach @(
        @{ Value = 'Ünï-日本' }
        @{ Value = 'a "quoted" \ back\slash' }
        @{ Value = ('x' * 5000) }
    ) {
        Save-DpcmPreferences -Prefs @{ DarkMode = $true; SiteCode = $Value; SMSProvider = '' }
        (Get-DpcmPreferences).SiteCode | Should -Be $Value
    }

    It 'creates the folder of the file when it is missing' {
        $global:PrefsPath = Join-Path $TestDrive 'a\b\prefs.json'
        Save-DpcmPreferences -Prefs @{ DarkMode = $true; SiteCode = 'X'; SMSProvider = '' }
        Test-Path -LiteralPath $global:PrefsPath | Should -BeTrue
    }

    It 'does not throw when the file cannot be written' {
        $global:PrefsPath = Join-Path $TestDrive 'dir-in-the-way'
        New-Item -ItemType Directory -Path $global:PrefsPath | Out-Null
        { Save-DpcmPreferences -Prefs @{ DarkMode = $true; SiteCode = 'X'; SMSProvider = '' } } | Should -Not -Throw
    }

    It 'reads a file with spaces and brackets in its path' {
        $global:PrefsPath = Join-Path $TestDrive 'my prefs [1]\p.json'
        Save-DpcmPreferences -Prefs @{ DarkMode = $true; SiteCode = 'SP'; SMSProvider = '' }
        (Get-DpcmPreferences).SiteCode | Should -Be 'SP'
    }
}
Describe 'Script contract' {
    It 'passes every action result through Assert-ContentActionResult' {
        $text = Get-Content -LiteralPath $script:ShellPath -Raw
        $calls = [regex]::Matches($text, 'param\(\$op\) (?<body>[^\r\n]+)')
        $calls.Count | Should -Be 5
        foreach ($c in $calls) { $c.Groups['body'].Value | Should -Match '^Assert-ContentActionResult -Result \(' }
    }

    It 'calls only commands that exist' {
        $own = @($script:Ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object Name)
        $called = @($script:Ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.CommandAst] }, $true) |
                ForEach-Object { $_.GetCommandName() } | Where-Object { $_ } | Sort-Object -Unique)
        Import-Module (Join-Path $script:DpcmRoot 'Lib\SuiteCommon\SuiteCommon.psd1') -Global -DisableNameChecking -ErrorAction SilentlyContinue
        $missing = @($called | Where-Object { $_ -notin $own -and -not (Get-Command $_ -ErrorAction SilentlyContinue) })
        # ConfigMgr cmdlets and Windows PowerShell 5.1-only commands are not on every machine.
        $missing | Where-Object { $_ -notmatch '^(Get|Set|New|Remove|Start|Invoke)-CM' } | Should -BeNullOrEmpty
    }
}
