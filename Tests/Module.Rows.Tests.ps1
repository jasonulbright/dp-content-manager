#Requires -Modules Pester
# Row shaping for the main window: glyphs, size text, DP / content / issue rows.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    # The size text uses the current culture (1,024.0 MB). Fix the culture so the test gives one result on every host.
    $script:OldCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
    [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('en-US')
}

AfterAll { [System.Threading.Thread]::CurrentThread.CurrentCulture = $script:OldCulture }

Describe 'Get-RollupGlyph' {
    It 'returns <Glyph> for failed=<Failed> in-progress=<InProgress> total=<Total>' -ForEach @(
        @{ Failed = 0; InProgress = 0; Total = 0; Glyph = 0x22EF }
        @{ Failed = 0; InProgress = 0; Total = 5; Glyph = 0x2713 }
        @{ Failed = 1; InProgress = 0; Total = 5; Glyph = 0x2717 }
        @{ Failed = 1; InProgress = 4; Total = 5; Glyph = 0x2717 }
        @{ Failed = 0; InProgress = 2; Total = 5; Glyph = 0x22EF }
        @{ Failed = 3; InProgress = 0; Total = 0; Glyph = 0x22EF }
    ) {
        [int](Get-RollupGlyph -Failed $Failed -InProgress $InProgress -Total $Total) | Should -Be $Glyph
    }
}

Describe 'Get-StatusRowGlyph' {
    It 'returns <Glyph> for state <State>' -ForEach @(
        @{ State = 0; Glyph = 0x2713 }
        @{ State = 8; Glyph = 0x2713 }
        @{ State = 3; Glyph = 0x2717 }
        @{ State = 6; Glyph = 0x2717 }
        @{ State = 1; Glyph = 0x22EF }
        @{ State = 2; Glyph = 0x22EF }
        @{ State = 4; Glyph = 0x22EF }
        @{ State = 5; Glyph = 0x22EF }
        @{ State = 7; Glyph = 0x22EF }
        @{ State = 99; Glyph = 0x22EF }
        @{ State = -1; Glyph = 0x22EF }
    ) {
        [int](Get-StatusRowGlyph -State $State) | Should -Be $Glyph
    }
}

Describe 'Format-ContentSize' {
    It 'formats <KB> KB as <Text>' -ForEach @(
        @{ KB = 0;          Text = '0 KB' }
        @{ KB = 1;          Text = '1 KB' }
        @{ KB = 1023;       Text = '1023 KB' }
        @{ KB = 1024;       Text = '1.0 MB' }
        @{ KB = 1536;       Text = '1.5 MB' }
        @{ KB = 1048575;    Text = '1,024.0 MB' }
        @{ KB = 1048576;    Text = '1.00 GB' }
        @{ KB = 5242880;    Text = '5.00 GB' }
        @{ KB = 4294967296; Text = '4,096.00 GB' }
    ) {
        Format-ContentSize -SizeKB $KB | Should -Be $Text
    }

    It 'treats a negative size as bytes below one KB' {
        Format-ContentSize -SizeKB -5 | Should -Be '-5 KB'
    }
}

Describe 'ConvertTo-DPRows' {
    It 'joins the status and storage rollups by server name' {
        $dps = @([PSCustomObject]@{ ServerName = 'dp1'; SiteCode = 'MCM'; IsPullDP = $true })
        $by = @([PSCustomObject]@{ DPName = 'dp1'; TotalContent = 10; InstalledCount = 8; InProgressCount = 1; FailedCount = 1; PctComplete = '80%' })
        $stor = @([PSCustomObject]@{ DPName = 'dp1'; TotalSizeGB = 12.5 })
        $r = @(ConvertTo-DPRows -DPs $dps -ByDP $by -StorageByDP $stor)
        $r.Count | Should -Be 1
        $r[0].ServerName | Should -Be 'dp1'
        $r[0].SiteCode | Should -Be 'MCM'
        $r[0].IsPullDP | Should -BeTrue
        $r[0].TotalContent | Should -Be 10
        $r[0].InstalledCount | Should -Be 8
        $r[0].FailedCount | Should -Be 1
        $r[0].PctComplete | Should -Be '80%'
        $r[0].TotalSizeGB | Should -Be 12.5
        [int]$r[0].StatusGlyph | Should -Be 0x2717
    }

    It 'shows a DP without any status as empty with 0%' {
        $r = @(ConvertTo-DPRows -DPs @([PSCustomObject]@{ ServerName = 'dp9'; SiteCode = 'MCM'; IsPullDP = $false }) -ByDP @() -StorageByDP @())
        $r[0].TotalContent | Should -Be 0
        $r[0].PctComplete | Should -Be '0%'
        $r[0].TotalSizeGB | Should -Be 0
        [int]$r[0].StatusGlyph | Should -Be 0x22EF
    }

    It 'returns no rows for <Name> DPs' -ForEach @(
        @{ Name = 'null';  Dps = $null }
        @{ Name = 'empty'; Dps = @() }
    ) {
        @(ConvertTo-DPRows -DPs $Dps -ByDP $null -StorageByDP $null).Count | Should -Be 0
    }

    It 'matches server names without regard to case' {
        $r = @(ConvertTo-DPRows -DPs @([PSCustomObject]@{ ServerName = 'DP1'; SiteCode = 'S'; IsPullDP = $false }) `
                -ByDP @([PSCustomObject]@{ DPName = 'dp1'; TotalContent = 4; InstalledCount = 4; InProgressCount = 0; FailedCount = 0; PctComplete = '100%' }) -StorageByDP @())
        $r[0].TotalContent | Should -Be 4
        [int]$r[0].StatusGlyph | Should -Be 0x2713
    }

    It 'keeps one row per DP for 500 DPs' {
        $dps = @(1..500 | ForEach-Object { [PSCustomObject]@{ ServerName = "dp$_"; SiteCode = 'S'; IsPullDP = $false } })
        @(ConvertTo-DPRows -DPs $dps -ByDP @() -StorageByDP @()).Count | Should -Be 500
    }
}

Describe 'ConvertTo-ContentRows' {
    It 'joins the rollup by package id and formats the size' {
        $content = @([PSCustomObject]@{ ContentName = 'Office'; ContentType = 'Application'; PackageID = 'MCM00001'; SourceSize = 2097152 })
        $by = @([PSCustomObject]@{ PackageID = 'MCM00001'; TotalDPs = 4; InstalledCount = 4; InProgressCount = 0; FailedCount = 0; PctComplete = '100%' })
        $r = @(ConvertTo-ContentRows -Content $content -ByContent $by)
        $r[0].Name | Should -Be 'Office'
        $r[0].ContentType | Should -Be 'Application'
        $r[0].SizeDisplay | Should -Be '2.00 GB'
        $r[0].SourceSize | Should -Be 2097152
        $r[0].TotalDPs | Should -Be 4
        [int]$r[0].StatusGlyph | Should -Be 0x2713
    }

    It 'shows content that no DP holds as empty' {
        $r = @(ConvertTo-ContentRows -Content @([PSCustomObject]@{ ContentName = 'X'; ContentType = 'Package'; PackageID = 'P'; SourceSize = $null }) -ByContent @())
        $r[0].TotalDPs | Should -Be 0
        $r[0].SizeDisplay | Should -Be '0 KB'
        $r[0].PctComplete | Should -Be '0%'
    }

    It 'does not fail on a size above the 32-bit range' {
        $r = @(ConvertTo-ContentRows -Content @([PSCustomObject]@{ ContentName = 'Big'; ContentType = 'OS Image'; PackageID = 'P'; SourceSize = 3000000000 }) -ByContent @())
        $r[0].SizeDisplay | Should -Be '2,861.02 GB'
    }

    It 'returns no rows for <Name> content' -ForEach @(
        @{ Name = 'null';  C = $null }
        @{ Name = 'empty'; C = @() }
    ) {
        @(ConvertTo-ContentRows -Content $C -ByContent $null).Count | Should -Be 0
    }
}

Describe 'ConvertTo-StatusIssueRows' {
    It 'keeps only states that are not 0 or 8' {
        $status = @(0..8 | ForEach-Object { New-StatusRow "P$_" 'dp1' $_ })
        $r = @(ConvertTo-StatusIssueRows -Status $status -Content @())
        ($r | ForEach-Object { [int]$_.State }) -join ',' | Should -Be '1,2,3,4,5,6,7'
    }

    It 'names the content of a known package' {
        $r = @(ConvertTo-StatusIssueRows -Status @(New-StatusRow 'P1' 'dp1' 3 'Install Failed') -Content @([PSCustomObject]@{ PackageID = 'P1'; ContentName = 'Office'; ContentType = 'Application' }))
        $r[0].ContentName | Should -Be 'Office'
        $r[0].ContentType | Should -Be 'Application'
        $r[0].DPName | Should -Be 'dp1'
        $r[0].StatusName | Should -Be 'Install Failed'
        [int]$r[0].StatusGlyph | Should -Be 0x2717
    }

    It 'marks a package with no content object as orphaned' {
        $r = @(ConvertTo-StatusIssueRows -Status @(New-StatusRow 'GONE' 'dp1' 1) -Content @())
        $r[0].ContentName | Should -Be '(unknown / orphaned)'
        $r[0].ContentType | Should -Be '?'
        $r[0].PackageID | Should -Be 'GONE'
    }

    It 'returns no rows when everything is installed' {
        @(ConvertTo-StatusIssueRows -Status @((New-StatusRow 'P1' 'd' 0), (New-StatusRow 'P2' 'd' 8)) -Content @()).Count | Should -Be 0
    }

    It 'returns no rows for <Name> status' -ForEach @(
        @{ Name = 'null';  S = $null }
        @{ Name = 'empty'; S = @() }
    ) {
        @(ConvertTo-StatusIssueRows -Status $S -Content $null).Count | Should -Be 0
    }
}

Describe 'Refresh pipeline' {
    It 'carries the server name from the DP query into the grid row' {
        $null = New-CmStub -Name Get-CMDistributionPoint
        try {
            Mock Write-Log {} -ModuleName DPContentMgrCommon
            Mock Get-CMDistributionPoint { [PSCustomObject]@{ NetworkOSPath = '\\dp7.contoso.com'; SiteCode = 'MCM'; IsPullDP = $false; NALPath = 'n' } } -ModuleName DPContentMgrCommon
            $status = @(New-StatusRow 'P1' 'dp7.contoso.com' 0)
            $rows = @(ConvertTo-DPRows -DPs @(Get-AllDistributionPoints) -ByDP @(ConvertTo-DPStatusSummary -StatusRows $status) -StorageByDP @())
            $rows[0].ServerName | Should -Be 'dp7.contoso.com'
            $rows[0].TotalContent | Should -Be 1
        }
        finally { Remove-CmStub -Name Get-CMDistributionPoint }
    }
}
