#Requires -Modules Pester
# Status queries and aggregation: Get-BulkDistributionStatus, ConvertTo-*StatusSummary,
# Find-OrphanedContent, Get-DPStorageAnalysis.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-DpcmModule
    Mock Write-Log {} -ModuleName DPContentMgrCommon
}

AfterAll { Remove-Variable -Name Fx -Scope Global -ErrorAction SilentlyContinue }

Describe 'Get-BulkDistributionStatus' {
    It 'names state <State> as <Name>' -ForEach @(
        @{ State = 0; Name = 'Installed' }
        @{ State = 1; Name = 'Install Pending' }
        @{ State = 2; Name = 'Install Retrying' }
        @{ State = 3; Name = 'Install Failed' }
        @{ State = 4; Name = 'Removal Pending' }
        @{ State = 5; Name = 'Removal Retrying' }
        @{ State = 6; Name = 'Removal Failed' }
        @{ State = 7; Name = 'Content Validating' }
        @{ State = 8; Name = 'Content Valid' }
        @{ State = 9; Name = 'Unknown (9)' }
        @{ State = -1; Name = 'Unknown (-1)' }
        @{ State = 4294967295; Name = 'Unknown (4294967295)' }
    ) {
        $global:Fx = $State
        Mock Get-CimInstance {
            [PSCustomObject]@{ PackageID = 'MCM00001'; ServerNALPath = '["Display=\\dp1.contoso.com\"]MSWNET:["SMS_SITE=MCM"]\\dp1.contoso.com\'; State = $global:Fx; SourceVersion = 3; SummaryDate = $null }
        } -ModuleName DPContentMgrCommon
        $r = @(Get-BulkDistributionStatus -SMSProvider 'cm01' -SiteCode 'MCM')
        $r[0].StatusName | Should -Be $Name
        $r[0].State | Should -Be $State
    }

    It 'takes the DP name from ServerNALPath: <Nal>' -ForEach @(
        @{ Nal = '["Display=\\dp1.contoso.com\"]MSWNET:["SMS_SITE=MCM"]\\dp1.contoso.com\'; Expected = 'dp1.contoso.com' }
        @{ Nal = '\\dp2\SMS_DP$';                                                          Expected = 'dp2' }
        @{ Nal = 'no-unc-path';                                                            Expected = '' }
        @{ Nal = '';                                                                       Expected = '' }
        @{ Nal = $null;                                                                    Expected = '' }
    ) {
        $global:Fx = $Nal
        Mock Get-CimInstance { [PSCustomObject]@{ PackageID = 'P'; ServerNALPath = $global:Fx; State = 0; SourceVersion = 1; SummaryDate = $null } } -ModuleName DPContentMgrCommon
        (@(Get-BulkDistributionStatus -SMSProvider 'cm01' -SiteCode 'MCM'))[0].DPName | Should -Be $Expected
    }

    It 'queries the site namespace on the provider' {
        Mock Get-CimInstance { } -ModuleName DPContentMgrCommon
        $null = Get-BulkDistributionStatus -SMSProvider 'cm01.contoso.com' -SiteCode 'MCM'
        Should -Invoke Get-CimInstance -ModuleName DPContentMgrCommon -Times 1 -Exactly -ParameterFilter {
            $ComputerName -eq 'cm01.contoso.com' -and $Namespace -eq 'root\SMS\site_MCM' -and $Query -match 'SMS_PackageStatusDistPointsSummarizer'
        }
    }

    It 'returns nothing when no content is distributed' {
        Mock Get-CimInstance { } -ModuleName DPContentMgrCommon
        @(Get-BulkDistributionStatus -SMSProvider 'cm01' -SiteCode 'MCM').Count | Should -Be 0
    }

    It 'lets a provider failure reach the caller' {
        Mock Get-CimInstance { throw 'RPC server unavailable' } -ModuleName DPContentMgrCommon
        { Get-BulkDistributionStatus -SMSProvider 'cm01' -SiteCode 'MCM' } | Should -Throw '*RPC server unavailable*'
    }

    It 'requires <Missing>' -ForEach @(
        @{ Missing = 'SMSProvider'; Args = @{ SiteCode = 'MCM';  SMSProvider = '' } }
        @{ Missing = 'SiteCode';    Args = @{ SiteCode = '';     SMSProvider = 'cm01' } }
    ) {
        { Get-BulkDistributionStatus @Args } | Should -Throw
    }
}

Describe 'ConvertTo-ContentStatusSummary' {
    It 'counts state <State> as <Bucket>' -ForEach @(
        @{ State = 0; Bucket = 'InstalledCount' }
        @{ State = 8; Bucket = 'InstalledCount' }
        @{ State = 1; Bucket = 'InProgressCount' }
        @{ State = 2; Bucket = 'InProgressCount' }
        @{ State = 4; Bucket = 'InProgressCount' }
        @{ State = 5; Bucket = 'InProgressCount' }
        @{ State = 7; Bucket = 'InProgressCount' }
        @{ State = 3; Bucket = 'FailedCount' }
        @{ State = 6; Bucket = 'FailedCount' }
    ) {
        $r = @(ConvertTo-ContentStatusSummary -StatusRows @(New-StatusRow 'P1' 'dp1' $State))
        foreach ($b in 'InstalledCount', 'InProgressCount', 'FailedCount') {
            $r[0].$b | Should -Be $(if ($b -eq $Bucket) { 1 } else { 0 })
        }
        $r[0].TotalDPs | Should -Be 1
    }

    It 'counts an unknown state in the total only' {
        $r = @(ConvertTo-ContentStatusSummary -StatusRows @(New-StatusRow 'P1' 'dp1' 42))
        $r[0].TotalDPs | Should -Be 1
        ($r[0].InstalledCount + $r[0].InProgressCount + $r[0].FailedCount) | Should -Be 0
        $r[0].PctComplete | Should -Be '0%'
    }

    It 'rounds the percentage of <Installed> installed in <Total> to <Pct>' -ForEach @(
        @{ Installed = 1; Total = 3; Pct = '33.3%' }
        @{ Installed = 2; Total = 3; Pct = '66.7%' }
        @{ Installed = 3; Total = 3; Pct = '100%' }
        @{ Installed = 0; Total = 5; Pct = '0%' }
        @{ Installed = 1; Total = 8; Pct = '12.5%' }
    ) {
        $rows = @(1..$Total | ForEach-Object { New-StatusRow 'P1' "dp$_" $(if ($_ -le $Installed) { 0 } else { 3 }) })
        (@(ConvertTo-ContentStatusSummary -StatusRows $rows))[0].PctComplete | Should -Be $Pct
    }

    It 'accepts an empty row list (a site that has not distributed anything yet)' {
        { ConvertTo-ContentStatusSummary -StatusRows @() } | Should -Not -Throw
        @(ConvertTo-ContentStatusSummary -StatusRows @()).Count | Should -Be 0
    }

    It 'keeps separate totals per package' {
        $rows = @((New-StatusRow 'A' 'd1' 0), (New-StatusRow 'A' 'd2' 3), (New-StatusRow 'B' 'd1' 0))
        $r = @(ConvertTo-ContentStatusSummary -StatusRows $rows)
        ($r | Where-Object PackageID -eq 'A').TotalDPs | Should -Be 2
        ($r | Where-Object PackageID -eq 'B').TotalDPs | Should -Be 1
    }

    It 'aggregates 20000 rows' {
        $rows = @(1..20000 | ForEach-Object { New-StatusRow "P$($_ % 100)" "dp$($_ % 50)" ($_ % 9) })
        $r = @(ConvertTo-ContentStatusSummary -StatusRows $rows)
        $r.Count | Should -Be 100
        ($r | Measure-Object TotalDPs -Sum).Sum | Should -Be 20000
    }
}

Describe 'ConvertTo-DPStatusSummary' {
    It 'counts state <State> as <Bucket>' -ForEach @(
        @{ State = 0; Bucket = 'InstalledCount' }
        @{ State = 8; Bucket = 'InstalledCount' }
        @{ State = 1; Bucket = 'InProgressCount' }
        @{ State = 7; Bucket = 'InProgressCount' }
        @{ State = 5; Bucket = 'InProgressCount' }
        @{ State = 3; Bucket = 'FailedCount' }
        @{ State = 6; Bucket = 'FailedCount' }
    ) {
        $r = @(ConvertTo-DPStatusSummary -StatusRows @(New-StatusRow 'P1' 'dp1' $State))
        foreach ($b in 'InstalledCount', 'InProgressCount', 'FailedCount') {
            $r[0].$b | Should -Be $(if ($b -eq $Bucket) { 1 } else { 0 })
        }
    }

    It 'accepts an empty row list' {
        @(ConvertTo-DPStatusSummary -StatusRows @()).Count | Should -Be 0
    }

    It 'skips rows with a <Kind> DP name' -ForEach @(
        @{ Kind = 'empty'; Name = '' }
        @{ Kind = 'null';  Name = $null }
    ) {
        @(ConvertTo-DPStatusSummary -StatusRows @(New-StatusRow 'P1' $Name 0)).Count | Should -Be 0
    }

}

Describe 'Find-OrphanedContent' {
    It 'returns nothing for empty inputs' {
        @(Find-OrphanedContent -StatusRows @() -ContentObjects @()).Count | Should -Be 0
    }

    It 'flags every package when the content list is empty' {
        $r = @(Find-OrphanedContent -StatusRows @((New-StatusRow 'P1' 'd1' 0), (New-StatusRow 'P2' 'd1' 0)) -ContentObjects @())
        $r.Count | Should -Be 2
    }

    It 'counts DPs per orphan' {
        $rows = @((New-StatusRow 'P1' 'd1' 0), (New-StatusRow 'P1' 'd2' 0), (New-StatusRow 'P1' 'd3' 3))
        $r = @(Find-OrphanedContent -StatusRows $rows -ContentObjects @())
        $r[0].DPCount | Should -Be 3
        $r[0].DPNames | Should -Be @('d1', 'd2', 'd3')
    }

    It 'matches package ids case-insensitively' {
        $r = @(Find-OrphanedContent -StatusRows @(New-StatusRow 'mcm00001' 'd1' 0) -ContentObjects @([PSCustomObject]@{ PackageID = 'MCM00001' }))
        $r.Count | Should -Be 0
    }
}

Describe 'Get-DPStorageAnalysis' {
    It 'accepts empty inputs' {
        @(Get-DPStorageAnalysis -StatusRows @() -ContentObjects @()).Count | Should -Be 0
    }

    It 'treats missing, null and zero sizes as no size' -ForEach @(
        @{ Size = $null }
        @{ Size = 0 }
        @{ Size = '' }
    ) {
        $r = @(Get-DPStorageAnalysis -StatusRows @(New-StatusRow 'P1' 'd1' 0) -ContentObjects @([PSCustomObject]@{ PackageID = 'P1'; SourceSize = $Size }))
        $r[0].TotalSizeGB | Should -Be 0
        $r[0].ContentCount | Should -Be 1
    }

    It 'ignores status rows whose package is not in the content list' {
        $r = @(Get-DPStorageAnalysis -StatusRows @(New-StatusRow 'ORPHAN' 'd1' 0) -ContentObjects @([PSCustomObject]@{ PackageID = 'P1'; SourceSize = 1048576 }))
        $r[0].TotalSizeGB | Should -Be 0
        $r[0].ContentCount | Should -Be 1
    }

    It 'rounds KB to GB with two decimals: <KB> KB is <GB> GB' -ForEach @(
        @{ KB = 1048576;  GB = 1 }
        @{ KB = 524288;   GB = 0.5 }
        @{ KB = 1572864;  GB = 1.5 }
        @{ KB = 10485;    GB = 0.01 }
        @{ KB = 4000;     GB = 0 }
    ) {
        $r = @(Get-DPStorageAnalysis -StatusRows @(New-StatusRow 'P1' 'd1' 0) -ContentObjects @([PSCustomObject]@{ PackageID = 'P1'; SourceSize = $KB }))
        $r[0].TotalSizeGB | Should -Be $GB
    }

    It 'sorts the DPs by size, largest first' {
        $content = @([PSCustomObject]@{ PackageID = 'S'; SourceSize = 1048576 }, [PSCustomObject]@{ PackageID = 'L'; SourceSize = 10485760 })
        $rows = @((New-StatusRow 'S' 'small' 0), (New-StatusRow 'L' 'large' 0), (New-StatusRow 'S' 'mid' 0), (New-StatusRow 'L' 'mid' 0))
        (@(Get-DPStorageAnalysis -StatusRows $rows -ContentObjects $content) | ForEach-Object DPName) -join ',' | Should -Be 'mid,large,small'
    }

    It 'skips rows with an empty DP name' {
        @(Get-DPStorageAnalysis -StatusRows @(New-StatusRow 'P1' '' 0) -ContentObjects @([PSCustomObject]@{ PackageID = 'P1'; SourceSize = 5 })).Count | Should -Be 0
    }

    It 'counts failed states 3 and 6 only' -ForEach @(
        @{ State = 3; Failed = 1 }
        @{ State = 6; Failed = 1 }
        @{ State = 0; Failed = 0 }
        @{ State = 1; Failed = 0 }
        @{ State = 4; Failed = 0 }
    ) {
        $r = @(Get-DPStorageAnalysis -StatusRows @(New-StatusRow 'P1' 'd1' $State) -ContentObjects @([PSCustomObject]@{ PackageID = 'P1'; SourceSize = 1 }))
        $r[0].FailedCount | Should -Be $Failed
    }
}
