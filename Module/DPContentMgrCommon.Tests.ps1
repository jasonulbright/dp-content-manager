#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x baseline for the DPContentMgrCommon shared module.

.DESCRIPTION
    Covers pure-logic exports: logging, status aggregation
    (ConvertTo-ContentStatusSummary, ConvertTo-DPStatusSummary), orphan
    detection (Find-OrphanedContent), DP storage analysis, and the
    plain-text status summary. CM-cmdlet integration (Connect-CMSite,
    Get-AllDistributionPoints, Get-AllContentObjects, Invoke-RedistributeContent,
    Remove-ContentFromDP, etc.) requires a live MECM site and is
    verified end-to-end on a CM-console-equipped client (CLIENT01)
    rather than mocked here.

.EXAMPLE
    Invoke-Pester .\DPContentMgrCommon.Tests.ps1
#>

BeforeAll {
    Import-Module "$PSScriptRoot\DPContentMgrCommon.psd1" -Force -DisableNameChecking

    function New-StatusRow {
        param([string]$PackageID, [string]$DPName, [int]$State, [string]$StatusName = 'Generated')
        [PSCustomObject]@{
            PackageID  = $PackageID
            DPName     = $DPName
            State      = $State
            StatusName = $StatusName
        }
    }

    function New-ContentRow {
        param([string]$PackageID, [string]$Name = 'Generated', [int]$SourceSize = 0)
        [PSCustomObject]@{
            PackageID  = $PackageID
            Name       = $Name
            SourceSize = $SourceSize
        }
    }
}

# ============================================================================
# Write-Log / Initialize-Logging
# ============================================================================

Describe 'Write-Log' {
    It 'writes formatted message to log file' {
        $logFile = Join-Path $TestDrive 'test.log'
        Initialize-Logging -LogPath $logFile
        Write-Log 'Hello world' -Quiet
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[INFO \] Hello world'
    }

    It 'tags WARN messages correctly' {
        $logFile = Join-Path $TestDrive 'warn.log'
        Initialize-Logging -LogPath $logFile
        Write-Log 'Slow' -Level WARN -Quiet
        (Get-Content -LiteralPath $logFile -Raw) | Should -Match '\[WARN \] Slow'
    }

    It 'tags ERROR messages correctly' {
        $logFile = Join-Path $TestDrive 'error.log'
        Initialize-Logging -LogPath $logFile
        Write-Log 'Boom' -Level ERROR -Quiet
        (Get-Content -LiteralPath $logFile -Raw) | Should -Match '\[ERROR\] Boom'
    }

    It 'accepts empty string message' {
        $logFile = Join-Path $TestDrive 'empty.log'
        Initialize-Logging -LogPath $logFile
        { Write-Log '' -Quiet } | Should -Not -Throw
    }
}

Describe 'Initialize-Logging' {
    It 'creates log file with header line' {
        $logFile = Join-Path $TestDrive 'init.log'
        Initialize-Logging -LogPath $logFile
        Test-Path -LiteralPath $logFile | Should -BeTrue
        (Get-Content -LiteralPath $logFile -Raw) | Should -Match '\[INFO \] === Log initialized ==='
    }

    It 'creates parent directories if missing' {
        $logFile = Join-Path $TestDrive 'sub\dir\deep.log'
        Initialize-Logging -LogPath $logFile
        Test-Path -LiteralPath $logFile | Should -BeTrue
    }

    It '-Attach preserves an externally-created log file' {
        $logFile = Join-Path $TestDrive 'attach.log'
        $sentinel = "[2026-05-02 00:00:00] [INFO ] Shell-managed header"
        Set-Content -LiteralPath $logFile -Value $sentinel -Encoding UTF8
        Initialize-Logging -LogPath $logFile -Attach
        Write-Log 'Module appended line' -Quiet
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match 'Shell-managed header'
        $content | Should -Match 'Module appended line'
    }
}

# ============================================================================
# Status / content row builders defined in the merged BeforeAll above.
# State codes: 0=Installed, 8=Installed, 1/2/7/4/5=InProgress, 3/6=Failed
# ============================================================================

Describe 'ConvertTo-ContentStatusSummary' {
    It 'aggregates per-package totals across many DPs' {
        $rows = @(
            (New-StatusRow -PackageID 'PKG001' -DPName 'DP01' -State 0)
            (New-StatusRow -PackageID 'PKG001' -DPName 'DP02' -State 0)
            (New-StatusRow -PackageID 'PKG001' -DPName 'DP03' -State 3)
            (New-StatusRow -PackageID 'PKG002' -DPName 'DP01' -State 0)
        )
        $summary = @(ConvertTo-ContentStatusSummary -StatusRows $rows)
        $pkg1 = $summary | Where-Object { $_.PackageID -eq 'PKG001' }
        $pkg1.TotalDPs   | Should -Be 3
        $pkg1.InstalledCount | Should -Be 2
        $pkg1.FailedCount    | Should -Be 1
        $pkg2 = $summary | Where-Object { $_.PackageID -eq 'PKG002' }
        $pkg2.TotalDPs       | Should -Be 1
        $pkg2.InstalledCount | Should -Be 1
    }

    It 'maps State 8 to Installed' {
        $rows = @(New-StatusRow -PackageID 'P' -DPName 'DP01' -State 8)
        $s = @(ConvertTo-ContentStatusSummary -StatusRows $rows)
        $s[0].InstalledCount | Should -Be 1
    }

    It 'maps States 1/2/7 to InProgress' {
        $rows = @(
            (New-StatusRow -PackageID 'P' -DPName 'A' -State 1)
            (New-StatusRow -PackageID 'P' -DPName 'B' -State 2)
            (New-StatusRow -PackageID 'P' -DPName 'C' -State 7)
        )
        $s = @(ConvertTo-ContentStatusSummary -StatusRows $rows)
        $s[0].InProgressCount | Should -Be 3
        $s[0].FailedCount     | Should -Be 0
    }

    It 'maps States 3/6 to Failed' {
        $rows = @(
            (New-StatusRow -PackageID 'P' -DPName 'A' -State 3)
            (New-StatusRow -PackageID 'P' -DPName 'B' -State 6)
        )
        $s = @(ConvertTo-ContentStatusSummary -StatusRows $rows)
        $s[0].FailedCount | Should -Be 2
    }
}

Describe 'ConvertTo-DPStatusSummary' {
    It 'aggregates per-DP totals across many content objects' {
        $rows = @(
            (New-StatusRow -PackageID 'P1' -DPName 'DP01' -State 0)
            (New-StatusRow -PackageID 'P2' -DPName 'DP01' -State 0)
            (New-StatusRow -PackageID 'P3' -DPName 'DP01' -State 3)
            (New-StatusRow -PackageID 'P1' -DPName 'DP02' -State 0)
        )
        $summary = @(ConvertTo-DPStatusSummary -StatusRows $rows)
        $dp1 = $summary | Where-Object { $_.DPName -eq 'DP01' }
        $dp1.TotalContent   | Should -Be 3
        $dp1.InstalledCount | Should -Be 2
        $dp1.FailedCount    | Should -Be 1
        $dp2 = $summary | Where-Object { $_.DPName -eq 'DP02' }
        $dp2.TotalContent   | Should -Be 1
        $dp2.InstalledCount | Should -Be 1
    }

    It 'skips rows with empty DPName' {
        $rows = @(
            (New-StatusRow -PackageID 'P' -DPName '' -State 0)
            (New-StatusRow -PackageID 'P' -DPName 'DP01' -State 0)
        )
        $summary = @(ConvertTo-DPStatusSummary -StatusRows $rows)
        $summary.Count | Should -Be 1
        $summary[0].DPName | Should -Be 'DP01'
    }
}

Describe 'Find-OrphanedContent' {
    It 'flags package IDs present on DPs but absent from the content object list' {
        $rows = @(
            (New-StatusRow -PackageID 'PKG001' -DPName 'DP01' -State 0)
            (New-StatusRow -PackageID 'PKG002' -DPName 'DP01' -State 0)
            (New-StatusRow -PackageID 'PKG_GHOST' -DPName 'DP01' -State 0)
        )
        $known = @(
            (New-ContentRow -PackageID 'PKG001')
            (New-ContentRow -PackageID 'PKG002')
        )
        $orphans = @(Find-OrphanedContent -StatusRows $rows -ContentObjects $known)
        $orphans.Count | Should -Be 1
        $orphans[0].PackageID | Should -Be 'PKG_GHOST'
    }

    It 'returns empty when every package on every DP has a matching content object' {
        $rows = @(
            (New-StatusRow -PackageID 'PKG001' -DPName 'DP01' -State 0)
            (New-StatusRow -PackageID 'PKG002' -DPName 'DP02' -State 0)
        )
        $known = @(
            (New-ContentRow -PackageID 'PKG001')
            (New-ContentRow -PackageID 'PKG002')
        )
        $orphans = @(Find-OrphanedContent -StatusRows $rows -ContentObjects $known)
        $orphans.Count | Should -Be 0
    }

    It 'reports the DP names where each orphan lives' {
        $rows = @(
            (New-StatusRow -PackageID 'GHOST' -DPName 'DP01' -State 0)
            (New-StatusRow -PackageID 'GHOST' -DPName 'DP02' -State 3)
            (New-StatusRow -PackageID 'GHOST' -DPName 'DP03' -State 0)
        )
        $orphans = @(Find-OrphanedContent -StatusRows $rows -ContentObjects @())
        $orphans[0].DPCount | Should -Be 3
        $orphans[0].DPNames | Should -Contain 'DP01'
        $orphans[0].DPNames | Should -Contain 'DP02'
        $orphans[0].DPNames | Should -Contain 'DP03'
    }
}

Describe 'Get-DPStorageAnalysis' {
    It 'sums SourceSize per DP from cross-referenced content objects (in GB)' {
        # SourceSize is in KB; output is in GB. Use whole-GB sizes so the
        # rounded output is exact: 1 GB = 1048576 KB, 2 GB = 2097152 KB.
        $oneGB = 1048576
        $twoGB = 2097152
        $rows = @(
            (New-StatusRow -PackageID 'P1' -DPName 'DP01' -State 0)
            (New-StatusRow -PackageID 'P2' -DPName 'DP01' -State 0)
            (New-StatusRow -PackageID 'P1' -DPName 'DP02' -State 0)
        )
        $known = @(
            (New-ContentRow -PackageID 'P1' -SourceSize $oneGB)
            (New-ContentRow -PackageID 'P2' -SourceSize $twoGB)
        )
        $analysis = @(Get-DPStorageAnalysis -StatusRows $rows -ContentObjects $known)
        $dp1 = $analysis | Where-Object { $_.DPName -eq 'DP01' }
        $dp1.TotalSizeGB  | Should -Be 3
        $dp1.ContentCount | Should -Be 2
        $dp2 = $analysis | Where-Object { $_.DPName -eq 'DP02' }
        $dp2.TotalSizeGB  | Should -Be 1
        $dp2.ContentCount | Should -Be 1
    }

    It 'counts failed content (states 3/6) into FailedCount' {
        $rows = @(
            (New-StatusRow -PackageID 'P1' -DPName 'DP01' -State 0)
            (New-StatusRow -PackageID 'P2' -DPName 'DP01' -State 3)
            (New-StatusRow -PackageID 'P3' -DPName 'DP01' -State 6)
        )
        $known = @((New-ContentRow -PackageID 'P1'),(New-ContentRow -PackageID 'P2'),(New-ContentRow -PackageID 'P3'))
        $analysis = @(Get-DPStorageAnalysis -StatusRows $rows -ContentObjects $known)
        $analysis[0].FailedCount | Should -Be 2
    }
}

Describe 'New-ContentStatusSummary' {
    It 'produces a header + total rows + failure section when failures exist' {
        $dt = New-Object System.Data.DataTable
        [void]$dt.Columns.Add('ContentName',[string])
        [void]$dt.Columns.Add('PackageID',[string])
        [void]$dt.Columns.Add('FailedCount',[int])
        [void]$dt.Rows.Add('App A','PKG001',0)
        [void]$dt.Rows.Add('App B','PKG002',2)
        $text = New-ContentStatusSummary -DataTable $dt
        $text | Should -Match '=== DP Content Status Summary ==='
        $text | Should -Match 'Total content objects: 2'
        $text | Should -Match 'Objects with failures: 1'
        $text | Should -Match 'App B \(PKG002\) - 2 DPs failed'
    }

    It 'omits the failure section when all rows have FailedCount = 0' {
        $dt = New-Object System.Data.DataTable
        [void]$dt.Columns.Add('ContentName',[string])
        [void]$dt.Columns.Add('PackageID',[string])
        [void]$dt.Columns.Add('FailedCount',[int])
        [void]$dt.Rows.Add('App A','PKG001',0)
        $text = New-ContentStatusSummary -DataTable $dt
        $text | Should -Match 'Objects with failures: 0'
        $text | Should -Not -Match '--- Failed Content ---'
    }
}
