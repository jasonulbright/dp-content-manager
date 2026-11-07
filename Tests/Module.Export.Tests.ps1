#Requires -Modules Pester
# Export-ContentStatusCsv, Export-ContentStatusHtml, New-ContentStatusSummary.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-DpcmModule
    Mock Write-Log {} -ModuleName DPContentMgrCommon
}

Describe 'Export-ContentStatusCsv' {
    It 'writes the columns in table order and one line per row' {
        $t = New-RollupTable -Column ContentName, PackageID, FailedCount -Row @(@('App A', 'P1', 0), @('App B', 'P2', 3))
        $path = Join-Path $TestDrive 'a.csv'
        Export-ContentStatusCsv -DataTable $t -OutputPath $path
        $rows = @(Import-Csv -LiteralPath $path)
        $rows.Count | Should -Be 2
        $rows[0].PSObject.Properties.Name -join ',' | Should -Be 'ContentName,PackageID,FailedCount'
        $rows[1].FailedCount | Should -Be '3'
    }

    It 'creates the missing parent folders' {
        $t = New-RollupTable -Column A -Row @(, @('x'))
        $path = Join-Path $TestDrive 'deep\er\still\a.csv'
        Export-ContentStatusCsv -DataTable $t -OutputPath $path
        Test-Path -LiteralPath $path | Should -BeTrue
    }

    It 'round-trips the hostile value <Name>' -ForEach @(
        @{ Name = 'comma';    Value = 'a,b' }
        @{ Name = 'quote';    Value = 'say "hi"' }
        @{ Name = 'newline';  Value = "line1`r`nline2" }
        @{ Name = 'unicode';  Value = 'Ünï 日本語' }
        @{ Name = 'html';     Value = '<script>alert(1)</script>' }
        @{ Name = 'long';     Value = ('x' * 20000) }
        @{ Name = 'spaces';   Value = '  padded  ' }
    ) {
        $t = New-RollupTable -Column V -Row @(, @($Value))
        $path = Join-Path $TestDrive "h-$Name.csv"
        Export-ContentStatusCsv -DataTable $t -OutputPath $path
        (Import-Csv -LiteralPath $path)[0].V | Should -Be $Value
    }

    It 'accepts a path with spaces and brackets' {
        $t = New-RollupTable -Column A -Row @(, @('x'))
        $path = Join-Path $TestDrive 'my report [1]\out file.csv'
        Export-ContentStatusCsv -DataTable $t -OutputPath $path
        Test-Path -LiteralPath $path | Should -BeTrue
    }

    It 'throws on an invalid path character' {
        $t = New-RollupTable -Column A -Row @(, @('x'))
        { Export-ContentStatusCsv -DataTable $t -OutputPath (Join-Path $TestDrive 'bad|name?.csv') } | Should -Throw
    }

    It 'requires a data table' {
        { Export-ContentStatusCsv -DataTable $null -OutputPath (Join-Path $TestDrive 'n.csv') } | Should -Throw
    }
}

Describe 'Export-ContentStatusHtml' {
    It 'writes one header cell per column and one row per data row' {
        $t = New-RollupTable -Column ContentName, FailedCount -Row @(@('App A', 0), @('App B', 2))
        $path = Join-Path $TestDrive 'a.html'
        Export-ContentStatusHtml -DataTable $t -OutputPath $path
        $html = Get-Content -LiteralPath $path -Raw
        ([regex]::Matches($html, '<th>')).Count | Should -Be 2
        ([regex]::Matches($html, '<tr>')).Count | Should -Be 3
        $html | Should -Match 'Rows: 2'
        $html | Should -Match '<title>\s*DP Content Status Report\s*</title>'
    }

    It 'encodes the value <Value>' -ForEach @(
        @{ Value = '<script>alert(1)</script>'; Expected = '&lt;script&gt;alert(1)&lt;/script&gt;' }
        @{ Value = 'A & B';                     Expected = 'A &amp; B' }
        @{ Value = '"quoted"';                  Expected = '&quot;quoted&quot;' }
        @{ Value = '</td><td>injected';         Expected = '&lt;/td&gt;&lt;td&gt;injected' }
    ) {
        $t = New-RollupTable -Column ContentName -Row @(, @($Value))
        $path = Join-Path $TestDrive 'enc.html'
        Export-ContentStatusHtml -DataTable $t -OutputPath $path
        $html = Get-Content -LiteralPath $path -Raw
        $html | Should -Match ([regex]::Escape("<td>$Expected</td>"))
    }

    It 'encodes a column name and the report title' {
        $t = New-RollupTable -Column '<b>Col</b>' -Row @(, @('x'))
        $path = Join-Path $TestDrive 'enc2.html'
        Export-ContentStatusHtml -DataTable $t -OutputPath $path -ReportTitle 'R&D <report>'
        $html = Get-Content -LiteralPath $path -Raw
        $html | Should -Match '<th>&lt;b&gt;Col&lt;/b&gt;</th>'
        $html | Should -Match '<title>\s*R&amp;D &lt;report&gt;\s*</title>'
        $html | Should -Match '<h1>R&amp;D &lt;report&gt;</h1>'
    }

    It 'marks <Column> = <Value> with class <Class>' -ForEach @(
        @{ Column = 'FailedCount';     Value = 3; Class = 'failed' }
        @{ Column = 'InProgressCount'; Value = 1; Class = 'inprog' }
        @{ Column = 'FailedCount';     Value = 0; Class = '' }
        @{ Column = 'InProgressCount'; Value = 0; Class = '' }
        @{ Column = 'FailedCount';     Value = 'n/a'; Class = '' }
        @{ Column = 'ContentName';     Value = 5; Class = '' }
    ) {
        $t = New-RollupTable -Column $Column -Row @(, @($Value))
        $path = Join-Path $TestDrive 'cls.html'
        Export-ContentStatusHtml -DataTable $t -OutputPath $path
        $html = Get-Content -LiteralPath $path -Raw
        if ($Class) { $html | Should -Match "<td class=`"$Class`">$Value</td>" }
        else { $html | Should -Match "<td>$([regex]::Escape([string]$Value))</td>" }
    }

    It 'writes a valid page for an empty table' {
        $t = New-RollupTable -Column A -Row @()
        $path = Join-Path $TestDrive 'empty.html'
        Export-ContentStatusHtml -DataTable $t -OutputPath $path
        $html = Get-Content -LiteralPath $path -Raw
        $html | Should -Match 'Rows: 0'
        $html | Should -Match '</html>\s*$'
    }

    It 'renders a null cell as empty text' {
        $t = New-RollupTable -Column A, B -Row @(, @('x', [DBNull]::Value))
        $path = Join-Path $TestDrive 'null.html'
        Export-ContentStatusHtml -DataTable $t -OutputPath $path
        (Get-Content -LiteralPath $path -Raw) | Should -Match '<td>x</td><td></td>'
    }

    It 'creates the missing parent folders' {
        $t = New-RollupTable -Column A -Row @(, @('x'))
        $path = Join-Path $TestDrive 'p\q\r.html'
        Export-ContentStatusHtml -DataTable $t -OutputPath $path
        Test-Path -LiteralPath $path | Should -BeTrue
    }
}

Describe 'New-ContentStatusSummary' {
    It 'reports zero objects and no failure section for an empty table' {
        $t = New-RollupTable -Column ContentName, PackageID, FailedCount -Row @()
        $s = New-ContentStatusSummary -DataTable $t
        $s | Should -Match 'Total content objects: 0'
        $s | Should -Match 'Objects with failures: 0'
        $s | Should -Not -Match 'Failed Content'
    }

    It 'lists up to 20 failed objects and counts the rest' {
        $rows = @(1..25 | ForEach-Object { , @("App $_", "P$_", 1) })
        $t = New-RollupTable -Column ContentName, PackageID, FailedCount -Row $rows
        $s = New-ContentStatusSummary -DataTable $t
        $s | Should -Match 'Objects with failures: 25'
        ([regex]::Matches($s, 'DPs failed')).Count | Should -Be 20
        $s | Should -Match '\.\.\. and 5 more'
    }

    It 'prints exactly 20 failures without a "more" line' {
        $rows = @(1..20 | ForEach-Object { , @("App $_", "P$_", 2) })
        $t = New-RollupTable -Column ContentName, PackageID, FailedCount -Row $rows
        (New-ContentStatusSummary -DataTable $t) | Should -Not -Match 'and \d+ more'
    }

    It 'names the content, id and failure count' {
        $t = New-RollupTable -Column ContentName, PackageID, FailedCount -Row @(, @('Office', 'MCM00042', 4))
        (New-ContentStatusSummary -DataTable $t) | Should -Match '  Office \(MCM00042\) - 4 DPs failed'
    }

    It 'fails with a message when the FailedCount column is missing' {
        $t = New-RollupTable -Column ContentName -Row @(, @('x'))
        { New-ContentStatusSummary -DataTable $t } | Should -Throw
    }
}
