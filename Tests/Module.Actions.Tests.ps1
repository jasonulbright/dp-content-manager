#Requires -Modules Pester
# Write actions: redistribute, remove, validate. ConfigMgr and CIM are mocked.
# Covers success, failure, partial failure, progress callback, WhatIf and hostile input.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    $script:Stubs = New-CmStub -Name Remove-CMContentDistribution, Invoke-CMContentValidation -Definition {
        [CmdletBinding()]
        param([string]$PackageId, [string]$DistributionPointName, [switch]$Force)
    }
    Mock Write-Log {} -ModuleName DPContentMgrCommon
}

AfterAll {
    Remove-CmStub -Name $script:Stubs
    Remove-Variable -Name Fx -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Invoke-RedistributeContent' {
    BeforeEach {
        Mock Get-CMConnectionInfo { [PSCustomObject]@{ SiteCode = 'MCM'; SMSProvider = 'cm01.contoso.com' } } -ModuleName DPContentMgrCommon
        # One distribution point instance per server in $global:Fx (array of names).
        $global:Fx = @('dp1.contoso.com')
        Mock Get-CimInstance {
            foreach ($s in $global:Fx) {
                [PSCustomObject]@{ PackageID = 'MCM00001'; ServerNALPath = ('["Display=\\{0}\"]MSWNET:["SMS_SITE=MCM"]\\{0}\' -f $s) }
            }
        } -ModuleName DPContentMgrCommon
        Mock Set-CimInstance {} -ModuleName DPContentMgrCommon -RemoveParameterType InputObject
    }

    It 'refreshes the distribution point and reports success' {
        $r = @(Invoke-RedistributeContent -PackageID 'MCM00001' -DPName 'dp1.contoso.com')
        $r.Count | Should -Be 1
        $r[0].Success | Should -BeTrue
        $r[0].Error | Should -Be ''
        $r[0].PackageID | Should -Be 'MCM00001'
        $r[0].DPName | Should -Be 'dp1.contoso.com'
        Should -Invoke Set-CimInstance -ModuleName DPContentMgrCommon -Times 1 -Exactly -ParameterFilter { $Property.RefreshNow -eq $true }
    }

    It 'queries the connected site on the connected provider' {
        $null = Invoke-RedistributeContent -PackageID 'MCM00001' -DPName 'dp1.contoso.com'
        Should -Invoke Get-CimInstance -ModuleName DPContentMgrCommon -ParameterFilter { $Namespace -eq 'root\SMS\site_MCM' -and $ComputerName -eq 'cm01.contoso.com' }
    }

    It 'throws a plain message when no site is connected' {
        Mock Get-CMConnectionInfo { $null } -ModuleName DPContentMgrCommon
        { Invoke-RedistributeContent -PackageID 'MCM00001' -DPName 'dp1' } | Should -Throw 'Not connected to a CM site; call Connect-CMSite first.'
        Should -Invoke Set-CimInstance -ModuleName DPContentMgrCommon -Times 0 -Exactly
    }

    It 'reports a missing distribution point instance as a failure' {
        $global:Fx = @()
        $r = @(Invoke-RedistributeContent -PackageID 'MCM00001' -DPName 'dp1.contoso.com')
        $r[0].Success | Should -BeFalse
        $r[0].Error | Should -Be 'DP instance not found'
        Should -Invoke Set-CimInstance -ModuleName DPContentMgrCommon -Times 0 -Exactly
    }

    It 'does not touch dp10 when asked for dp1' {
        $global:Fx = @('dp10.contoso.com', 'dp1.contoso.com', 'dp11.contoso.com')
        $r = @(Invoke-RedistributeContent -PackageID 'MCM00001' -DPName 'dp1.contoso.com')
        $r[0].Success | Should -BeTrue
        Should -Invoke Set-CimInstance -ModuleName DPContentMgrCommon -Times 1 -Exactly -ParameterFilter { $InputObject.ServerNALPath -like '*\\dp1.contoso.com\*' }
    }

    It 'finds a DP given by its short name' {
        $global:Fx = @('dp1.contoso.com', 'dp10.contoso.com')
        $r = @(Invoke-RedistributeContent -PackageID 'MCM00001' -DPName 'dp1')
        $r[0].Success | Should -BeTrue
        Should -Invoke Set-CimInstance -ModuleName DPContentMgrCommon -Times 1 -Exactly -ParameterFilter { $InputObject.ServerNALPath -like '*\\dp1.contoso.com\*' }
    }

    It 'reports a short name that only prefixes another DP as not found' {
        $global:Fx = @('dp10.contoso.com')
        (@(Invoke-RedistributeContent -PackageID 'MCM00001' -DPName 'dp1'))[0].Error | Should -Be 'DP instance not found'
        Should -Invoke Set-CimInstance -ModuleName DPContentMgrCommon -Times 0 -Exactly
    }

    It 'escapes a quote in the package id so the WQL stays one literal' {
        $null = Invoke-RedistributeContent -PackageID "X' OR PackageID LIKE '%" -DPName 'dp1.contoso.com'
        # WQL escapes a quote with a backslash.
        Should -Invoke Get-CimInstance -ModuleName DPContentMgrCommon -ParameterFilter { $Query -like "*PackageID = 'X\' OR PackageID LIKE \'%'" }
    }

    It 'continues after a failure and reports each DP (partial failure)' {
        $global:Fx = @('dp1.contoso.com', 'dp2.contoso.com', 'dp3.contoso.com')
        Mock Set-CimInstance { if ($InputObject.ServerNALPath -like '*dp2*') { throw 'access denied' } } -ModuleName DPContentMgrCommon -RemoveParameterType InputObject
        $r = @(Invoke-RedistributeContent -PackageID 'MCM00001' -DPName 'dp1.contoso.com', 'dp2.contoso.com', 'dp3.contoso.com')
        ($r | ForEach-Object Success) -join ',' | Should -Be 'True,False,True'
        $r[1].Error | Should -BeLike '*access denied*'
        Should -Invoke Write-Log -ModuleName DPContentMgrCommon -ParameterFilter { $Level -eq 'ERROR' } -Times 1 -Exactly
    }

    It 'reports a query failure per DP' {
        Mock Get-CimInstance { throw 'RPC server unavailable' } -ModuleName DPContentMgrCommon
        $r = @(Invoke-RedistributeContent -PackageID 'MCM00001' -DPName 'a', 'b')
        $r.Count | Should -Be 2
        ($r | Where-Object Success).Count | Should -Be 0
        $r[0].Error | Should -BeLike '*RPC server unavailable*'
    }

    It 'calls the progress callback with index, total and name' {
        $global:Fx = @('a', 'b')
        $script:calls = [System.Collections.ArrayList]::new()
        $null = Invoke-RedistributeContent -PackageID 'MCM00001' -DPName 'a', 'b' -ProgressCallback { param($i, $n, $dp) [void]$script:calls.Add("$i/$n/$dp") }
        $script:calls -join ';' | Should -Be '0/2/a;1/2/b'
    }

    It 'does nothing under -WhatIf' {
        $r = @(Invoke-RedistributeContent -PackageID 'MCM00001' -DPName 'dp1.contoso.com' -WhatIf)
        $r.Count | Should -Be 0
        Should -Invoke Set-CimInstance -ModuleName DPContentMgrCommon -Times 0 -Exactly
    }

    It 'rejects a <Kind> argument' -ForEach @(
        @{ Kind = 'empty package id';       Params = @{ PackageID = '';  DPName = 'dp1' } }
        @{ Kind = 'null DP list';           Params = @{ PackageID = 'P'; DPName = $null } }
        @{ Kind = 'empty DP list';          Params = @{ PackageID = 'P'; DPName = @() } }
    ) {
        { Invoke-RedistributeContent @Params } | Should -Throw
    }

    It 'handles hostile DP names without a crash: <Name>' -ForEach @(
        @{ Name = "dp'1" }
        @{ Name = 'dp%1' }
        @{ Name = ('d' * 3000) }
        @{ Name = 'dp\..\x' }
        @{ Name = 'dp-日本' }
    ) {
        $global:Fx = @('dp1.contoso.com')
        $r = @(Invoke-RedistributeContent -PackageID 'MCM00001' -DPName $Name)
        $r.Count | Should -Be 1
        $r[0].Success | Should -BeFalse
        Should -Invoke Set-CimInstance -ModuleName DPContentMgrCommon -Times 0 -Exactly
    }
}

Describe 'Remove-ContentFromDP' {
    BeforeEach { Mock Remove-CMContentDistribution {} -ModuleName DPContentMgrCommon }

    It 'removes the content from each DP and reports success' {
        $r = @(Remove-ContentFromDP -PackageID 'MCM00001' -DPName 'a', 'b')
        ($r | ForEach-Object Success) -join ',' | Should -Be 'True,True'
        Should -Invoke Remove-CMContentDistribution -ModuleName DPContentMgrCommon -Times 1 -Exactly -ParameterFilter { $PackageId -eq 'MCM00001' -and $DistributionPointName -eq 'a' -and $Force }
        Should -Invoke Remove-CMContentDistribution -ModuleName DPContentMgrCommon -Times 1 -Exactly -ParameterFilter { $DistributionPointName -eq 'b' }
    }

    It 'continues after a failure and keeps the message (partial failure)' {
        Mock Remove-CMContentDistribution { if ($DistributionPointName -eq 'b') { throw 'content in use by a task sequence' } } -ModuleName DPContentMgrCommon
        $r = @(Remove-ContentFromDP -PackageID 'MCM00001' -DPName 'a', 'b', 'c')
        ($r | ForEach-Object Success) -join ',' | Should -Be 'True,False,True'
        $r[1].Error | Should -BeLike '*content in use by a task sequence*'
        $r[0].Error | Should -Be ''
    }

    It 'reports every DP as failed when the cmdlet always fails' {
        Mock Remove-CMContentDistribution { throw 'no connection' } -ModuleName DPContentMgrCommon
        @(Remove-ContentFromDP -PackageID 'P' -DPName 'a', 'b' | Where-Object Success).Count | Should -Be 0
    }

    It 'calls the progress callback before each removal' {
        $script:calls = [System.Collections.ArrayList]::new()
        $null = Remove-ContentFromDP -PackageID 'P' -DPName 'a', 'b', 'c' -ProgressCallback { param($i, $n, $dp) [void]$script:calls.Add("$i/$n/$dp") }
        $script:calls -join ';' | Should -Be '0/3/a;1/3/b;2/3/c'
    }

    It 'removes nothing under -WhatIf' {
        $r = @(Remove-ContentFromDP -PackageID 'MCM00001' -DPName 'a' -WhatIf)
        $r.Count | Should -Be 0
        Should -Invoke Remove-CMContentDistribution -ModuleName DPContentMgrCommon -Times 0 -Exactly
    }

    It 'rejects a <Kind> argument' -ForEach @(
        @{ Kind = 'empty package id'; Params = @{ PackageID = ''; DPName = 'a' } }
        @{ Kind = 'empty DP list';    Params = @{ PackageID = 'P'; DPName = @() } }
    ) {
        { Remove-ContentFromDP @Params } | Should -Throw
    }

    It 'passes a hostile DP name to the cmdlet as data, unchanged' {
        $null = Remove-ContentFromDP -PackageID 'P' -DPName "dp';--"
        Should -Invoke Remove-CMContentDistribution -ModuleName DPContentMgrCommon -ParameterFilter { $DistributionPointName -eq "dp';--" }
    }
}

Describe 'Invoke-ContentValidation' {
    BeforeEach { Mock Invoke-CMContentValidation {} -ModuleName DPContentMgrCommon }

    It 'starts validation on each DP and reports success' {
        $r = @(Invoke-ContentValidation -PackageID 'MCM00001' -DPName 'a', 'b')
        ($r | ForEach-Object Success) -join ',' | Should -Be 'True,True'
        Should -Invoke Invoke-CMContentValidation -ModuleName DPContentMgrCommon -Times 2 -Exactly -ParameterFilter { $PackageId -eq 'MCM00001' }
    }

    It 'continues after a failure (partial failure)' {
        Mock Invoke-CMContentValidation { if ($DistributionPointName -eq 'a') { throw 'dp offline' } } -ModuleName DPContentMgrCommon
        $r = @(Invoke-ContentValidation -PackageID 'P' -DPName 'a', 'b')
        ($r | ForEach-Object Success) -join ',' | Should -Be 'False,True'
        $r[0].Error | Should -BeLike '*dp offline*'
    }

    It 'calls the progress callback' {
        $script:calls = [System.Collections.ArrayList]::new()
        $null = Invoke-ContentValidation -PackageID 'P' -DPName 'a' -ProgressCallback { param($i, $n, $dp) [void]$script:calls.Add("$i/$n/$dp") }
        $script:calls -join ';' | Should -Be '0/1/a'
    }

    It 'validates nothing under -WhatIf' {
        @(Invoke-ContentValidation -PackageID 'P' -DPName 'a' -WhatIf).Count | Should -Be 0
        Should -Invoke Invoke-CMContentValidation -ModuleName DPContentMgrCommon -Times 0 -Exactly
    }

    It 'rejects an empty DP list' {
        { Invoke-ContentValidation -PackageID 'P' -DPName @() } | Should -Throw
    }
}

Describe 'Assert-ContentActionResult' {
    It 'passes when every result succeeded' {
        { Assert-ContentActionResult -Result @([PSCustomObject]@{ DPName = 'a'; Success = $true; Error = '' }) } | Should -Not -Throw
    }

    It 'passes for no result (a WhatIf run)' -ForEach @(
        @{ R = $null }
        @{ R = @() }
    ) {
        { Assert-ContentActionResult -Result $R } | Should -Not -Throw
    }

    It 'throws with the DP name and message of the failed result' {
        $rows = @([PSCustomObject]@{ DPName = 'a'; Success = $true; Error = '' }, [PSCustomObject]@{ DPName = 'b'; Success = $false; Error = 'DP instance not found' })
        { Assert-ContentActionResult -Result $rows } | Should -Throw '*b: DP instance not found*'
    }

    It 'names every failed DP' {
        $rows = @([PSCustomObject]@{ DPName = 'a'; Success = $false; Error = 'x' }, [PSCustomObject]@{ DPName = 'b'; Success = $false; Error = 'y' })
        { Assert-ContentActionResult -Result $rows } | Should -Throw '*a: x*b: y*'
    }

    It 'accepts the result as a pipeline value' {
        { [PSCustomObject]@{ DPName = 'a'; Success = $false; Error = 'bad' } | Assert-ContentActionResult } | Should -Throw '*a: bad*'
    }
}
