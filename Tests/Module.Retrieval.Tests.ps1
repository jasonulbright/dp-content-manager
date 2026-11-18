#Requires -Modules Pester
# Content retrieval: every Get-* function, with the ConfigMgr cmdlets mocked.
# Fixture values reach the mock bodies through $global:Fx, because a mock body
# runs in the module scope and cannot see test-scope variables.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    $script:Stubs = New-CmStub -Name Get-CMDistributionPoint, Get-CMApplication, Get-CMPackage,
        Get-CMSoftwareUpdateDeploymentPackage, Get-CMBootImage, Get-CMOperatingSystemImage,
        Get-CMDriverPackage, Get-CMTaskSequence, Get-CMDistributionStatus
    Mock Write-Log {} -ModuleName DPContentMgrCommon
}

AfterAll {
    Remove-CmStub -Name $script:Stubs
    Remove-Variable -Name Fx -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Get-AllDistributionPoints' {
    It 'normalizes the server name from NetworkOSPath <Path>' -ForEach @(
        @{ Path = '\\dp1.contoso.com';    Expected = 'dp1.contoso.com' }
        @{ Path = ('\\DP-' + [char]0x00DC + 'n' + [char]0x00EF + '.contoso.com'); Expected = ('DP-' + [char]0x00DC + 'n' + [char]0x00EF + '.contoso.com') }
        @{ Path = '\\a';                  Expected = 'a' }
        @{ Path = 'dp1.contoso.com';      Expected = '' }
        @{ Path = '';                     Expected = '' }
        @{ Path = $null;                  Expected = '' }
    ) {
        $global:Fx = $Path
        Mock Get-CMDistributionPoint { [PSCustomObject]@{ NetworkOSPath = $global:Fx; SiteCode = 'MCM'; IsPullDP = $false; NALPath = 'x' } } -ModuleName DPContentMgrCommon
        $r = @(Get-AllDistributionPoints)
        $r.Count | Should -Be 1
        $r[0].Name | Should -Be $Expected
    }

    It 'exposes ServerName, the property the main window reads' {
        Mock Get-CMDistributionPoint { [PSCustomObject]@{ NetworkOSPath = '\\dp1.contoso.com'; SiteCode = 'MCM'; IsPullDP = $true; NALPath = 'n' } } -ModuleName DPContentMgrCommon
        $r = @(Get-AllDistributionPoints)
        $r[0].ServerName | Should -Be 'dp1.contoso.com'
        $r[0].Name | Should -Be 'dp1.contoso.com'
    }

    It 'converts IsPullDP <Value> to a boolean' -ForEach @(
        @{ Value = $true;  Expected = $true }
        @{ Value = $false; Expected = $false }
        @{ Value = $null;  Expected = $false }
        @{ Value = 1;      Expected = $true }
        @{ Value = 0;      Expected = $false }
    ) {
        $global:Fx = $Value
        Mock Get-CMDistributionPoint { [PSCustomObject]@{ NetworkOSPath = '\\d'; SiteCode = 'S'; IsPullDP = $global:Fx; NALPath = '' } } -ModuleName DPContentMgrCommon
        (@(Get-AllDistributionPoints))[0].IsPullDP | Should -Be $Expected
    }

    It 'returns nothing for a site with no distribution points' {
        Mock Get-CMDistributionPoint { } -ModuleName DPContentMgrCommon
        @(Get-AllDistributionPoints).Count | Should -Be 0
    }

    It 'returns every distribution point in order' {
        Mock Get-CMDistributionPoint {
            1..3 | ForEach-Object { [PSCustomObject]@{ NetworkOSPath = "\\dp$_"; SiteCode = 'MCM'; IsPullDP = $false; NALPath = "n$_" } }
        } -ModuleName DPContentMgrCommon
        (@(Get-AllDistributionPoints) | ForEach-Object Name) -join ',' | Should -Be 'dp1,dp2,dp3'
    }

    It 'lets a cmdlet failure reach the caller (ErrorAction Stop)' {
        Mock Get-CMDistributionPoint { throw 'provider unreachable' } -ModuleName DPContentMgrCommon
        { Get-AllDistributionPoints } | Should -Throw '*provider unreachable*'
    }
}

Describe 'Typed content getters <Fn>' -ForEach @(
    @{ Fn = 'Get-ContentPackages';       Cmd = 'Get-CMPackage';                         Type = 'Package' }
    @{ Fn = 'Get-ContentSUDPs';          Cmd = 'Get-CMSoftwareUpdateDeploymentPackage'; Type = 'SU Deployment Pkg' }
    @{ Fn = 'Get-ContentBootImages';     Cmd = 'Get-CMBootImage';                       Type = 'Boot Image' }
    @{ Fn = 'Get-ContentOSImages';       Cmd = 'Get-CMOperatingSystemImage';            Type = 'OS Image' }
    @{ Fn = 'Get-ContentDriverPackages'; Cmd = 'Get-CMDriverPackage';                   Type = 'Driver Package' }
) {
    It 'maps name, id, size and source date' {
        Mock -CommandName $Cmd -ModuleName DPContentMgrCommon -MockWith {
            [PSCustomObject]@{ Name = 'Pkg A'; PackageID = 'MCM00001'; PackageSize = 2048; SourceDate = '2026-01-02' }
        }
        $r = @(& $Fn)
        $r.Count | Should -Be 1
        $r[0].ContentName | Should -Be 'Pkg A'
        $r[0].ContentType | Should -Be $Type
        $r[0].PackageID | Should -Be 'MCM00001'
        $r[0].ObjectID | Should -Be 'MCM00001'
        $r[0].SourceSize | Should -Be 2048
        $r[0].SourceVersion | Should -Be '2026-01-02'
    }

    It 'returns nothing when the site has none' {
        Mock -CommandName $Cmd -ModuleName DPContentMgrCommon -MockWith { }
        @(& $Fn).Count | Should -Be 0
    }

    It 'keeps the hostile name verbatim' -ForEach @(
        @{ Name = "O'Brien ""Quoted"" <b>&amp;</b>" }
        @{ Name = ('x' * 5000) }
        @{ Name = ('' + [char]0x00DC + 'n' + [char]0x00EF + 'c' + [char]0x00F6 + 'd' + [char]0x00E9 + ' ' + [char]0x65E5 + [char]0x672C + [char]0x8A9E + ' ' + [char]::ConvertFromUtf32(0x1F642)) }
        @{ Name = '..\..\Windows\System32' }
        @{ Name = '' }
    ) {
        $global:Fx = $Name
        Mock -CommandName $Cmd -ModuleName DPContentMgrCommon -MockWith { [PSCustomObject]@{ Name = $global:Fx; PackageID = 'P1'; PackageSize = 1; SourceDate = $null } }
        (@(& $Fn))[0].ContentName | Should -Be $Name
    }

    It 'handles many objects' {
        Mock -CommandName $Cmd -ModuleName DPContentMgrCommon -MockWith {
            1..2000 | ForEach-Object { [PSCustomObject]@{ Name = "n$_"; PackageID = "P$_"; PackageSize = $_; SourceDate = $null } }
        }
        @(& $Fn).Count | Should -Be 2000
    }
}

Describe 'Get-ContentApplications' {
    It 'maps display name, CI_ID and package size' {
        Mock Get-CMApplication {
            [PSCustomObject]@{ LocalizedDisplayName = 'App A'; PackageID = 'MCM0000A'; PackageSize = 10; SourceDate = 'd'; CI_ID = 16777300 }
        } -ModuleName DPContentMgrCommon
        $r = @(Get-ContentApplications)
        $r[0].ContentName | Should -Be 'App A'
        $r[0].ContentType | Should -Be 'Application'
        $r[0].ObjectID | Should -Be 16777300
        $r[0].PackageID | Should -Be 'MCM0000A'
    }

    It 'returns nothing when there are no applications' {
        Mock Get-CMApplication { } -ModuleName DPContentMgrCommon
        @(Get-ContentApplications).Count | Should -Be 0
    }
}

Describe 'Get-ContentTaskSequenceRefs' {
    It 'lists the package ids a task sequence references' {
        Mock Get-CMTaskSequence {
            [PSCustomObject]@{ Name = 'TS'; PackageID = 'MCM00010'; PackageSize = 0; SourceDate = $null
                References = @([PSCustomObject]@{ Package = 'MCM00001' }, [PSCustomObject]@{ Package = 'MCM00002' }) }
        } -ModuleName DPContentMgrCommon
        $r = @(Get-ContentTaskSequenceRefs)
        $r[0].ContentType | Should -Be 'Task Sequence'
        $r[0].ReferencedIDs | Should -Be @('MCM00001', 'MCM00002')
    }

    It 'gives an empty reference list when References is <Name>' -ForEach @(
        @{ Name = 'null';  Refs = $null }
        @{ Name = 'empty'; Refs = @() }
    ) {
        $global:Fx = $Refs
        Mock Get-CMTaskSequence { [PSCustomObject]@{ Name = 'TS'; PackageID = 'MCM00010'; PackageSize = 0; SourceDate = $null; References = $global:Fx } } -ModuleName DPContentMgrCommon
        @((Get-ContentTaskSequenceRefs)[0].ReferencedIDs).Count | Should -Be 0
    }

    It 'returns nothing when there are no task sequences' {
        Mock Get-CMTaskSequence { } -ModuleName DPContentMgrCommon
        @(Get-ContentTaskSequenceRefs).Count | Should -Be 0
    }
}

Describe 'Get-AllContentObjects' {
    BeforeEach {
        $script:order = 'Application|Package|SU Deployment Pkg|Boot Image|OS Image|Driver Package|Task Sequence'
        Mock Get-ContentApplications { [PSCustomObject]@{ ContentType = 'Application'; PackageID = 'P1' } } -ModuleName DPContentMgrCommon
        Mock Get-ContentPackages { [PSCustomObject]@{ ContentType = 'Package'; PackageID = 'P2' } } -ModuleName DPContentMgrCommon
        Mock Get-ContentSUDPs { [PSCustomObject]@{ ContentType = 'SU Deployment Pkg'; PackageID = 'P3' } } -ModuleName DPContentMgrCommon
        Mock Get-ContentBootImages { [PSCustomObject]@{ ContentType = 'Boot Image'; PackageID = 'P4' } } -ModuleName DPContentMgrCommon
        Mock Get-ContentOSImages { [PSCustomObject]@{ ContentType = 'OS Image'; PackageID = 'P5' } } -ModuleName DPContentMgrCommon
        Mock Get-ContentDriverPackages { [PSCustomObject]@{ ContentType = 'Driver Package'; PackageID = 'P6' } } -ModuleName DPContentMgrCommon
        Mock Get-ContentTaskSequenceRefs { [PSCustomObject]@{ ContentType = 'Task Sequence'; PackageID = 'P7' } } -ModuleName DPContentMgrCommon
    }

    It 'merges all seven types when no filter is given' {
        $r = @(Get-AllContentObjects)
        $r.Count | Should -Be 7
        ($r.ContentType -join '|') | Should -Be $script:order
    }

    It 'runs only the getter of the requested types <Types>' -ForEach @(
        @{ Types = @('Package');                      Expected = 'Package' }
        @{ Types = @('package');                      Expected = 'Package' }
        @{ Types = @('Boot Image', 'OS Image');       Expected = 'Boot Image|OS Image' }
        @{ Types = @('Task Sequence', 'Application'); Expected = 'Application|Task Sequence' }
        @{ Types = @('Nope');                         Expected = '' }
    ) {
        $r = @(Get-AllContentObjects -Types $Types)
        ($r.ContentType -join '|') | Should -Be $Expected
    }

    It 'skips a type whose getter returns nothing' {
        Mock Get-ContentPackages { } -ModuleName DPContentMgrCommon
        $r = @(Get-AllContentObjects)
        $r.Count | Should -Be 6
        $r.ContentType | Should -Not -Contain 'Package'
    }

    It 'gives a note for a type that fails for good and still loads the other six' {
        Mock Get-ContentBootImages { throw 'boot image query failed' } -ModuleName DPContentMgrCommon
        $state = [hashtable]::Synchronized(@{ Notes = (New-Object System.Collections.ArrayList); Step = ''; Cancel = $false })
        $r = @(Get-AllContentObjects -State $state)
        $r.Count | Should -Be 6
        $r.ContentType | Should -Not -Contain 'Boot Image'
        $state['Notes'].Count | Should -Be 1
        $state['Notes'][0] | Should -Match "Content of type 'Boot Image' failed after 1 attempt.*boot image query failed"
    }

    It 'retries a transient fault of one type and gives no note' {
        $global:Fx = @{ Hit = 0 }
        Mock Get-ContentOSImages { $global:Fx.Hit++; if ($global:Fx.Hit -eq 1) { throw 'The operation has timed out.' }; [PSCustomObject]@{ ContentType = 'OS Image'; PackageID = 'P5' } } -ModuleName DPContentMgrCommon
        $state = [hashtable]::Synchronized(@{ Notes = (New-Object System.Collections.ArrayList); Step = ''; Cancel = $false })
        @(Get-AllContentObjects -State $state).Count | Should -Be 7
        $state['Notes'].Count | Should -Be 0
        $global:Fx.Hit | Should -Be 2
    }

    It 'shows the type in the progress and stops when the cancel flag is set' {
        $global:Fx = @{ State = [hashtable]::Synchronized(@{ Notes = (New-Object System.Collections.ArrayList); Step = ''; Cancel = $false }) }
        Mock Get-ContentPackages { $global:Fx.State['Cancel'] = $true; [PSCustomObject]@{ ContentType = 'Package'; PackageID = 'P2' } } -ModuleName DPContentMgrCommon
        { Get-AllContentObjects -State $global:Fx.State } | Should -Throw '*cancelled*'
        $global:Fx.State['Step'] | Should -Be 'Loading content objects: SU Deployment Pkg...'
    }

    It 'flattens a getter that returns several objects' {
        Mock Get-ContentPackages { 1..3 | ForEach-Object { [PSCustomObject]@{ ContentType = 'Package'; PackageID = "X$_" } } } -ModuleName DPContentMgrCommon
        @(Get-AllContentObjects -Types 'Package').Count | Should -Be 3
    }
}

Describe 'Get-ContentDistributionStatus' {
    It 'pipes the content object to Get-CMDistributionStatus' {
        Mock Get-CMDistributionStatus { 'status-for-' + $InputObject.PackageID } -ModuleName DPContentMgrCommon -RemoveParameterType InputObject
        Get-ContentDistributionStatus -ContentObject ([PSCustomObject]@{ PackageID = 'MCM00001' }) | Should -Be 'status-for-MCM00001'
    }

    It 'requires a content object' {
        { Get-ContentDistributionStatus -ContentObject $null } | Should -Throw
    }
}

Describe 'Content getters report errors' {
    It 'has no SilentlyContinue in the content getters' {
        $text = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'Module\DPContentMgrCommon.psm1') -Raw
        $m = [regex]::Match($text, '(?s)function Get-ContentApplications.*?function Get-AllContentObjects')
        $m.Success | Should -BeTrue
        $m.Value | Should -Not -Match 'SilentlyContinue'
        $m.Value | Should -Not -Match '\$all \+='
    }

    It 'asks the cmdlets for the fast form where it exists' {
        Mock Get-CMApplication { } -ModuleName DPContentMgrCommon
        $null = Get-ContentApplications
        Should -Invoke Get-CMApplication -ModuleName DPContentMgrCommon -ParameterFilter { $Fast -and $ErrorAction -eq 'Stop' }
    }
}