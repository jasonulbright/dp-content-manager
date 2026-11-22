# Realistic rows for the grid data tests. The rows come out of the real module functions, which read ConfigMgr through
# stubs of the ConfigMgr commands that return data in the shape of the real commands. No row is written by hand, so
# a property name that the module stops producing shows in the grid tests.
# Call it where Write-Log of the module is mocked and the module is imported.

function New-GridFixture {
    $stubs = @('Get-CMDistributionPoint', 'Get-CMApplication', 'Get-CMPackage', 'Get-CimInstance')
    try {
        Set-Item -Path 'function:global:Get-CMDistributionPoint' -Value {
            [CmdletBinding()] param($Fast)
            [PSCustomObject]@{ NetworkOSPath = '\\dp1.contoso.com'; SiteCode = 'MCM'; IsPullDP = $false; NALPath = '["Display=\\dp1.contoso.com\"]MSWNET:["SMS_SITE=MCM"]\\dp1.contoso.com\' }
            [PSCustomObject]@{ NetworkOSPath = '\\dp2.contoso.com'; SiteCode = 'MCM'; IsPullDP = $true; NALPath = '["Display=\\dp2.contoso.com\"]MSWNET:["SMS_SITE=MCM"]\\dp2.contoso.com\' }
        }
        Set-Item -Path 'function:global:Get-CMApplication' -Value {
            [CmdletBinding()] param([switch]$Fast)
            [PSCustomObject]@{ LocalizedDisplayName = 'Contoso Viewer'; PackageID = 'MCM00010'; PackageSize = 51200; SourceDate = [datetime]'2026-11-01'; CI_ID = 16777300 }
        }
        Set-Item -Path 'function:global:Get-CMPackage' -Value {
            [CmdletBinding()] param([switch]$Fast)
            [PSCustomObject]@{ Name = 'Contoso Agent Package'; PackageID = 'MCM00020'; PackageSize = 2048; SourceDate = [datetime]'2026-11-02' }
        }
        Set-Item -Path 'function:global:Get-CimInstance' -Value {
            [CmdletBinding()] param($ComputerName, $Namespace, $Query)
            [PSCustomObject]@{ PackageID = 'MCM00010'; ServerNALPath = '["Display=\\dp1.contoso.com\"]MSWNET:["SMS_SITE=MCM"]\\dp1.contoso.com\'; State = 0; SourceVersion = 3; SummaryDate = [datetime]'2026-11-20 08:15:00' }
            [PSCustomObject]@{ PackageID = 'MCM00010'; ServerNALPath = '["Display=\\dp2.contoso.com\"]MSWNET:["SMS_SITE=MCM"]\\dp2.contoso.com\'; State = 3; SourceVersion = 3; SummaryDate = [datetime]'2026-11-20 09:30:00' }
            [PSCustomObject]@{ PackageID = 'MCM00020'; ServerNALPath = '["Display=\\dp1.contoso.com\"]MSWNET:["SMS_SITE=MCM"]\\dp1.contoso.com\'; State = 1; SourceVersion = 1; SummaryDate = [datetime]'2026-11-21 10:45:00' }
        }
        $dps = @(Get-AllDistributionPoints)
        $content = @(Get-AllContentObjects -Types 'Application', 'Package')
        $status = @(Get-BulkDistributionStatus -SMSProvider 'cm01.contoso.com' -SiteCode 'MCM')
        $byDP = @(ConvertTo-DPStatusSummary -StatusRows $status)
        $byContent = @(ConvertTo-ContentStatusSummary -StatusRows $status)
        # The same calls as the refresh handler and the selection handler of start-dpcontentmgr.ps1.
        @{
            gridDPs          = @(ConvertTo-DPRows -DPs $dps -ByDP $byDP -StorageByDP @())
            gridContent      = @(ConvertTo-ContentRows -Content $content -ByContent $byContent)
            gridContentPerDP = @(ConvertTo-ContentPerDPRows -Status $status -PackageID 'MCM00010')
            gridStatusIssues = @(ConvertTo-StatusIssueRows -Status $status -Content $content)
        }
    }
    finally { foreach ($s in $stubs) { Remove-Item -Path "function:global:$s" -ErrorAction SilentlyContinue } }
}

# Columns that a fixture cannot fill on any row, each with the reason. The grid test accepts only these.
function Get-GridEmptyAllowed {
    @{}
}
