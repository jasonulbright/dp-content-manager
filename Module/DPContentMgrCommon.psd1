@{
    RootModule        = 'DPContentMgrCommon.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'b2c3d4e5-f6a7-8901-bcde-f23456789012'
    Author            = 'Jason Ulbright'
    Description       = 'MECM distribution point content status, redistribution, validation, and orphan detection.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging
        'Initialize-Logging'
        'Write-Log'

        # CM Connection
        'Connect-CMSite'
        'Disconnect-CMSite'
        'Test-CMConnection'

        # Content Retrieval
        'Get-AllDistributionPoints'
        'Get-AllContentObjects'
        'Get-ContentApplications'
        'Get-ContentPackages'
        'Get-ContentSUDPs'
        'Get-ContentBootImages'
        'Get-ContentOSImages'
        'Get-ContentDriverPackages'
        'Get-ContentTaskSequenceRefs'

        # Status Queries
        'Get-BulkDistributionStatus'
        'Get-ContentDistributionStatus'
        'ConvertTo-ContentStatusSummary'
        'ConvertTo-DPStatusSummary'

        # Actions
        'Invoke-RedistributeContent'
        'Remove-ContentFromDP'
        'Invoke-ContentValidation'
        'Find-OrphanedContent'
        'Get-DPStorageAnalysis'

        # Export
        'Export-ContentStatusCsv'
        'Export-ContentStatusHtml'
        'New-ContentStatusSummary'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
