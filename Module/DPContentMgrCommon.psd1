@{
    RootModule        = 'DPContentMgrCommon.psm1'
    ModuleVersion     = '1.2.1'
    GUID              = 'b2c3d4e5-f6a7-8901-bcde-f23456789012'
    Author            = 'Jason Ulbright'
    Description       = 'MECM distribution point content status, redistribution, validation, and orphan detection.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging and CM connection come from the vendored SuiteCommon
        # module (Lib\SuiteCommon), imported globally by the root module.

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
