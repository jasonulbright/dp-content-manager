# App-specific names for Window.Tests.ps1.
$script:AppConfig = @{
    Title         = 'DP Content Manager'
    Shell         = 'start-dpcontentmgr.ps1'
    Xaml          = 'MainWindow.xaml'
    OptionsButton = 'btnOptions'
    LogLabel      = 'lblLogOutput'
    WorkflowButtons = @('btnRefresh', 'btnExportCsv', 'btnExportHtml', 'btnRedistribute', 'btnRemoveFromDP', 'btnValidate', 'btnRedistributeIssues', 'btnRemoveIssues')
    Views         = @(
        @{ Name = 'DPs';           Container = 'viewDPs';          Button = 'btnViewDPs' }
        @{ Name = 'Content';       Container = 'viewContent';      Button = 'btnViewContent' }
        @{ Name = 'Status Issues'; Container = 'viewStatusIssues'; Button = 'btnViewStatusIssues' }
    )
}
