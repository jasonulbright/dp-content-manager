#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x tests for the main window text (XAML and script only, no window is opened).
#>

BeforeAll {
    $script:strRoot = Split-Path -Path $PSScriptRoot -Parent
    $script:xmlWindow = [xml](Get-Content -LiteralPath (Join-Path -Path $script:strRoot -ChildPath 'MainWindow.xaml') -Raw)
    $script:nsWpf = New-Object System.Xml.XmlNamespaceManager($script:xmlWindow.NameTable)
    $script:nsWpf.AddNamespace('w', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
    $script:nsWpf.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
}

Describe 'Page subtitle' {
    It 'starts with the subtitle of the DPs view, so that no hint about Refresh stays after the connection' {
        $strScript = Get-Content -LiteralPath (Join-Path -Path $script:strRoot -ChildPath 'start-dpcontentmgr.ps1') -Raw
        $strDpsSubtitle = [regex]::Match($strScript, "'DPs'\s*=\s*@\{ Title = 'DPs';\s*Subtitle = '([^']+)'").Groups[1].Value
        $strDpsSubtitle | Should -Not -BeNullOrEmpty
        $nodeSubtitle = $script:xmlWindow.SelectSingleNode("//w:TextBlock[@x:Name='txtModuleSubtitle']", $script:nsWpf)
        $nodeSubtitle.GetAttribute('Text') | Should -BeExactly $strDpsSubtitle
    }
}

Describe 'Content type filter' {
    It 'offers only content types that the module gives to the rows, so that every choice can match' {
        $strScript = Get-Content -LiteralPath (Join-Path -Path $script:strRoot -ChildPath 'start-dpcontentmgr.ps1') -Raw
        $strModule = Get-Content -LiteralPath (Join-Path -Path $script:strRoot -ChildPath 'Module\DPContentMgrCommon.psm1') -Raw
        $strList = [regex]::Match($strScript, "(?s)'Content' \{.*?foreach \(\`$it in @\(([^)]+)\)\)").Groups[1].Value
        $strList | Should -Not -BeNullOrEmpty
        $arrChoice = @([regex]::Matches($strList, "'([^']+)'") | ForEach-Object { $_.Groups[1].Value } | Where-Object { $_ -ne 'All' })
        $arrChoice.Count | Should -BeGreaterThan 0
        $arrType = @([regex]::Matches($strModule, "ContentType\s*=\s*'([^']+)'") | ForEach-Object { $_.Groups[1].Value })
        foreach ($strChoice in $arrChoice) { $arrType | Should -Contain $strChoice -Because "the filter choice $strChoice" }
    }
}
