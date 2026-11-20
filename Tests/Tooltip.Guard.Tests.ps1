#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Guard tests for tooltips. They parse the XAML and the XAML here-strings of the script. They create
# no WPF control and open no window, so they run on Windows PowerShell 5.1 and on PowerShell 7.
#
# Rules:
#   1. Every control in the audited list has a ToolTip.
#   2. A ToolTip is plain ASCII, is not empty, ends in a period and is at most 159 characters long.


BeforeAll {
    $script:Root = Split-Path -Path $PSScriptRoot -Parent
    $script:XamlNs = 'http://schemas.microsoft.com/winfx/2006/xaml'

    # ---- Per-app configuration ------------------------------------------------
    # The x:Name of every control that must explain itself. A dialog button whose label says it all
    # (OK, Cancel, Close), a read-only output box and a grid column are not in the list.
    $script:AuditedName = @(
        'toggleTheme', 'btnViewDPs', 'btnViewContent', 'btnViewStatusIssues', 'btnOptions',
        'btnRefresh', 'txtFilter', 'cboStatusFilter', 'btnExportCsv', 'btnExportHtml',
        'btnRedistribute', 'btnRemoveFromDP', 'btnValidate', 'btnRedistributeIssues', 'btnRemoveIssues',
        'btnCancelWork',
        'btnRun', 'btnCatConnection', 'btnCatAbout', 'txtSiteCode', 'txtSmsProvider'
    )
    $script:MaxLength = 159

    # Every XAML document of the app: each .xaml file and each XAML here-string of a script or module.
    function Get-XamlDocument {
        $skip = '\\(Tests|tools|docs|Lib|\.git|\.github|\.claude|TestResults)\\'
        $files = @(Get-ChildItem -LiteralPath $script:Root -Recurse -File -Include '*.xaml', '*.ps1', '*.psm1' |
                Where-Object { $_.FullName.Substring($script:Root.Length) -notmatch $skip })
        $docs = New-Object System.Collections.Generic.List[object]
        foreach ($f in $files) {
            $rel = $f.FullName.Substring($script:Root.Length).TrimStart('\')
            $text = [IO.File]::ReadAllText($f.FullName)
            $chunks = New-Object System.Collections.Generic.List[string]
            if ($f.Extension -eq '.xaml') { $chunks.Add($text) }
            else {
                foreach ($m in [regex]::Matches($text, "(?ms)@(?:'|`")[ \t]*\r?\n(\s*<.*?)\r?\n(?:'|`")@")) { $chunks.Add($m.Groups[1].Value) }
            }
            $i = 0
            foreach ($c in $chunks) {
                $i++
                $x = New-Object System.Xml.XmlDocument
                $err = ''
                try { $x.LoadXml($c) }
                catch {
                    $wrapped = '<Wrap xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" xmlns:Controls="http://metro.mahapps.com/winfx/xaml/controls">' + ($c -replace '^\s*<\?xml[^>]*\?>', '') + '</Wrap>'
                    try { $x.LoadXml($wrapped) } catch { $err = $_.Exception.Message }
                }
                $label = if ($f.Extension -eq '.xaml') { $rel } else { '{0} xaml {1}' -f $rel, $i }
                $docs.Add([pscustomobject]@{ Label = $label; Xml = $x; ParseError = $err })
            }
        }
        , $docs.ToArray()
    }

    # The elements that carry an x:Name, with the name and the ToolTip text (null when there is none).
    function Get-NamedElement {
        param($Document)
        foreach ($el in @($Document.Xml.SelectNodes('//*[@*[local-name()="Name"]]'))) {
            $name = $el.GetAttribute('Name', $script:XamlNs)
            if (-not $name) { continue }
            $tip = $null
            if ($el.HasAttribute('ToolTip')) { $tip = $el.GetAttribute('ToolTip') }
            [pscustomobject]@{ Label = $Document.Label; Name = $name; Element = $el.LocalName; ToolTip = $tip }
        }
    }

    # The problems of one ToolTip text: empty, not ASCII, no final period, too long.
    function Get-ToolTipProblem {
        param([AllowNull()][string]$Text, [int]$Max = 159)
        if ([string]::IsNullOrWhiteSpace($Text)) { return 'is empty' }
        if ($Text -cmatch '[^\x20-\x7E]') { 'has a character that is not ASCII' }
        if ($Text.Trim() -notmatch '\.$') { 'does not end in a period' }
        if ($Text.Length -gt $Max) { 'is longer than {0} characters' -f $Max }
    }

    # The audited names that the documents do not tool-tip correctly. A name that no document holds is a problem too.
    function Get-AuditProblem {
        param($Document, [string[]]$Name)
        $all = @(foreach ($d in $Document) { Get-NamedElement -Document $d })
        foreach ($n in $Name) {
            $hits = @($all | Where-Object { $_.Name -eq $n })
            if ($hits.Count -eq 0) { '{0}: no element has this name' -f $n; continue }
            foreach ($h in $hits) {
                foreach ($p in @(Get-ToolTipProblem -Text $h.ToolTip -Max $script:MaxLength)) { '{0} ({1}): the ToolTip {2}' -f $h.Name, $h.Label, $p }
            }
        }
    }

    function New-XmlSample {
        param([string]$Body)
        $x = New-Object System.Xml.XmlDocument
        $x.LoadXml('<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">' + $Body + '</Window>')
        [pscustomobject]@{ Label = 'sample'; Xml = $x; ParseError = '' }
    }

    $script:Documents = Get-XamlDocument
}

Describe 'The tooltips of the window' {
    It 'finds the XAML of the app (the scan is not empty)' {
        @($script:Documents).Count | Should -BeGreaterThan 1
        @($script:Documents | Where-Object { $_.Label -eq 'MainWindow.xaml' }).Count | Should -Be 1
        @($script:AuditedName).Count | Should -BeGreaterThan 10
    }

    It 'parses every XAML document' {
        @($script:Documents | Where-Object { $_.ParseError } | ForEach-Object { '{0}: {1}' -f $_.Label, $_.ParseError }) | Should -BeNullOrEmpty
    }

    It 'writes every audited ToolTip as ASCII, ends it in a period and keeps it short' {
        @(Get-AuditProblem -Document $script:Documents -Name $script:AuditedName) | Should -BeNullOrEmpty
    }

    It 'writes every other ToolTip in XAML as ASCII with a final period' {
        $bad = foreach ($d in $script:Documents) {
            foreach ($e in @(Get-NamedElement -Document $d | Where-Object { $null -ne $_.ToolTip })) {
                foreach ($p in @(Get-ToolTipProblem -Text $e.ToolTip -Max $script:MaxLength)) { '{0} ({1}): the ToolTip {2}' -f $e.Name, $e.Label, $p }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }
}

Describe 'The tooltip check itself' {
    It 'fails a control without a ToolTip (bad sample)' {
        $d = New-XmlSample '<Button x:Name="btnGo" Content="Go"/>'
        @(Get-AuditProblem -Document @($d) -Name @('btnGo')).Count | Should -Be 1
    }

    It 'fails an empty ToolTip, a ToolTip without a period and a long ToolTip (bad sample)' {
        $long = ('A' * 160) + '.'
        $d = New-XmlSample ('<StackPanel><Button x:Name="btnA" ToolTip=""/><Button x:Name="btnB" ToolTip="Run it"/><Button x:Name="btnC" ToolTip="' + $long + '"/></StackPanel>')
        @(Get-AuditProblem -Document @($d) -Name @('btnA', 'btnB', 'btnC')).Count | Should -Be 3
    }

    It 'fails a ToolTip with a character that is not ASCII (bad sample)' {
        $d = New-XmlSample ('<Button x:Name="btnA" ToolTip="Run it' + [char]0x2013 + 'now."/>')
        @(Get-AuditProblem -Document @($d) -Name @('btnA')).Count | Should -Be 1
    }

    It 'fails a name that no element holds (bad sample)' {
        $d = New-XmlSample '<Button x:Name="btnA" ToolTip="Run it."/>'
        @(Get-AuditProblem -Document @($d) -Name @('btnGone')).Count | Should -Be 1
    }

    It 'passes a control with a good ToolTip' {
        $d = New-XmlSample '<Button x:Name="btnA" ToolTip="Run the operation."/>'
        @(Get-AuditProblem -Document @($d) -Name @('btnA')).Count | Should -Be 0
    }
}
