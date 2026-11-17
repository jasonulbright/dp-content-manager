# Helpers of Theme.Guard.Tests.ps1. Dot-source from BeforeDiscovery and from BeforeAll.

$script:RepoRoot = Split-Path -Parent $PSScriptRoot
$script:ThemeProperty = @('FontSize', 'FontWeight', 'FontStyle', 'TextDecorations', 'FontFamily', 'FontStretch')
$script:BannedPattern = '(?i)\b(easily|simply|seamless(ly)?|powerful|robust|streamline[sd]?|streamlining|quickly|just|effortless(ly)?|intuitive(ly)?|handy|helps you|allows you to)\b'

# True when a brush value names an accent or a blue color.
function Test-BlueBrush {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    if ($Value -match 'MahApps\.Brushes\.(Accent\d*|AccentBase|Highlight|Button\.Flat\.Background\.Pressed)\b') { return $true }
    if ($Value -match '(?i)^\s*(blue|dodgerblue|royalblue|steelblue|navy|mediumblue|darkblue|cornflowerblue)\s*$') { return $true }
    if ($Value -match '^\s*#(?:[0-9A-Fa-f]{2})?([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})\s*$') {
        $r = [Convert]::ToInt32($Matches[1], 16) / 255.0
        $g = [Convert]::ToInt32($Matches[2], 16) / 255.0
        $b = [Convert]::ToInt32($Matches[3], 16) / 255.0
        $max = [Math]::Max($r, [Math]::Max($g, $b)); $min = [Math]::Min($r, [Math]::Min($g, $b))
        $d = $max - $min
        if ($d -lt 0.2 -or $max -eq 0) { return $false }
        if ($max -ne $b) { return $false }
        $hue = 60.0 * ((4.0 + (($r - $g) / $d)))
        return ($hue -ge 190 -and $hue -le 260)
    }
    return $false
}

function Test-WhiteBrush {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value -match '(?i)^\s*(white|#FFFFFF|#FFFFFFFF)\s*$' -or $Value -match 'MahApps\.Brushes\.(Ideal|ThemeIdeal)Foreground')
}

# The value of a Setter: the Value attribute, or the Color of a nested brush.
function Get-SetterValue {
    param($Setter)
    if ($Setter.HasAttribute('Value')) { return $Setter.GetAttribute('Value') }
    $node = $Setter.SelectSingleNode('.//*[@Color]')
    if ($node) { return $node.GetAttribute('Color') }
    return ''
}

# Every XAML document of the repository: the XAML files and the XAML here-strings of the scripts.
function Get-XamlDocument {
    $docs = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $script:RepoRoot -Filter '*.xaml' -File -Recurse | Where-Object { $_.FullName -notmatch '\\(\.claude|\.git)\\' })) {
        $docs += [PSCustomObject]@{ Source = $f.Name; Text = (Get-Content -LiteralPath $f.FullName -Raw) }
    }
    foreach ($f in @(Get-ChildItem -LiteralPath $script:RepoRoot -File -Recurse | Where-Object { $_.Extension -in '.ps1', '.psm1' -and $_.FullName -notmatch '\\(\.claude|\.git|Tests|tools)\\' })) {
        $raw = Get-Content -LiteralPath $f.FullName -Raw
        $i = 0
        foreach ($m in [regex]::Matches($raw, "(?s)@'\r?\n(\s*<[^\r\n]*Window.*?)\r?\n'@")) {
            $i++
            $docs += [PSCustomObject]@{ Source = ('{0} here-string {1}' -f $f.Name, $i); Text = $m.Groups[1].Value }
        }
    }
    return $docs
}

# Returns one text per blue-background rule that has no white foreground.
function Get-BlueWithoutWhite {
    param([string]$Text, [string]$Source)
    [xml]$x = $Text
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($el in @($x.SelectNodes('//*'))) {
        if ($el.LocalName -eq 'Setter' -and $el.GetAttribute('Property') -eq 'Background') {
            if (Test-BlueBrush (Get-SetterValue $el)) {
                $fg = $el.ParentNode.ChildNodes | Where-Object { $_.LocalName -eq 'Setter' -and $_.GetAttribute('Property') -eq 'Foreground' } | Select-Object -First 1
                if (-not $fg -or -not (Test-WhiteBrush (Get-SetterValue $fg))) {
                    [void]$found.Add(('{0}: Setter Background {1} in <{2}> has no white Foreground' -f $Source, (Get-SetterValue $el), $el.ParentNode.LocalName))
                }
            }
        }
        elseif ($el.HasAttribute('Background') -and $el.LocalName -ne 'Setter') {
            if (Test-BlueBrush $el.GetAttribute('Background')) {
                if (-not (Test-WhiteBrush $el.GetAttribute('Foreground'))) {
                    [void]$found.Add(('{0}: <{1}> Background {2} has no white Foreground' -f $Source, $el.LocalName, $el.GetAttribute('Background')))
                }
            }
        }
    }
    return $found.ToArray()
}

# Every string of a script: the constant strings and the text parts of expandable strings.
function Get-ScriptString {
    param([string]$Path)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    $nodes = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst] }, $true)
    foreach ($n in $nodes) { [string]$n.Value }
}

# The text a user can read in a XAML document: tooltips and text attributes.
function Get-XamlUiText {
    param([string]$Text)
    [xml]$x = $Text
    foreach ($el in @($x.SelectNodes('//*'))) {
        foreach ($a in @($el.Attributes)) {
            if ($a.LocalName -match '^(ToolTip|Text|Content|Header|Title|Watermark|OnContent|OffContent|ToolTipService.*)$' -or $a.Name -match 'Watermark$') {
                if ($a.Value -notmatch '^\{') { $a.Value }
            }
        }
    }
}

# One text per catch block that sends raw exception text to a UI sink.
# UI sinks: Add-LogLine, Set-StatusText, the themed dialogs, and a UI property such as Text or Detail.
# The log file (Write-Log) and a state field that the UI does not show are allowed.
function Get-RawExceptionToUi {
    param([string]$Path, [string]$ScriptText)
    $ast = if ($ScriptText) { [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$null, [ref]$null) } else { [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null) }
    $rawPattern = '\$_(?![A-Za-z0-9_])|\$PSItem|\$Error\b|\.Exception\b|\$ex\b|\$err\b'
    $uiProperty = @('Text', 'Content', 'Detail', 'ToolTip', 'Title', 'Header', 'ErrorMsg')
    $sinkCommand = @('Add-LogLine', 'Set-StatusText', 'Show-ThemedMessage', 'Show-ConfirmDialog', 'Show-Dialog')
    $found = @()
    foreach ($c in @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CatchClauseAst] }, $true))) {
        foreach ($cmd in @($c.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true))) {
            $name = [string]$cmd.GetCommandName()
            if ($sinkCommand -contains $name -and $cmd.Extent.Text -match $rawPattern) { $found += ('line {0}: {1}' -f $cmd.Extent.StartLineNumber, $cmd.Extent.Text) }
        }
        foreach ($as in @($c.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] }, $true))) {
            $member = [string]$as.Left.Member.Value
            if ($uiProperty -contains $member -and $as.Right.Extent.Text -match $rawPattern) { $found += ('line {0}: {1}' -f $as.Extent.StartLineNumber, $as.Extent.Text) }
        }
    }
    return $found
}

# The title of a value box: the TextBlock that comes before it in the same panel (the rule of SuiteCommon).
function Get-FieldTitleElement {
    param($Box)
    $prev = $Box.PreviousSibling
    while ($prev -and $prev.NodeType -ne 'Element') { $prev = $prev.PreviousSibling }
    if ($prev -and $prev.LocalName -in 'TextBlock', 'Label') { return $prev }
    return $null
}

# Lists every hard-coded color in the Setters and the attributes of the given nodes.
# Only a DynamicResource or a TemplateBinding is allowed.
function Get-HardColor {
    param($Nodes, [string]$Label)
    $brushProps = 'Background|Foreground|BorderBrush|Fill|Stroke|OpacityMask|CaretBrush|SelectionBrush'
    foreach ($n in $Nodes) {
        foreach ($setter in @($n.SelectNodes('descendant-or-self::*[local-name()="Setter"]'))) {
            if ($setter.GetAttribute('Property') -notmatch ('^(' + $brushProps + ')$')) { continue }
            $v = $setter.GetAttribute('Value')
            if ($v) {
                if ($v -notmatch '^\{(DynamicResource|TemplateBinding)\b') { '{0}: Setter {1}="{2}"' -f $Label, $setter.GetAttribute('Property'), $v }
            }
            else {
                foreach ($c in @($setter.SelectNodes('descendant::*[@Color]'))) {
                    if ($c.GetAttribute('Color') -notmatch '^\{DynamicResource\b') { '{0}: Setter {1} holds the color {2}' -f $Label, $setter.GetAttribute('Property'), $c.GetAttribute('Color') }
                }
            }
        }
        foreach ($el in @($n.SelectNodes('descendant-or-self::*[not(local-name()="Setter")]'))) {
            foreach ($attr in @($el.Attributes)) {
                if ($attr.LocalName -match ('^(' + $brushProps + ')$') -and $attr.Value -notmatch '^\{(DynamicResource|TemplateBinding)\b') { '{0}: {1} {2}="{3}"' -f $Label, $el.LocalName, $attr.LocalName, $attr.Value }
            }
        }
    }
}
# True when a UI title is in Title Case: each word starts with a capital letter or a digit, except the small words in the
# middle. A word of capitals (CSV) and a word with a dot (cm01.contoso.com) pass. The text may end with dots.
function Test-TitleCase {
    param([string]$Text)
    $clean = ($Text -replace '\.\.\.$', '').Trim()
    if ([string]::IsNullOrWhiteSpace($clean)) { return $true }
    $small = @('a', 'an', 'the', 'and', 'or', 'of', 'to', 'in', 'on', 'for', 'per', 'by', 'at', 'as', 'but', 'nor', 'with', 'from')
    $words = @($clean -split '\s+')
    for ($i = 0; $i -lt $words.Count; $i++) {
        $word = $words[$i].Trim(',', ':', '(', ')')
        if (-not $word) { continue }
        $edge = ($i -eq 0 -or $i -eq $words.Count - 1)
        if (-not $edge -and ($small -contains $word)) { continue }
        if ($word -cmatch '^v\d[\d.]*$') { continue }
        if ($word -cnotmatch '^[A-Z0-9#<\[\(]') { return $false }
    }
    return $true
}

# The titles of a XAML document that must be in Title Case: button content, column headers, tab and group headers, window
# titles, check box and radio button text, and the bold heading text blocks. A binding stays out.
function Get-XamlTitle {
    param([string]$Text)
    [xml]$xml = $Text
    foreach ($el in @($xml.SelectNodes('//*'))) {
        $name = $el.LocalName
        $value = ''
        if ($name -in 'Button', 'CheckBox', 'RadioButton', 'ToggleButton') { $value = $el.GetAttribute('Content') }
        elseif ($name -in 'DataGridTextColumn', 'DataGridCheckBoxColumn', 'DataGridTemplateColumn', 'GroupBox', 'TabItem', 'Expander') { $value = $el.GetAttribute('Header') }
        elseif ($name -eq 'MetroWindow' -or $name -eq 'Window') { $value = $el.GetAttribute('Title') }
        elseif ($name -eq 'TextBlock' -and $el.GetAttribute('FontWeight') -in 'Bold', 'SemiBold') { $value = $el.GetAttribute('Text') }
        if ($value -and $value -notmatch '^\{') { $value }
    }
}


# Returns one text per button problem in a vendored SuiteCommon dialog (a here-string of Lib\SuiteCommon\SuiteCommon.psm1).
# The dialogs do not use the shared styles of MainWindow.xaml: each one defines its own ThemedButtonStyle. Every Button has
# a Style whose chain of BasedOn styles ends at ThemedButtonStyle. ThemedButtonStyle is the only button style that is based
# on MahApps directly. No button, style or control template has a literal color.
function Get-SuiteDialogButtonViolation {
    param([string]$Text, [string]$Source)
    $keyNs = 'http://schemas.microsoft.com/winfx/2006/xaml'
    [xml]$xml = $Text
    $found = New-Object System.Collections.Generic.List[string]
    $leaf = { param($n) if ([string]::IsNullOrEmpty($n)) { '' } else { ($n -split '[.:]')[-1] } }
    $resKey = { param($v) if ($v -match '^\{(?:Static|Dynamic)Resource\s+([^}\s]+)\}$') { $Matches[1] } else { '' } }
    $isLiteral = { param($v) ($v -match '^\s*#[0-9A-Fa-f]{3,8}\s*$') -or ($v -match '^\s*[A-Za-z]+\s*$') }
    $styleBase = @{}
    foreach ($st in @($xml.SelectNodes("//*[local-name()='Style'][@TargetType='Button']"))) { $styleBase[$st.GetAttribute('Key', $keyNs)] = $st.GetAttribute('BasedOn') }
    $reachesThemed = {
        param($key)
        $seen = @{}
        while ($key -and -not $seen.ContainsKey($key)) {
            if ($key -eq 'ThemedButtonStyle') { return $true }
            $seen[$key] = $true
            if (-not $styleBase.ContainsKey($key)) { return $false }
            $key = & $resKey $styleBase[$key]
        }
        return $false
    }
    foreach ($key in @($styleBase.Keys)) {
        if ($key -ne 'ThemedButtonStyle' -and $styleBase[$key] -match 'MahApps\.') { [void]$found.Add(('{0}: Button style {1} is based on MahApps and not on ThemedButtonStyle' -f $Source, $key)) }
    }
    foreach ($btn in @($xml.SelectNodes("//*[local-name()='Button']"))) {
        $name = $btn.GetAttribute('Name', $keyNs)
        $key = & $resKey $btn.GetAttribute('Style')
        if (-not $key -or -not (& $reachesThemed $key)) { [void]$found.Add(('{0}: Button {1} does not take ThemedButtonStyle' -f $Source, $name)) }
        foreach ($attr in @($btn.Attributes)) {
            if ((& $leaf $attr.Name) -match '(Background|Foreground|Brush)$' -and (& $isLiteral $attr.Value)) { [void]$found.Add(('{0}: Button {1} sets {2} to the literal {3}' -f $Source, $name, $attr.Name, $attr.Value)) }
        }
    }
    $scopes = @($xml.SelectNodes("//*[local-name()='Style']")) + @($xml.SelectNodes("//*[local-name()='ControlTemplate']"))
    foreach ($scope in $scopes) {
        $key = $scope.GetAttribute('Key', $keyNs)
        foreach ($node in @($scope.SelectNodes('descendant-or-self::*'))) {
            foreach ($attr in @($node.Attributes)) {
                $l = & $leaf $attr.Name
                if ($l -eq 'Value') {
                    if ((& $leaf $node.GetAttribute('Property')) -notmatch '(Background|Foreground|Brush)$') { continue }
                }
                elseif ($l -notmatch '^(Color|Fill|Stroke)$' -and $l -notmatch '(Background|Foreground|Brush)$') { continue }
                if (& $isLiteral $attr.Value) { [void]$found.Add(('{0}: {1} {2} uses the literal color {3}' -f $Source, $scope.LocalName, $key, $attr.Value)) }
            }
        }
    }
    return $found.ToArray()
}

# The XAML here-strings of the vendored SuiteCommon: its dialogs and its control templates.
function Get-SuiteCommonXamlDocument {
    $raw = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1') -Raw
    $docs = @()
    $i = 0
    foreach ($m in [regex]::Matches($raw, "(?s)@'\r?\n(\s*<.*?)\r?\n'@")) {
        $text = $m.Groups[1].Value
        if ($text -notmatch 'xmlns="http://schemas\.microsoft\.com/winfx/2006/xaml/presentation"') { continue }
        $i++
        $docs += [PSCustomObject]@{ Source = ('SuiteCommon.psm1 here-string {0}' -f $i); Text = $text }
    }
    return $docs
}
