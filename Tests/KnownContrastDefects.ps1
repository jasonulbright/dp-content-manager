# Contrast violations that exist today and are not fixed yet. Each entry is
# "<control>:<metric>" or "text:<visible text>". The window test fails when a
# violation is not listed (a new defect) and when a listed one no longer happens
# (remove it here). Run Show-ContrastViolations.ps1 to print the current list.
#
# Text must reach 4.5:1 and a control border 3:1, enabled and disabled, in both
# themes. The causes are shared by every suite tool: the SuiteCommon palette
# (sidebar button border #555555 on #1E1E1E is 2.06:1), the MahApps disabled
# style, the light text box border and the unselected tab header. The fix
# belongs in app-packager-suite, then a sync of the vendored copy, not here.

function Get-KnownContrastDefect {
    param([Parameter(Mandatory)][ValidateSet('dark', 'light')][string]$Theme)
    switch ($Theme) {
        'dark' {
            @(
                'btnExportCsv:DisabledText', 'btnExportHtml:DisabledText', 'btnOptions:DisabledBorder',
                'btnOptions:EnabledBorder', 'btnRedistribute:DisabledText', 'btnRedistributeIssues:DisabledText',
                'btnRefresh:DisabledText', 'btnRemoveFromDP:DisabledText', 'btnRemoveIssues:DisabledText',
                'btnValidate:DisabledText', 'btnViewContent:DisabledBorder', 'btnViewContent:DisabledText',
                'btnViewContent:EnabledBorder', 'btnViewDPs:DisabledBorder', 'btnViewDPs:DisabledText',
                'btnViewDPs:EnabledBorder', 'btnViewStatusIssues:DisabledBorder', 'btnViewStatusIssues:DisabledText',
                'btnViewStatusIssues:EnabledBorder', 'titlebar:inactive'
            )
        }
        'light' {
            @(
                'btnExportCsv:DisabledText', 'btnExportHtml:DisabledText', 'btnOptions:DisabledText',
                'btnRedistribute:DisabledText', 'btnRedistributeIssues:DisabledText', 'btnRefresh:DisabledText',
                'btnRemoveFromDP:DisabledText', 'btnRemoveIssues:DisabledText', 'btnValidate:DisabledText',
                'btnViewContent:DisabledText', 'btnViewDPs:DisabledText', 'btnViewStatusIssues:DisabledText',
                'cboStatusFilter:DisabledBorder', 'cboStatusFilter:DisabledText', 'cboStatusFilter:EnabledBorder',
                'cboStatusFilter:EnabledText', 'titlebar:inactive', 'txtFilter:DisabledBorder',
                'txtFilter:EnabledBorder'
            )
        }
    }
}