$win = Start-App 'start-dpcontentmgr.ps1' 'DP Content Manager'
Size-Win $win 1440 900
$win = Find-Top 'DP Content Manager' 10
function Idle([string]$id = 'btnRefresh', [int]$sec = 300) { Start-Sleep -Seconds 2; $null = Wait-Until { $b = Get-El $win $id -Quiet; $b -and $b.Current.IsEnabled } $sec; Start-Sleep -Seconds 2 }
function Rows($grid) { @((Get-El $win $grid).FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataItem)))) }
function Pick($combo, [string]$item) { Expand-El (Get-El $win $combo); Start-Sleep -Milliseconds 500; Click-El (Find-Role $win 'ListItem' $item); Start-Sleep -Milliseconds 700 }
function WinNames { $c = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window); (@($root.FindAll('Descendants', $c)) | ForEach-Object { $_.Current.Name }) -join ' ; ' }
# Open the preview dialog of a bulk action, capture it, dump its tree and close it with a button that does not run anything.
# The step never selects the button that runs the operations. It stops when no safe button exists.
function Preview-Dlg([string]$btn, [string]$shot) {
  Invoke-El (Get-El $win $btn); Start-Sleep -Seconds 4
  Log ('windows: ' + (WinNames))
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
  $d = @($root.FindAll('Descendants', $c)) | Where-Object { $_.Current.Name -notin 'DP Content Manager', 'Program Manager' -and $_.Current.Name -ne '' -and $_.Current.Name -notlike 'Administrator*' -and $_.Current.Name -notlike 'GlowWindow*' } | Select-Object -First 1
  if (-not $d) { Log ('No dialog after ' + $btn + ': the selection has no matching operations.'); return $false }
  Shot $d $shot; Dump-Tree $d $shot
  $safe = @('btnCancel', 'btnClose') | ForEach-Object { Get-El $d $_ -Quiet } | Where-Object { $_ -and $_.Current.IsEnabled } | Select-Object -First 1
  if (-not $safe) { throw ('No Cancel or Close button in the dialog of ' + $btn + '. Nothing was answered.') }
  Invoke-El $safe; Start-Sleep -Seconds 1
  return $true
}
Section 'first run' {
  Shot $win 'start-first-run'
  Invoke-El (Get-El $win 'btnRefresh'); Start-Sleep -Seconds 2; Shot $win 'refresh-without-connection'
  Invoke-El (Get-El $win 'btnOptions'); $d = Find-Last 'Options' 10; Shot $d 'options-empty'; Dump-Tree $d 'options'
  Set-Text (Get-El $d 'txtSiteCode') $Capture.SiteCode; Set-Text (Get-El $d 'txtSmsProvider') $Capture.SMSProvider; Shot $d 'options-connection'
  Invoke-El (Get-El $d 'btnCatAbout'); Start-Sleep -Milliseconds 600; Shot $d 'options-about'
  Invoke-El (Get-El $d 'btnCatConnection'); Invoke-El (Get-El $d 'btnOk'); Start-Sleep -Seconds 1
}
Section 'refresh' {
  Invoke-El (Get-El $win 'btnRefresh'); Start-Sleep -Seconds 3; Shot $win 'refresh-running'; Idle; Shot $win 'dps'
}
Section 'dps detail' {
  $r = Rows 'gridDPs'; Log ('dp rows: ' + $r.Count); if ($r.Count) { Select-El $r[0]; Start-Sleep -Seconds 2; Shot $win 'dp-detail' }
  Set-Text (Get-El $win 'txtFilter') 'CM01'; Start-Sleep -Seconds 1; Shot $win 'dps-filter-text'; Set-Text (Get-El $win 'txtFilter') ''
  Expand-El (Get-El $win 'cboStatusFilter'); Start-Sleep -Milliseconds 500; Shot $win 'dps-filter-list' -Hover; Collapse-El (Get-El $win 'cboStatusFilter')
}
Section 'exports' {
  Invoke-El (Get-El $win 'btnExportCsv'); File-Dialog $null $null; Shot $win 'dps-exported'
  Invoke-El (Get-El $win 'btnExportHtml'); File-Dialog $null $null; Start-Sleep -Seconds 1
}
Section 'content' {
  Invoke-El (Get-El $win 'btnViewContent'); Start-Sleep -Seconds 3; Shot $win 'content-before'
  Invoke-El (Get-El $win 'btnRefresh'); Idle; Shot $win 'content'
  $r = Rows 'gridContent'; Log ('content rows: ' + $r.Count);
  Set-Text (Get-El $win 'txtFilter') 'Doc'; Start-Sleep -Seconds 1; Shot $win 'content-filter-text'; Set-Text (Get-El $win 'txtFilter') ''
  Expand-El (Get-El $win 'cboStatusFilter'); Start-Sleep -Milliseconds 500; Shot $win 'content-filter-list' -Hover; Collapse-El (Get-El $win 'cboStatusFilter')
}
Section 'content actions' {
  # Find a content row that has DP status rows. Each action only opens a preview dialog, which is closed without a run.
  $picked = $null
  foreach ($x in (Rows 'gridContent')) {
    Select-El $x; Start-Sleep -Milliseconds 700
    if (Preview-Dlg 'btnValidate' 'dialog-validate-preview') { $picked = $x; break }
  }
  if (-not $picked) { Log 'ERROR: no content row gave a Validate preview.'; return }
  Select-El $picked; Start-Sleep -Seconds 2; Shot $win 'content-detail'
  [void](Preview-Dlg 'btnRedistribute' 'dialog-redistribute-preview')
  [void](Preview-Dlg 'btnRemoveFromDP' 'dialog-remove-preview')
  Shot $win 'content-actions-log'
}
Section 'status issues' {
  Invoke-El (Get-El $win 'btnViewStatusIssues'); Start-Sleep -Seconds 3; Shot $win 'issues-before'
  Invoke-El (Get-El $win 'btnRefresh'); Idle; Shot $win 'issues'
  $r = Rows 'gridStatusIssues'; Log ('issue rows: ' + $r.Count)
}
Section 'tab order' { Invoke-El (Get-El $win 'btnViewDPs'); Focus-Win $win; Click-El (Get-El $win 'btnViewDPs'); 1..14 | ForEach-Object { Send-Keys '{TAB}'; $f = [System.Windows.Automation.AutomationElement]::FocusedElement; Log ('tab -> ' + $f.Current.AutomationId + ' / ' + $f.Current.Name) } }
Section 'light' {
  Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 1
  foreach ($v in @(@('DPs', 'light-dps'), @('Content', 'light-content'), @('StatusIssues', 'light-issues'))) { Invoke-El (Get-El $win ('btnView' + $v[0])); Start-Sleep -Seconds 2; Shot $win $v[1] }
  Invoke-El (Get-El $win 'btnOptions'); $d = Find-Last 'Options' 10; Shot $d 'light-options'; Focus-Win $d; Send-Keys '{ESC}'; Start-Sleep -Seconds 1
  Toggle-El (Get-El $win 'toggleTheme')
}
Section 'persist' { Invoke-El (Get-El $win 'btnViewContent'); Start-Sleep -Milliseconds 600; Invoke-El (Get-El $win 'Close'); Start-Sleep -Seconds 8; $win = Start-App 'start-dpcontentmgr.ps1' 'DP Content Manager'; Shot $win 'start-reopen' }
Section 'files' {
  New-Item -ItemType Directory "$OutDir\files" -Force | Out-Null
  Get-ChildItem $appDir -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notlike '*\Lib\*' -and $_.FullName -notlike '*\Module\*' -and $_.Extension -in '.json', '.jsonl', '.log', '.csv', '.html', '.txt' } | ForEach-Object { Log ('file: ' + $_.FullName.Substring($appDir.Length) + ' ' + $_.Length); Copy-Item $_.FullName ("$OutDir\files\" + $_.Name) -Force }
}
Stop-App 'start-dpcontentmgr.ps1'
