# Changelog

All notable changes to DP Content Manager are documented in this file.

## [1.1.0] - 2026-08-14

### Changed

- **Shared plumbing moved to the vendored `SuiteCommon` module.** Logging
  (`Initialize-Logging`, `Write-Log`) and CM site connection
  (`Connect-CMSite`, `Disconnect-CMSite`, `Test-CMConnection`) now load
  from `Lib\SuiteCommon\`, shared across the tool suite and synced from
  the suite-core repository instead of hand-edited per repo. Behavior is
  unchanged — same log format, same connection flow (globally scoped
  CMSite PSDrive, `Get-CMSite` verification) — and the connection now
  survives a stale drive whose provider connection died: `Connect-CMSite`
  tears it down and rebuilds it instead of failing.
- **Redistribution reads the connection via `Get-CMConnectionInfo`**
  instead of module-private state, keeping the bulk WMI redistribute path
  working across the module split.

## [1.0.0] - 2026-05-02

DP Content Manager is a MahApps.Metro WPF GUI for auditing, redistributing,
validating, and removing MECM distribution point content at scale. Extract
the zip and run `start-dpcontentmgr.ps1`.

### Features

- **DPs view** -- per-DP rollup grid: status glyph, DP name, site, total
  content, installed, in progress, failed, % OK, Pull DP flag. Detail
  panel shows storage analysis (TotalSizeGB) for the selected DP.
- **Content view** -- per-content-object rollup: glyph, name, type
  (Application / Package / SUDP / BootImage / OSImage / DriverPackage /
  TaskSequence), PackageID, size (auto-scaled KB / MB / GB), DPs targeted,
  installed, failed, % OK. Selection drives a per-DP status grid below
  showing exactly where the content is healthy / failed / in-progress.
  Action bar: Redistribute Failed, Remove from DP, Validate.
- **Status Issues view** -- the triage list: every (DP, Content) row
  currently in a non-OK state, with status filters (Failed / In progress /
  All). Multi-select rows + Redistribute Selected or Remove Failures for
  bulk remediation.
- **Bulk operation safety** -- every destructive bulk action (Redistribute,
  Remove from DP, Validate) opens a preview dialog showing the full
  (DP, Content) impact list BEFORE running. The "Run N operations" button
  executes per-row with live OK / FAIL status and per-row error detail.
  No accidental "remove all failed across 200 DPs" without seeing the
  list first.
- **Modal dialogs** -- Bulk Op Confirm (preview-before-commit),
  Options (Connection, About). All MetroWindow inline-XAML, theme-
  honoring, drag-fallback installed.
- **Core module** `DPContentMgrCommon.psm1` with 28 exported functions
  covering logging, CM connection, 7-content-type retrieval, WMI bulk
  status aggregation (`Get-BulkDistributionStatus`), per-content and
  per-DP summary aggregation, redistribute / validate / remove / orphan
  detection / storage analysis, and CSV / HTML export.
- **WPF brand alignment** -- MahApps.Metro shell with sidebar navigation,
  glyph status (no red / green for state), animated ProgressRing during
  refresh, log drawer, status bar, dark and light themes with runtime
  toggle, window-state persistence with legacy WinForms schema bridge.

### Stack

- PowerShell 5.1 + .NET Framework 4.7.2+
- WPF + MahApps.Metro 2.4.10 (vendored under `Lib/`)
- ConfigurationManager PowerShell module + raw CIM for the bulk status
  summarizer (CM console required on the host machine)
