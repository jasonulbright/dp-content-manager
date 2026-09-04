# Changelog

All notable changes to DP Content Manager are documented in this file.

## [1.2.3] - 2026-09-04

### Fixed

- **Version labels read the script header.** The sidebar version and the
  About panel carried literal version strings that no release updated;
  both now read the entry script's `Version` header at startup, so the
  window always names the version that is actually installed.

## [1.2.2] - 2026-09-04

### Changed

- **Vendored `SuiteCommon` 0.4.3.** The module repairs the process
  PSModulePath at import: a Windows PowerShell process launched from
  PowerShell 7 inherits the 7.x module directories, and the background
  runspace opened later autoloaded a Microsoft.PowerShell.Utility without
  Get-FileHash or ConvertFrom-Json, so background operations failed with
  an unrelated "term not recognized". A background runspace whose module
  import fails is disposed and the original error thrown instead of being
  returned as an opened but unusable worker.

## [1.2.1] - 2026-08-16

### Changed

- **Vendored `SuiteCommon` 0.3.2.** Window restore applies the saved
  geometry before maximizing, so un-maximizing returns to the saved size
  instead of the XAML defaults; background runspace bootstrap failures
  are named in the log instead of surfacing later as an unrelated
  "term not recognized".

## [1.2.0] - 2026-08-16

### Changed

- **Window chrome, theming, dialogs, and background-work now come from
  the vendored `SuiteCommon` module** (0.3.0): the title-bar drag block,
  sidebar/title-bar theming, dialog theming, and window-state
  persistence load from `Lib\SuiteCommon\`, and the background runspace
  lifecycle delegates to the shared helpers. Behavior gains from the
  shared layer: hook state no longer leaks when a window closes, a
  maximized close persists the pre-maximize geometry instead of
  full-screen extents, an off-screen saved position is clamped into the
  nearest monitor instead of discarding the saved size, and teardown no
  longer blocks the UI thread while a background pipeline is stuck
  inside a provider call.

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
