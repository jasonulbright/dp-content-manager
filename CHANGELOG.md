# Changelog

All notable changes to DP Content Manager are documented in this file.

## [1.0.0] - 2026-05-02

DP Content Manager is a MahApps.Metro WPF GUI for auditing, redistributing,
validating, and removing MECM distribution point content at scale. Ships as
a zip + `install.ps1` wrapper; no MSI, no code signing required.

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
  detection / storage analysis, and CSV / HTML export. 20 Pester 5.x
  tests in `Module/DPContentMgrCommon.Tests.ps1` covering pure-logic
  exports (status aggregation, orphan detection, storage math, summary).
- **WPF brand alignment** -- MahApps.Metro shell with sidebar navigation,
  glyph status (no red / green for state), animated ProgressRing during
  refresh, log drawer, status bar, dark and light themes with runtime
  toggle, window-state persistence with legacy WinForms schema bridge.

### Stack

- PowerShell 5.1 + .NET Framework 4.7.2+
- WPF + MahApps.Metro 2.4.10 (vendored under `Lib/`)
- ConfigurationManager PowerShell module + raw CIM for the bulk status
  summarizer (CM console required on the host machine)
