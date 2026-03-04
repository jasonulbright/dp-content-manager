# Changelog

All notable changes to the DP Content Manager are documented in this file.

## [1.0.2] - 2026-03-04

### Fixed
- Redistribute action now uses WMI `RefreshNow` on `SMS_DistributionPoint` (same mechanism as the CM console) -- `Start-CMContentDistribution` only works for initial distribution and failed on content already present on a DP
- Content validation parameter corrected from `-PackageId` to `-Id` on `Invoke-CMContentValidation`
- Dock Z-order: reversed `BringToFront()` sequence so header panel renders at top instead of below the filter bar

---

## [1.0.1] - 2026-02-26

### Fixed

- Added ClearType anti-aliasing (`TextRenderingHint.ClearTypeGridFit`) to owner-draw tab headers for smoother text rendering in dark mode

## [1.0.0] - 2026-02-26

### Added

- **GUI application** (`start-dpcontentmgr.ps1`) with WinForms interface
  - Header panel with title and subtitle
  - Connection bar showing site code, SMS provider, and connection status
  - Content type and status filter dropdowns with real-time text filter
  - TabControl with Content View and DP View tabs
  - DataGridView with color-coded rows (red = failed, orange = in progress)
  - Detail panels with RichTextBox info and per-DP/per-content sub-grids
  - Live log console with timestamped progress messages
  - Status bar with summary counts

- **Content View tab**
  - Rows: content objects (apps, packages, SUDPs, boot images, OS images, driver packages, task sequences)
  - Columns: Name, Type, Package ID, Size, Total DPs, Installed, In Progress, Failed, % Complete
  - Drill-down: per-DP status sub-grid for selected content

- **DP View tab**
  - Rows: distribution points
  - Columns: DP Name, Site, Pull DP, Total Content, Installed, In Progress, Failed, Size (GB), % Complete
  - Drill-down: content sub-grid for selected DP

- **Bulk status queries** via WMI `SMS_PackageStatusDistPointsSummarizer` for scale (300+ DPs)

- **Actions**
  - Redistribute failed content to failed DPs
  - Validate content integrity (server-side hash check)
  - Remove content from distribution points
  - Detect and remove orphaned content (content on DPs with no source object)

- **Storage analysis** modal dialog with per-DP size breakdown and CSV export

- **Dark mode** with full theme support
  - Custom `DarkToolStripRenderer` for MenuStrip and StatusStrip
  - Owner-draw TabControl headers
  - Configurable via File > Preferences
  - Persisted in `DPContentMgr.prefs.json`

- **Export**
  - CSV export of active tab grid data
  - HTML export with color-coded status cells and embedded CSS

- **Menu bar** with File (Preferences, Exit), Actions (Refresh, Redistribute, Validate, Remove Orphaned), View (Content View, DP View, Storage Analysis), Help (About)

- **Window state persistence** across sessions (`DPContentMgr.windowstate.json`)

- **Core module** (`DPContentMgrCommon.psm1`) with 27 exported functions
  - CM site connection management
  - Content retrieval for 7 content types
  - Bulk WMI status queries with O(n) hashtable aggregation
  - Redistribution, removal, and validation actions
  - Orphaned content detection
  - Storage analysis
  - CSV and HTML export
