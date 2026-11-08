# DP Content Manager: test plan

Measured 2026-11-06 with `Invoke-Checks.ps1` (Pester 5.7.1, PSScriptAnalyzer 1.25.0). Counts are tests as run; each data row counts as one test. Coverage is the share of module commands that the tests run.

| | Before | After |
|---|---|---|
| Tests | 20 | 442 on Windows PowerShell 5.1 and on PowerShell 7 (11 contract tests skip on a machine without the console) |
| Module coverage | 31.2% (138 of 442 commands) | 99.5% (597 of 600 commands) |
| Lab steps | none | 17 (`Tests/Lab`) |

Run: `.\Invoke-Checks.ps1` (lint, tests, coverage floor 80%). It runs in CI on both hosts (`.github/workflows/checks.yml`).

Window tests: `Window.Tests.ps1` opens windows only when `SUITE_ALLOW_WINDOW_TESTS` is `1`. The CI workflow sets it. To run the window tests by hand, set it first: `$env:SUITE_ALLOW_WINDOW_TESTS = '1'; .\Invoke-Checks.ps1`. The lab script `Tests/Lab/Invoke-DPContentLab.ps1` does not set it and does not run `Window.Tests.ps1`. Otherwise Pester reports those tests as Skipped with that reason, not as Failed. `Test-WindowTestsAllowed` in `Tests/WindowHarness.ps1` makes the decision, and `Get-WpfSession` refuses to create a runspace when it says no. `Invoke-Checks.ps1` prints one line when it skips them, and the run still passes. The coverage floor stays in force: the window tests run no module code, so skipping them does not lower the module coverage. `WindowGuard.Tests.ps1` tests the guard with injected opt-in values, tests the environment default when no opt-in is passed, and checks that every test file that can open a window carries it. It opens no window.

## What is covered

| Area | File | Contents |
|---|---|---|
| Content retrieval, all 11 `Get-*` functions | `Module.Retrieval.Tests.ps1` | Property mapping, empty site, hostile names (quotes, unicode, 5000 characters, path traps), 2000 objects, cmdlet failure, type filter |
| Status queries and rollups | `Module.Status.Tests.ps1` | All 9 states and unknown states, DP name from the NAL path, percentage rounding, empty input, 20000 rows, orphans, storage |
| Write actions: redistribute, remove, validate | `Module.Actions.Tests.ps1` | Success, failure, partial failure, progress callback, `-WhatIf`, empty and null arguments, hostile DP names, exact DP match, WQL escaping, `Assert-ContentActionResult` |
| Export and summary | `Module.Export.Tests.ps1` | CSV and HTML round trips, HTML encoding, column classes, empty table, folders that do not exist |
| Row shaping for the window | `Module.Rows.Tests.ps1` | Glyphs, size text, DP, content and issue rows, orphans, 32-bit overflow |
| Script logic | `Shell.Logic.Tests.ps1` | Filters for all three views, status bar text, export routing, preferences file (first run, corrupt, unknown keys, round trip), script contract (every call resolves, every action result is checked) |
| Start without the console | `Startup.Tests.ps1` | Plain message, `$false`, no raw exception |
| Windows | `Window.Tests.ps1` | Main window and both dialogs: every named control, titles, no clipped text and nothing off screen at 1024x720 and 1440x900 in both themes and all three views, text contrast, control contrast |
| Real cmdlet signatures | `CmCmdlet.Contract.Tests.ps1` | Every ConfigMgr call, with the keys of a splatted hashtable, names only parameters that the real cmdlet has. Runs where the console is installed |
| Lab | `Lab/Invoke-DPContentLab.ps1` | Creates one `ZZTEST-` package on a lab machine, distributes it, redistributes (full name, short name, prefix), validates, removes under `-WhatIf` and for real, removes the package, checks that no other package changed |

## Defects found

| # | What | Fix |
|---|---|---|
| 1 | `Invoke-CMContentValidation -Id` does not exist, so every validation failed | `-PackageId` |
| 2 | The operation dialog ignored `Success = $false`, so a failed action showed OK | `Assert-ContentActionResult` on the five action calls |
| 3 | `LIKE '%dp1%'` also redistributed dp10 and dp11 | Exact server name match, or short name |
| 4 | The DP query returned `Name`, the grid read `ServerName`, so the DP names were empty | Both properties |
| 5 | The rollup functions threw on an empty status list | `AllowEmptyCollection` |
| 6 | A quote in the package ID broke the redistribute query | WQL escaping |

## Not covered, and why

- Contrast of disabled controls, the sidebar button border and the light text box border fails the standard (text 4.5:1, borders 3:1). The cause is the shared SuiteCommon palette and the MahApps disabled style. The window test lists each violation in `Tests/KnownContrastDefects.ps1`, fails on a new one, and fails when a listed one is fixed. The fix belongs in `app-packager-suite`.
- The refresh timer, the busy overlay and the background runspace live inside `Invoke-Refresh`. They need a running window with a site and are covered by the lab run only.
- The settings save in SuiteCommon is not atomic (`Set-Content` on the target file). The fix belongs in `app-packager-suite`.
- Dialog windows are measured for structure and titles only. Their layout and contrast need the dialog to be shown with data.
- The title bar of an inactive window misses 4.5:1 (listed as `titlebar:inactive`). The border of a disabled sidebar button is read from the button, not from the template border at half opacity.
- Only the Content-view buttons Redistribute, Remove and Validate run against a real site in the lab test, through the module. The window click path is not driven.
