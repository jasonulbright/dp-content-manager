# Files and logs

This page lists every folder and file that DP Content Manager writes.

## What it does

The tool writes global files in the folder of the script and site files in the data folder of the site. It writes nothing in the registry. When the tool is installed by the suite installer, the daily update check also keeps a cache file, `%LOCALAPPDATA%\AppPackagerSuite\update-check.json`, outside the script folder. A plain extract of the zip does not create it. A bulk action that you run changes the site, not these files.

## When to use it

Use it to find a log or a report, or to clean up.

## Requirements

None.

## Files

| Path | Content | Written when |
|---|---|---|
| `DPContentMgr.prefs.json` | The global settings (the theme). With no site chosen, also the site code and the SMS Provider. With a site, the site code and the provider in this file are old copies that the tool no longer changes. | You select **OK** in Options, or change the theme. |
| `DPContentMgr.windowstate.json` | The size, the place and the open view of the window. | You close the window. |
| `Logs\DPCM-<yyyyMMdd-HHmmss>.log` | The log of one session. | During the session. |
| `Logs\DPCM-startup-<yyyyMMdd-HHmmss>.txt` | The PowerShell transcript of the start. | At the start. |
| `Logs\DPCM-crash-<yyyyMMdd-HHmmss>.txt` | The error text when the window closes on an unhandled error. | An unhandled error. |
| `%LOCALAPPDATA%\AppPackagerData\sites\<SITE>\tools\dp-content-manager\prefs.json` | The site code and the SMS Provider of the site. | You select **OK** in Options, or change the theme. |
| `%LOCALAPPDATA%\AppPackagerData\sites\<SITE>\tools\dp-content-manager\reports\DPCM-<view>-<date>-<time>.csv` or `.html` | An export of the active view. | **Export CSV** or **Export HTML**. |

With no site chosen, the tool uses `Reports\` beside the script for the exports and keeps the site settings in `DPContentMgr.prefs.json`. The first start for a site copies the old `Reports` files into the site folder and keeps the old files. See [Sites](../core/options.md#sites).

## Steps

### Clean up

1. Close the tool.
2. Delete the logs and the reports that you do not need. The reports of a site are in the `reports` folder of that site.

## Limits

- The tool never deletes a log or a report.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| A start leaves no log. | The folder is not writable. | Move the tool to a writable folder. |

## Links

- [Settings](settings.md)
- [DPs](../core/dps.md)
