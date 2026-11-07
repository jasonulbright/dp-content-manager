# Troubleshooting

This page lists the most common problems of DP Content Manager with cause and fix. See [Error messages](reference/error-messages.md) for every message.

## What it does

It gives a short fix for each common problem.

## When to use it

Use it when a refresh, a bulk action or an export does not work.

## Requirements

None.

## Problems

| Message or symptom | Cause | Fix |
|---|---|---|
| Refresh does nothing and the log asks you to open Options. | The site code or the provider is empty. | Set the connection. |
| `Refresh failed: Failed to connect to site <code> (provider <name>).` | The console is not installed on this computer, or the provider name is wrong, or `SMS_ADMIN_UI_PATH` does not point to the `bin\i386` folder of the console. | Check the provider name. Install the console. If it is installed, set `SMS_ADMIN_UI_PATH` to its `bin\i386` folder and start the tool again. |
| A view is empty. | No refresh has run. | Select **Refresh**. |
| Status Issues shows `0 issue rows`. | No (DP, content) pair is in a non-OK state. | None. |
| A bulk button opens no dialog. | The selection has no matching operations, or nothing is selected. | Select rows that are on a DP. Read the log. |
| The dialog still shows operations in progress. | A run has not ended. | Wait until the status line shows `Done.` before you select **Close**. |
| The settings are lost. | The folder is not writable. | Move the tool to a writable folder. |

## Steps

1. Read the log pane. It names the cause of most problems.
2. Read the session log in `Logs`.
3. Do the fix of the table.

## Limits

None.

## Links

- [Error messages](reference/error-messages.md)
- [Files and logs](reference/files-and-logs.md)
