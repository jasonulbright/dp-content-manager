# Permissions

This page lists the rights that DP Content Manager needs.

## What it does

A refresh reads the site. The bulk actions change content on distribution points with the rights of the account that started the tool.

## When to use it

Use it to set up an account, or to find out why a read or an action is refused.

## Requirements

| Action | Right |
|---|---|
| Start the tool | Read and write in the folder of the script. |
| **Refresh** and all views | The Configuration Manager console on the computer, and read access to the SMS Provider and to DP content under your Configuration Manager role. |
| **Redistribute Failed...**, **Redistribute Selected...** | The right to redistribute content to distribution points. |
| **Remove From DP...**, **Remove Failures...** | The right to remove content from distribution points. |
| **Validate...** | The right to validate content on distribution points. |

## Steps

1. Start the tool as the account that has the Configuration Manager role.
2. Select **Refresh** to test the read access.

## Limits

- The tool does not check the rights before an action. A refused operation shows as FAIL in the dialog.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| `Refresh failed:` | The account has no access to the provider, or the console is missing. | Check the role, the provider name and the console. |
| A row shows FAIL. | The action failed. | Read the log file for the reason. Check the permissions only when it shows an access failure. |

## Links

- [Network](network.md)
- [Bulk action dialog](../core/bulk-actions.md)
