# Glossary

This page defines the terms of DP Content Manager.

## What it does

It explains the terms that the tool and this guide use.

## When to use it

Use it when a term is not clear.

## Requirements

None.

## Terms

| Term | Meaning |
|---|---|
| Distribution point (DP) | A site system that holds content for clients. |
| Pull DP | A pull-distribution point. It gets its content from another distribution point. |
| Content object | An application, package, software update deployment package (SUDP), boot image, OS image, driver package or task sequence content. |
| (DP, content) pair | The state of one content object on one distribution point. |
| State | The distribution state of a pair, for example installed, in progress or failed. |
| Redistribute | Sends the content to the distribution point again. |
| Validate | Asks the distribution point to check the content hash. |
| Preview | The dialog that lists every pair before an action runs. |
| Status summarizer | The class `SMS_PackageStatusDistPointsSummarizer`. One query on it gives the status of all pairs. |
| SMS Provider | The server role that gives access to the site data. |
| Site folder | The folder `%LOCALAPPDATA%\AppPackagerData\sites\<SITE>\tools\dp-content-manager`. It holds the settings and the reports of one site. |

## Steps

None.

## Limits

None.

## Troubleshooting

None.

## Links

- [Content](../core/content.md)
