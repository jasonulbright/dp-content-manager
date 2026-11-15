<#
.SYNOPSIS
    Copies the captured guide images over the images that the README shows.

.DESCRIPTION
    readme-images.json maps each README image path to the docs\images file it is a copy of.
    Run the guide capture first, then this script. It stops with an error when a source image is missing.

.PARAMETER RepoRoot
    The repository folder. The default is the repository that holds this script.
#>
[CmdletBinding()]
param(
    [string]$RepoRoot
)
$ErrorActionPreference = 'Stop'
if (-not $RepoRoot) { $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path }
$map = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'readme-images.json') -Raw | ConvertFrom-Json
foreach ($entry in $map.PSObject.Properties) {
    $source = Join-Path $RepoRoot $entry.Value
    $target = Join-Path $RepoRoot $entry.Name
    if (-not (Test-Path -LiteralPath $source)) { throw "The source image is missing: $($entry.Value)" }
    Copy-Item -LiteralPath $source -Destination $target -Force
    Write-Host ('{0} <- {1}' -f $entry.Name, $entry.Value)
}
