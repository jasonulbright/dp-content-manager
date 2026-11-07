# Shared helpers for the DP Content Manager tests. Dot-source from BeforeAll.

$script:DpcmRoot = Split-Path -Parent $PSScriptRoot

function Import-DpcmModule {
    Import-Module (Join-Path $script:DpcmRoot 'Module\DPContentMgrCommon.psd1') -Force -DisableNameChecking
}

# Defines a global stub for each ConfigMgr command that the machine does not
# have, so Pester can mock it. Returns the names it created; pass them to
# Remove-CmStub afterwards so the stubs never outlive the test file.
function New-CmStub {
    param(
        [Parameter(Mandatory)][string[]]$Name,
        # The stub accepts any named parameter and pipeline input, so a mock of
        # it binds whatever the code under test passes.
        [scriptblock]$Definition = { [CmdletBinding()] param([Parameter(ValueFromPipeline)]$InputObject, [Parameter(ValueFromRemainingArguments)]$Rest) process { } }
    )
    $created = @()
    foreach ($n in $Name) {
        if (-not (Get-Command -Name $n -ErrorAction SilentlyContinue)) {
            Set-Item -Path "function:global:$n" -Value $Definition
            $created += $n
        }
    }
    return $created
}

function Remove-CmStub {
    param([string[]]$Name)
    foreach ($n in $Name) { Remove-Item -Path "function:global:$n" -ErrorAction SilentlyContinue }
}

function New-DpNalPath {
    param([Parameter(Mandatory)][string]$Server, [string]$Site = 'MCM')
    return ('["Display=\\{0}\"]MSWNET:["SMS_SITE={1}"]\\{0}\' -f $Server, $Site)
}

function New-StatusRow {
    param([string]$PackageID, [string]$DPName, [int]$State, [string]$StatusName = 'Generated')
    [PSCustomObject]@{ PackageID = $PackageID; DPName = $DPName; State = $State; StatusName = $StatusName }
}

function New-RollupTable {
    # DataTable with the column names the shell passes to the export functions.
    param([string[]]$Column, [object[][]]$Row)
    $t = New-Object System.Data.DataTable
    foreach ($c in $Column) { [void]$t.Columns.Add($c) }
    foreach ($r in $Row) { [void]$t.Rows.Add($r) }
    # The comma keeps PowerShell from unrolling the table into its rows.
    return , $t
}
