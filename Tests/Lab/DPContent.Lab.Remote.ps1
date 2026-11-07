# Runs ON the lab site server (Windows PowerShell 5.1). Started by Invoke-DPContentLab.ps1.
#
# Creates one ZZTEST- package, distributes it to the lab DP, runs the module's
# write actions against it (redistribute, validate, remove, WhatIf), and removes
# everything it created, also when a step fails. It then compares the site with a
# snapshot taken before the first change. Only objects named ZZTEST-* change.
param(
    [Parameter(Mandatory)][string]$ModuleRoot,
    [Parameter(Mandatory)][string]$SiteCode,
    [Parameter(Mandatory)][string]$Provider,
    [Parameter(Mandatory)][string]$DPName
)

$ErrorActionPreference = 'Stop'
$results = [System.Collections.Generic.List[object]]::new()
function Add-Result { param([string]$Step, [bool]$Ok, [string]$Detail = '') $results.Add([PSCustomObject]@{ Step = $Step; Ok = $Ok; Detail = $Detail }) }

$pkgName = 'ZZTEST-DPCM-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
$srcDir = 'C:\ZZTEST-src-' + $pkgName.Substring(12)
$srcUnc = '\\{0}\{1}' -f $env:COMPUTERNAME, ($srcDir -replace '^C:', 'C$')
$pkgId = $null
$snapshot = $null

function Get-SiteSnapshot {
    $ns = "root\SMS\site_$SiteCode"
    $pk = @(Get-CimInstance -Namespace $ns -ClassName SMS_Package | Where-Object { $_.Name -notlike 'ZZTEST-*' } | ForEach-Object { '{0}|{1}' -f $_.PackageID, $_.Name } | Sort-Object)
    [PSCustomObject]@{ Packages = $pk }
}

try {
    Import-Module (Join-Path $ModuleRoot 'Module\DPContentMgrCommon.psd1') -Force -DisableNameChecking
    Initialize-Logging -LogPath (Join-Path $ModuleRoot 'lab.log')

    # Precondition: no other ZZTEST- package. Another run may be in progress.
    $others = @(Get-CimInstance -Namespace "root\SMS\site_$SiteCode" -ClassName SMS_Package -Filter "Name LIKE 'ZZTEST-%'")
    if ($others.Count -gt 0) { throw ('ZZTEST- package already exists: ' + (($others | ForEach-Object Name) -join ', ') + '. Another test may be running.') }

    Add-Result 'connect' ([bool](Connect-CMSite -SiteCode $SiteCode -SMSProvider $Provider)) 'Connect-CMSite'
    $snapshot = Get-SiteSnapshot

    # Create the ZZTEST package with a small source folder.
    New-Item -ItemType Directory -Path $srcDir -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $srcDir 'zztest.txt') -Value 'ZZTEST content'
    $pkg = New-CMPackage -Name $pkgName -Description 'ZZTEST temporary package; safe to delete' -Path $srcUnc
    $pkgId = $pkg.PackageID
    Add-Result 'create package' ($pkgId -match '^[A-Z0-9]{8}$') $pkgId

    Start-CMContentDistribution -PackageId $pkgId -DistributionPointName $DPName
    $row = $null
    for ($i = 0; $i -lt 40; $i++) {
        $row = @(Get-BulkDistributionStatus -SMSProvider $Provider -SiteCode $SiteCode | Where-Object PackageID -eq $pkgId) | Select-Object -First 1
        if ($row -and $row.State -in 0, 8) { break }
        Start-Sleep -Seconds 6
    }
    Add-Result 'distribute: row appears and installs' ([bool]($row -and $row.State -in 0, 8)) ($(if ($row) { $row.StatusName } else { 'no status row' }))
    Add-Result 'status row has the DP name' ([bool]($row -and $row.DPName -like "$($DPName.Split('.')[0])*")) ($(if ($row) { $row.DPName }))

    # Module functions against the live site.
    $dps = @(Get-AllDistributionPoints)
    Add-Result 'Get-AllDistributionPoints returns ServerName' ([bool]($dps.Count -ge 1 -and $dps[0].ServerName)) (($dps | ForEach-Object ServerName) -join ',')
    $content = @(Get-AllContentObjects -Types 'Package')
    Add-Result 'Get-AllContentObjects lists the ZZTEST package' ([bool]($content | Where-Object PackageID -eq $pkgId)) ''
    $status = @(Get-BulkDistributionStatus -SMSProvider $Provider -SiteCode $SiteCode)
    $rows = @(ConvertTo-DPRows -DPs $dps -ByDP @(ConvertTo-DPStatusSummary -StatusRows $status) -StorageByDP @(Get-DPStorageAnalysis -StatusRows $status -ContentObjects @(Get-AllContentObjects)))
    Add-Result 'DP grid rows carry the server name and content' ([bool]($rows | Where-Object { $_.ServerName -and $_.TotalContent -ge 1 })) (($rows | ForEach-Object { '{0}:{1}' -f $_.ServerName, $_.TotalContent }) -join ',')

    # Write actions.
    $r = @(Invoke-RedistributeContent -PackageID $pkgId -DPName $DPName)
    Add-Result 'redistribute by full name' ([bool]$r[0].Success) $r[0].Error
    $r = @(Invoke-RedistributeContent -PackageID $pkgId -DPName $DPName.Split('.')[0])
    Add-Result 'redistribute by short name' ([bool]$r[0].Success) $r[0].Error
    $shortName = $DPName.Split('.')[0]
    if ($shortName.Length -ge 2) {
        # A prefix that is shorter than the short name and is not a whole name finds no DP.
        $r = @(Invoke-RedistributeContent -PackageID $pkgId -DPName $shortName.Substring(0, $shortName.Length - 1))
        Add-Result 'redistribute by name prefix finds no DP' (-not $r[0].Success -and $r[0].Error -eq 'DP instance not found') $r[0].Error
    }
    $r = @(Invoke-ContentValidation -PackageID $pkgId -DPName $DPName)
    Add-Result 'validate' ([bool]$r[0].Success) $r[0].Error
    $r = @(Remove-ContentFromDP -PackageID $pkgId -DPName $DPName -WhatIf)
    $still = @(Get-BulkDistributionStatus -SMSProvider $Provider -SiteCode $SiteCode | Where-Object PackageID -eq $pkgId)
    Add-Result 'remove under WhatIf changes nothing' ($r.Count -eq 0 -and $still.Count -eq 1) "rows=$($still.Count)"
    $r = @(Remove-ContentFromDP -PackageID $pkgId -DPName $DPName)
    Add-Result 'remove from DP' ([bool]$r[0].Success) $r[0].Error
    $gone = $false
    for ($i = 0; $i -lt 40; $i++) {
        $left = @(Get-BulkDistributionStatus -SMSProvider $Provider -SiteCode $SiteCode | Where-Object PackageID -eq $pkgId)
        if ($left.Count -eq 0) { $gone = $true; break }
        Start-Sleep -Seconds 6
    }
    Add-Result 'status row is gone after remove' $gone ''
    $r = @(Remove-ContentFromDP -PackageID $pkgId -DPName 'no-such-dp.contoso.com')
    Add-Result 'remove from an unknown DP reports a failure' (-not $r[0].Success) $r[0].Error
}
catch {
    Add-Result 'run' $false $_.Exception.Message
}
finally {
    # Clean up everything this run created, also after a failure.
    try {
        if (-not (Test-CMConnection)) { [void](Connect-CMSite -SiteCode $SiteCode -SMSProvider $Provider) }
        # Only the package that this run created: its name holds this run's tag.
        $left = @(Get-CMPackage -Fast -Name $pkgName -ErrorAction SilentlyContinue)
        foreach ($p in $left) {
            if ($p.Name -eq $pkgName) {
                # CM raises a terminating error when the content is not on the DP (already removed).
                try { Remove-CMContentDistribution -PackageId $p.PackageID -DistributionPointName $DPName -Force -ErrorAction Stop } catch { $null = $_ }
                Remove-CMPackage -Id $p.PackageID -Force -ErrorAction Stop
            }
        }
    } catch { Add-Result 'cleanup package' $false $_.Exception.Message }
    try { if (Test-Path -LiteralPath $srcDir) { Remove-Item -LiteralPath $srcDir -Recurse -Force -ErrorAction Stop } } catch { Add-Result 'cleanup source folder' $false $_.Exception.Message }
    try {
        $remaining = @(Get-CimInstance -Namespace "root\SMS\site_$SiteCode" -ClassName SMS_Package -Filter "Name = '$pkgName'")
        Add-Result 'the ZZTEST- package of this run is gone' ($remaining.Count -eq 0) (($remaining | ForEach-Object Name) -join ',')
        if ($snapshot) {
            $after = Get-SiteSnapshot
            $same = (@($snapshot.Packages) -join "`n") -ceq (@($after.Packages) -join "`n")
            Add-Result 'other packages are unchanged' $same ''
        }
    } catch { Add-Result 'verify cleanup' $false $_.Exception.Message }
}

$results
