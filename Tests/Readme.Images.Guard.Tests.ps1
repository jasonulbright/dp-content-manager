#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
# Headless guard: every image of the README is a docs image or a mapped copy of one, and the copy is current.

BeforeDiscovery {
    $script:Root = Split-Path -Parent $PSScriptRoot
    $readme = Get-Content -LiteralPath (Join-Path $script:Root 'README.md') -Raw
    $script:Refs = @([regex]::Matches($readme, '!\[[^\]]*\]\(([^)\s]+\.(?:png|gif|jpg))\)') | ForEach-Object { @{ Ref = $_.Groups[1].Value } })
    $mapPath = Join-Path $script:Root 'tools\guide-capture\readme-images.json'
    $map = if (Test-Path -LiteralPath $mapPath) { Get-Content -LiteralPath $mapPath -Raw | ConvertFrom-Json } else { $null }
    $script:Pairs = @(if ($map) { $map.PSObject.Properties | ForEach-Object { @{ Target = $_.Name; Source = $_.Value } } })
}

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    $mapPath = Join-Path $script:Root 'tools\guide-capture\readme-images.json'
    $script:Map = Get-Content -LiteralPath $mapPath -Raw | ConvertFrom-Json
}

Describe 'README images' {
    It 'finds at least one image in the README' {
        @($script:Refs).Count | Should -BeGreaterThan 0
    }

    It '<Ref> is under docs/images or listed in readme-images.json' -ForEach $script:Refs {
        $inDocs = $Ref -like 'docs/images/*'
        $mapped = $null -ne $script:Map.PSObject.Properties[$Ref]
        ($inDocs -or $mapped) | Should -BeTrue
        (Test-Path -LiteralPath (Join-Path $script:Root $Ref)) | Should -BeTrue
    }

    It '<Target> is a byte-identical copy of <Source>' -ForEach $script:Pairs {
        (Test-Path -LiteralPath (Join-Path $script:Root $Source)) | Should -BeTrue
        (Get-FileHash -LiteralPath (Join-Path $script:Root $Target) -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath (Join-Path $script:Root $Source) -Algorithm SHA256).Hash
    }
}
