#Requires -Modules Pester
# Tests of the window-test guard (Test-WindowTestsAllowed in WindowHarness.ps1).
# The guard runs window tests only when SUITE_ALLOW_WINDOW_TESTS is 1. This file injects the
# opt-in value, so it never reads the real environment and never opens a window.

BeforeDiscovery {
    # Every other test file that can open a window must carry the guard.
    $script:WindowTestFiles = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Tests.ps1' -File |
            Where-Object { $_.Name -ne 'WindowGuard.Tests.ps1' -and (Get-Content -LiteralPath $_.FullName -Raw) -match 'Invoke-WpfSta|New-WpfWindow|ShowDialog|\.Show\(\)' } |
            ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    function Get-Guard { param($Value) Test-WindowTestsAllowed -OptIn $Value }
}

Describe 'Test-WindowTestsAllowed' {
    It 'has no ComputerName parameter and decides on the opt-in only' {
        $params = (Get-Command Test-WindowTestsAllowed).Parameters.Keys
        $params | Should -Not -Contain 'ComputerName'
        $params | Should -Contain 'OptIn'
        (Test-WindowTestsAllowed -OptIn '1').Allowed | Should -BeTrue
        (Test-WindowTestsAllowed -OptIn '').Allowed | Should -BeFalse
    }

    It 'allows the opt-in 1' {
        $r = Get-Guard '1'
        $r.Allowed | Should -BeTrue
        $r.Reason | Should -BeNullOrEmpty
    }

    It 'refuses when the opt-in is <Label>, and names the opt-in' -ForEach @(
        @{ Label = 'empty'; Value = '' }
        @{ Label = 'null'; Value = $null }
        @{ Label = '0'; Value = '0' }
        @{ Label = 'true'; Value = 'true' }
    ) {
        $r = Get-Guard $Value
        $r.Allowed | Should -BeFalse
        $r.Reason | Should -Be 'SUITE_ALLOW_WINDOW_TESTS is not 1'
    }
}

Describe 'Test-WindowTestsAllowed environment default' {
    BeforeAll {
        $script:SavedOptIn = $env:SUITE_ALLOW_WINDOW_TESTS
    }
    AfterAll {
        $env:SUITE_ALLOW_WINDOW_TESTS = $script:SavedOptIn
        if ($null -eq $script:SavedOptIn) { Remove-Item Env:\SUITE_ALLOW_WINDOW_TESTS -ErrorAction SilentlyContinue }
    }
    AfterEach {
        Remove-Item Env:\SUITE_ALLOW_WINDOW_TESTS -ErrorAction SilentlyContinue
    }

    It 'allows when the environment value is 1 and no opt-in is passed' {
        $env:SUITE_ALLOW_WINDOW_TESTS = '1'
        $r = Test-WindowTestsAllowed
        $r.Allowed | Should -BeTrue
        $r.Reason | Should -BeNullOrEmpty
    }

    It 'refuses when the environment value is <Label> and no opt-in is passed' -ForEach @(
        @{ Label = '0'; Value = '0' }
        @{ Label = 'true'; Value = 'true' }
        @{ Label = 'empty'; Value = '' }
    ) {
        $env:SUITE_ALLOW_WINDOW_TESTS = $Value
        $r = Test-WindowTestsAllowed
        $r.Allowed | Should -BeFalse
        $r.Reason | Should -Be 'SUITE_ALLOW_WINDOW_TESTS is not 1'
    }

    It 'refuses when the environment variable is not set and no opt-in is passed' {
        Remove-Item Env:\SUITE_ALLOW_WINDOW_TESTS -ErrorAction SilentlyContinue
        (Test-WindowTestsAllowed).Allowed | Should -BeFalse
    }
}

Describe 'Window test files carry the guard' {
    It 'finds at least one test file that can open a window' -ForEach @(@{ Found = @($script:WindowTestFiles).Count }) {
        $Found | Should -BeGreaterThan 0
    }

    It 'prints the skip reason once: Invoke-Checks sets the marker and Window.Tests reads it' {
        (Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-Checks.ps1') -Raw) | Should -Match '\$env:SUITE_WINDOW_GUARD_REPORTED\s*=\s*''1'''
        (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Window.Tests.ps1') -Raw) | Should -Match '-not \$env:SUITE_WINDOW_GUARD_REPORTED'
    }

    It '<Name> calls Test-WindowTestsAllowed and applies -Skip' -ForEach $script:WindowTestFiles {
        $text = Get-Content -LiteralPath $Path -Raw
        $text | Should -Match 'Test-WindowTestsAllowed'
        $text | Should -Match '-Skip:'
    }
}
