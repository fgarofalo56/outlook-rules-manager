<#
.SYNOPSIS
    Runs all Pester tests for Outlook Rules Manager.

.DESCRIPTION
    Executes all *.Tests.ps1 files in the tests directory and generates
    a test report with coverage information.

.PARAMETER OutputPath
    Path to save the test results file (default: ./tests/TestResults.xml)

.PARAMETER Coverage
    Enable code coverage reporting

.PARAMETER Tags
    Run only tests with specific tags

.EXAMPLE
    .\tests\Run-Tests.ps1
    # Run all tests

.EXAMPLE
    .\tests\Run-Tests.ps1 -Coverage
    # Run all tests with code coverage

.EXAMPLE
    .\tests\Run-Tests.ps1 -Tags "Security"
    # Run only tests tagged with "Security"
#>

[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot "TestResults.xml"),
    [switch]$Coverage,
    [string[]]$Tags
)

# Ensure Pester is installed
if (-not (Get-Module -ListAvailable -Name Pester | Where-Object { $_.Version -ge "5.0.0" })) {
    Write-Host "Installing Pester 5.x..." -ForegroundColor Yellow
    Install-Module Pester -MinimumVersion 5.0.0 -Force -SkipPublisherCheck -Scope CurrentUser
}

Import-Module Pester -MinimumVersion 5.0.0

# Configure Pester
$config = New-PesterConfiguration

# Test paths
$config.Run.Path = $PSScriptRoot
$config.Run.Exit = $false

# Output configuration
$config.Output.Verbosity = "Detailed"
$config.TestResult.Enabled = $true
$config.TestResult.OutputPath = $OutputPath
$config.TestResult.OutputFormat = "NUnitXml"

# Tags filter
if ($Tags) {
    $config.Filter.Tag = $Tags
}

# Code coverage
if ($Coverage) {
    $config.CodeCoverage.Enabled = $true
    $config.CodeCoverage.Path = @(
        (Join-Path $PSScriptRoot "..\src\modules\SecurityHelpers.psm1")
    )
    $config.CodeCoverage.OutputPath = (Join-Path $PSScriptRoot "coverage.xml")
    $config.CodeCoverage.OutputFormat = "JaCoCo"
}

# Run tests
Write-Host "`n=== Running Pester Tests ===" -ForegroundColor Cyan
Write-Host "Test Path: $PSScriptRoot" -ForegroundColor Gray
Write-Host "Output: $OutputPath" -ForegroundColor Gray
if ($Coverage) {
    Write-Host "Coverage: Enabled" -ForegroundColor Gray
}
Write-Host ""

$result = Invoke-Pester -Configuration $config

# Summary
Write-Host "`n=== Test Summary ===" -ForegroundColor Cyan
Write-Host "Passed:  $($result.PassedCount)" -ForegroundColor Green
Write-Host "Failed:  $($result.FailedCount)" -ForegroundColor $(if ($result.FailedCount -gt 0) { "Red" } else { "Gray" })
Write-Host "Skipped: $($result.SkippedCount)" -ForegroundColor Gray
Write-Host "Total:   $($result.TotalCount)" -ForegroundColor White

if ($result.FailedCount -gt 0) {
    Write-Host "`nFailed Tests:" -ForegroundColor Red
    $result.Failed | ForEach-Object {
        Write-Host "  - $($_.Name)" -ForegroundColor Red
    }
    exit 1
}

Write-Host "`nAll tests passed!" -ForegroundColor Green
exit 0
