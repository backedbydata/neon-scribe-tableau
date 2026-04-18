# Script to compare expected vs actual parsing results
# Useful for validating your TWB parser implementation

param(
    [Parameter(Mandatory=$true)]
    [string]$TestFile,
    [Parameter(Mandatory=$true)]
    [string]$ParserOutputJson,
    [switch]$Verbose
)

# This script expects your parser to output a JSON file with the following structure:
<#
{
    "worksheets": 4,
    "dashboards": 1,
    "dataSources": 2,
    "calculatedFields": 6,
    "filters": 3,
    "actions": 4,
    "parameters": 3,
    "lodCalculations": 2,
    "fields": [
        {
            "name": "Sales Amount",
            "internalName": "[Sales Amount]",
            "dataType": "real",
            "role": "measure"
        }
    ]
}
#>

Write-Host "`nTableau Parser Test Comparison" -ForegroundColor Cyan
Write-Host "==============================`n" -ForegroundColor Cyan

# Check if files exist
if (-not (Test-Path $TestFile)) {
    Write-Error "Test file not found: $TestFile"
    exit 1
}

if (-not (Test-Path $ParserOutputJson)) {
    Write-Error "Parser output JSON not found: $ParserOutputJson"
    exit 1
}

# Run the analyzer on the test file to get expected values
Write-Host "Analyzing test file to get expected values..." -ForegroundColor Yellow
$analyzerScript = Join-Path $PSScriptRoot "analyze-twb.ps1"
$tempCsv = Join-Path $env:TEMP "analyzer-output-$([guid]::NewGuid()).csv"

& $analyzerScript -Path $TestFile -OutputCsv $tempCsv | Out-Null

if (-not (Test-Path $tempCsv)) {
    Write-Error "Failed to analyze test file"
    exit 1
}

# Read expected values
$expected = Import-Csv $tempCsv | Select-Object -First 1
Remove-Item $tempCsv -Force

# Read actual values from parser
try {
    $actual = Get-Content $ParserOutputJson -Raw | ConvertFrom-Json
}
catch {
    Write-Error "Failed to parse JSON output: $_"
    exit 1
}

# Comparison results
$results = @{
    Passed = @()
    Failed = @()
    Missing = @()
}

# Define test cases
$testCases = @(
    @{
        Name = "Worksheets"
        Expected = [int]$expected.Worksheets
        Actual = $actual.worksheets
        Critical = $true
    },
    @{
        Name = "Dashboards"
        Expected = [int]$expected.Dashboards
        Actual = $actual.dashboards
        Critical = $true
    },
    @{
        Name = "Data Sources"
        Expected = [int]$expected.DataSourcesWithData
        Actual = $actual.dataSources
        Critical = $true
    },
    @{
        Name = "Parameters"
        Expected = [int]$expected.Parameters
        Actual = $actual.parameters
        Critical = $false
    },
    @{
        Name = "Calculated Fields"
        Expected = [int]$expected.CalculatedFields
        Actual = $actual.calculatedFields
        Critical = $true
    },
    @{
        Name = "LOD Calculations"
        Expected = [int]$expected.LODCalculations
        Actual = $actual.lodCalculations
        Critical = $false
    },
    @{
        Name = "Total Filters"
        Expected = [int]$expected.TotalFilters
        Actual = $actual.filters
        Critical = $true
    },
    @{
        Name = "Total Actions"
        Expected = [int]$expected.TotalActions
        Actual = $actual.actions
        Critical = $false
    },
    @{
        Name = "Total Fields"
        Expected = [int]$expected.TotalFields
        Actual = $actual.totalFields
        Critical = $false
    }
)

# Run comparisons
Write-Host "Test Results:" -ForegroundColor Cyan
Write-Host "============`n" -ForegroundColor Cyan

$totalTests = 0
$passedTests = 0
$failedTests = 0
$missingTests = 0

foreach ($test in $testCases) {
    $totalTests++
    $testName = $test.Name
    $expected = $test.Expected
    $actual = $test.Actual
    $critical = $test.Critical

    $status = ""
    $color = "White"
    $symbol = ""

    if ($null -eq $actual) {
        $status = "MISSING"
        $color = "Yellow"
        $symbol = "?"
        $results.Missing += $test
        $missingTests++
    }
    elseif ($expected -eq $actual) {
        $status = "PASS"
        $color = "Green"
        $symbol = "✓"
        $results.Passed += $test
        $passedTests++
    }
    else {
        $status = "FAIL"
        $color = if ($critical) { "Red" } else { "Yellow" }
        $symbol = "✗"
        $results.Failed += $test
        $failedTests++
    }

    $criticalMarker = if ($critical) { "[CRITICAL]" } else { "" }

    Write-Host ("[{0}] {1,-20} Expected: {2,3}  Actual: {3,3}  {4}" -f $symbol, $testName, $expected, $(if ($null -eq $actual) { "N/A" } else { $actual }), $criticalMarker) -ForegroundColor $color

    if ($Verbose -and $status -eq "FAIL") {
        Write-Host "    Difference: $($actual - $expected)" -ForegroundColor Gray
    }
}

# Field-level validation (if provided)
if ($actual.fields) {
    Write-Host "`nField Validation:" -ForegroundColor Cyan
    Write-Host "=================" -ForegroundColor Cyan

    $fieldCount = $actual.fields.Count
    $expectedFieldCount = [int]$expected.TotalFields

    Write-Host "Fields provided: $fieldCount / $expectedFieldCount" -ForegroundColor $(if ($fieldCount -eq $expectedFieldCount) { "Green" } else { "Yellow" })

    if ($Verbose) {
        Write-Host "`nField Details:" -ForegroundColor Yellow
        foreach ($field in $actual.fields) {
            $hasCaption = -not [string]::IsNullOrEmpty($field.name)
            $captionStatus = if ($hasCaption) { "✓" } else { "✗" }
            $captionColor = if ($hasCaption) { "Green" } else { "Red" }

            Write-Host "  [$captionStatus] $($field.name)" -ForegroundColor $captionColor
            if ($Verbose) {
                Write-Host "      Internal: $($field.internalName)" -ForegroundColor Gray
                Write-Host "      Type: $($field.dataType), Role: $($field.role)" -ForegroundColor Gray
            }
        }
    }

    # Check for cryptic names
    $crypticFields = $actual.fields | Where-Object {
        $_.name -match 'calculation_\d+|federated\.\w+\.\w+'
    }

    if ($crypticFields) {
        Write-Host "`n[WARNING] Found $($crypticFields.Count) fields with cryptic names:" -ForegroundColor Yellow
        foreach ($field in $crypticFields) {
            Write-Host "  - $($field.name)" -ForegroundColor Yellow
        }
        Write-Host "  These should be resolved to friendly names using caption attributes" -ForegroundColor Gray
    }
}

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Total Tests: $totalTests" -ForegroundColor White
Write-Host "Passed: $passedTests" -ForegroundColor Green
Write-Host "Failed: $failedTests" -ForegroundColor $(if ($failedTests -gt 0) { "Red" } else { "White" })
Write-Host "Missing: $missingTests" -ForegroundColor $(if ($missingTests -gt 0) { "Yellow" } else { "White" })

$passRate = [Math]::Round(($passedTests / $totalTests) * 100, 1)
Write-Host "`nPass Rate: $passRate%" -ForegroundColor $(
    if ($passRate -eq 100) { "Green" }
    elseif ($passRate -ge 80) { "Yellow" }
    else { "Red" }
)

# Critical failures
$criticalFailures = $results.Failed | Where-Object { $_.Critical }
if ($criticalFailures) {
    Write-Host "`n[CRITICAL FAILURES]" -ForegroundColor Red
    foreach ($failure in $criticalFailures) {
        Write-Host "  - $($failure.Name): Expected $($failure.Expected), got $($failure.Actual)" -ForegroundColor Red
    }
}

# Exit code
if ($criticalFailures) {
    Write-Host "`nTests FAILED with critical errors" -ForegroundColor Red
    exit 1
}
elseif ($failedTests -gt 0 -or $missingTests -gt 0) {
    Write-Host "`nTests completed with warnings" -ForegroundColor Yellow
    exit 0
}
else {
    Write-Host "`nAll tests PASSED!" -ForegroundColor Green
    exit 0
}
