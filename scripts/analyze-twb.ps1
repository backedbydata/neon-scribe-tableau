# PowerShell script to analyze Tableau TWB/TWBX files and report complexity metrics

param(
    [Parameter(Mandatory=$true)]
    [string]$Path,
    [switch]$Recursive,
    [switch]$Detailed,
    [string]$OutputCsv
)

# Function to extract TWB from TWBX
function Extract-TwbFromTwbx {
    param([string]$TwbxPath)

    $tempDir = Join-Path $env:TEMP ([System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    try {
        # TWBX is just a zip file
        Expand-Archive -Path $TwbxPath -DestinationPath $tempDir -Force

        # Find the TWB file inside
        $twbFile = Get-ChildItem -Path $tempDir -Filter "*.twb" -Recurse | Select-Object -First 1

        if ($twbFile) {
            return $twbFile.FullName, $tempDir
        }
        else {
            Write-Warning "No TWB file found in TWBX: $TwbxPath"
            return $null, $tempDir
        }
    }
    catch {
        Write-Warning "Failed to extract TWBX: $_"
        return $null, $tempDir
    }
}

# Function to analyze a single TWB file
function Analyze-TwbFile {
    param(
        [string]$FilePath,
        [string]$OriginalPath
    )

    try {
        [xml]$xml = Get-Content -Path $FilePath -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to parse XML in $OriginalPath : $_"
        return $null
    }

    $analysis = [PSCustomObject]@{
        FileName = [System.IO.Path]::GetFileName($OriginalPath)
        FilePath = $OriginalPath
        FileSize = (Get-Item $OriginalPath).Length
        Version = $xml.workbook.version
        SourceBuild = $xml.workbook.'source-build'

        # Counts
        DataSources = 0
        DataSourcesWithData = 0
        Parameters = 0
        Worksheets = 0
        Dashboards = 0
        Stories = 0

        # Field counts
        TotalFields = 0
        Dimensions = 0
        Measures = 0
        CalculatedFields = 0
        LODCalculations = 0
        TableCalculations = 0

        # Filter counts
        TotalFilters = 0
        CategoricalFilters = 0
        QuantitativeFilters = 0
        DateFilters = 0
        TopNFilters = 0
        ContextFilters = 0

        # Action counts
        TotalActions = 0
        FilterActions = 0
        HighlightActions = 0
        URLActions = 0
        ParameterActions = 0

        # Visual elements
        MarkTypes = @()
        ColorEncodings = 0
        HasTooltips = $false
        HasCustomShapes = $false

        # Dashboard elements
        DashboardZones = 0
        ParameterControls = 0
        FilterControls = 0
        TextBoxes = 0
        Images = 0
        WebObjects = 0

        # Data source details
        ConnectionTypes = @()
        HasCustomSQL = $false

        # Complexity indicators
        HasAliases = $false
        HasHierarchies = $false
        MaxCalculationNesting = 0

        # Warnings/Issues
        Warnings = @()
    }

    # Analyze datasources
    $datasources = $xml.workbook.datasources.datasource
    if ($datasources) {
        $analysis.DataSources = @($datasources).Count

        foreach ($ds in $datasources) {
            # Check if it's a parameter datasource or real data
            if ($ds.connection.class -eq 'parameters') {
                # Count parameters
                $params = $ds.column | Where-Object { $_.GetAttribute('param-domain-type') }
                if ($params) {
                    $analysis.Parameters += @($params).Count
                }
            }
            else {
                $analysis.DataSourcesWithData++

                # Connection types
                $connType = $ds.connection.class
                if ($connType -and $connType -notin $analysis.ConnectionTypes) {
                    $analysis.ConnectionTypes += $connType
                }

                # Check for custom SQL
                if ($ds.connection.relation.sql) {
                    $analysis.HasCustomSQL = $true
                }

                # Analyze columns/fields
                $columns = $ds.column
                if ($columns) {
                    foreach ($col in $columns) {
                        $analysis.TotalFields++

                        # Role (dimension/measure)
                        if ($col.role -eq 'dimension') { $analysis.Dimensions++ }
                        if ($col.role -eq 'measure') { $analysis.Measures++ }

                        # Calculated fields
                        if ($col.calculation) {
                            $analysis.CalculatedFields++

                            $formula = $col.calculation.formula
                            if ($formula) {
                                # LOD calculations
                                if ($formula -match '\{.*FIXED.*\}|\{.*INCLUDE.*\}|\{.*EXCLUDE.*\}') {
                                    $analysis.LODCalculations++
                                }

                                # Table calculations
                                if ($formula -match 'RUNNING_|WINDOW_|RANK|INDEX|FIRST|LAST') {
                                    $analysis.TableCalculations++
                                }

                                # Calculation complexity (count nested brackets/references)
                                $nesting = ([regex]::Matches($formula, '\[')).Count
                                if ($nesting -gt $analysis.MaxCalculationNesting) {
                                    $analysis.MaxCalculationNesting = $nesting
                                }
                            }
                        }
                    }
                }

                # Check for aliases
                if ($ds.aliases -or $ds.column.aliases) {
                    $analysis.HasAliases = $true
                }
            }
        }
    }

    # Analyze worksheets
    $worksheets = $xml.workbook.worksheets.worksheet
    if ($worksheets) {
        $analysis.Worksheets = @($worksheets).Count

        foreach ($ws in $worksheets) {
            # Filters
            $filters = $ws.filter
            if ($filters) {
                foreach ($filter in $filters) {
                    $analysis.TotalFilters++

                    $filterClass = $filter.class
                    switch ($filterClass) {
                        'categorical' { $analysis.CategoricalFilters++ }
                        'quantitative' { $analysis.QuantitativeFilters++ }
                        'relative-date-filter' { $analysis.DateFilters++ }
                    }

                    # Top N filters
                    if ($filter.groupfilter.function -contains 'top') {
                        $analysis.TopNFilters++
                    }

                    # Context filters
                    if ($filter.GetAttribute('context') -eq 'true') {
                        $analysis.ContextFilters++
                    }
                }
            }

            # Mark types
            $mark = $ws.table.panes.pane.mark
            if ($mark) {
                $markType = $mark.class
                if ($markType -and $markType -notin $analysis.MarkTypes) {
                    $analysis.MarkTypes += $markType
                }
            }

            # Color encodings
            $colorEncodings = $ws.SelectNodes("//encodings/color")
            if ($colorEncodings) {
                $analysis.ColorEncodings += $colorEncodings.Count
            }

            # Tooltips
            $tooltip = $ws.SelectNodes("//format[@attr='tooltip']")
            if ($tooltip -and $tooltip.value) {
                $analysis.HasTooltips = $true
            }
        }
    }

    # Analyze dashboards
    $dashboards = $xml.workbook.dashboards.dashboard
    if ($dashboards) {
        $analysis.Dashboards = @($dashboards).Count

        foreach ($dashboard in $dashboards) {
            # Count zones
            $zones = $dashboard.zones.zone
            if ($zones) {
                $analysis.DashboardZones += @($zones).Count

                foreach ($zone in $zones) {
                    $zoneType = $zone.type
                    switch ($zoneType) {
                        'parameter-control' { $analysis.ParameterControls++ }
                        'filter' { $analysis.FilterControls++ }
                        'text' { $analysis.TextBoxes++ }
                        'bitmap' { $analysis.Images++ }
                        'web' { $analysis.WebObjects++ }
                    }
                }
            }

            # Actions
            $actions = $dashboard.actions.action
            if ($actions) {
                foreach ($action in $actions) {
                    $analysis.TotalActions++

                    # Determine action type
                    if ($action.target -and $action.target.GetAttribute('type') -eq 'worksheet') {
                        if ($action.SelectSingleNode("activation/command[@command='tsl:tableselect']")) {
                            $analysis.FilterActions++
                        }
                    }

                    if ($action.'url-action') {
                        $analysis.URLActions++
                    }

                    if ($action.'parameter-action') {
                        $analysis.ParameterActions++
                    }

                    # Highlight actions (harder to detect, often implicit)
                    $actionName = $action.name
                    if ($actionName -match 'highlight' -or $action.caption -match 'highlight') {
                        $analysis.HighlightActions++
                    }
                }
            }
        }
    }

    # Analyze stories
    $stories = $xml.workbook.stories.story
    if ($stories) {
        $analysis.Stories = @($stories).Count
    }

    # Calculate complexity score (0-100)
    $complexityScore = 0
    $complexityScore += [Math]::Min($analysis.Worksheets * 2, 20)
    $complexityScore += [Math]::Min($analysis.Dashboards * 5, 20)
    $complexityScore += [Math]::Min($analysis.CalculatedFields * 1, 15)
    $complexityScore += [Math]::Min($analysis.LODCalculations * 3, 15)
    $complexityScore += [Math]::Min($analysis.TotalFilters * 1, 10)
    $complexityScore += [Math]::Min($analysis.TotalActions * 2, 10)
    $complexityScore += [Math]::Min($analysis.DataSourcesWithData * 5, 10)

    $analysis | Add-Member -MemberType NoteProperty -Name ComplexityScore -Value $complexityScore

    # Add warnings
    if ($analysis.CalculatedFields -eq 0 -and $analysis.Worksheets -gt 0) {
        $analysis.Warnings += "No calculated fields (very simple workbook)"
    }
    if ($analysis.MaxCalculationNesting -gt 10) {
        $analysis.Warnings += "Complex nested calculations detected"
    }
    if ($analysis.LODCalculations -gt 5) {
        $analysis.Warnings += "Many LOD calculations (advanced)"
    }
    if ($analysis.DataSourcesWithData -gt 3) {
        $analysis.Warnings += "Multiple data sources (potential complexity)"
    }
    if ($analysis.ContextFilters -gt 0) {
        $analysis.Warnings += "Context filters in use (advanced filtering)"
    }

    return $analysis
}

# Function to display analysis results
function Display-Analysis {
    param(
        [object]$Analysis,
        [switch]$Detailed
    )

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "FILE: $($Analysis.FileName)" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Path: $($Analysis.FilePath)" -ForegroundColor Gray
    Write-Host "Size: $([Math]::Round($Analysis.FileSize / 1KB, 2)) KB" -ForegroundColor Gray
    Write-Host "Tableau Version: $($Analysis.Version) ($($Analysis.SourceBuild))" -ForegroundColor Gray

    Write-Host "`nCOMPLEXITY SCORE: $($Analysis.ComplexityScore)/100" -ForegroundColor $(
        if ($Analysis.ComplexityScore -lt 30) { 'Green' }
        elseif ($Analysis.ComplexityScore -lt 60) { 'Yellow' }
        else { 'Red' }
    )

    Write-Host "`nOVERVIEW:" -ForegroundColor Yellow
    Write-Host "  Worksheets: $($Analysis.Worksheets)" -ForegroundColor White
    Write-Host "  Dashboards: $($Analysis.Dashboards)" -ForegroundColor White
    Write-Host "  Stories: $($Analysis.Stories)" -ForegroundColor White
    Write-Host "  Data Sources: $($Analysis.DataSourcesWithData) (+ $($Analysis.Parameters) parameters)" -ForegroundColor White

    if ($Analysis.ConnectionTypes.Count -gt 0) {
        Write-Host "  Connection Types: $($Analysis.ConnectionTypes -join ', ')" -ForegroundColor Gray
    }

    Write-Host "`nFIELDS:" -ForegroundColor Yellow
    Write-Host "  Total Fields: $($Analysis.TotalFields)" -ForegroundColor White
    Write-Host "  Dimensions: $($Analysis.Dimensions)" -ForegroundColor White
    Write-Host "  Measures: $($Analysis.Measures)" -ForegroundColor White
    Write-Host "  Calculated Fields: $($Analysis.CalculatedFields)" -ForegroundColor $(if ($Analysis.CalculatedFields -gt 0) { 'Cyan' } else { 'Gray' })

    if ($Analysis.LODCalculations -gt 0) {
        Write-Host "  LOD Calculations: $($Analysis.LODCalculations)" -ForegroundColor Magenta
    }
    if ($Analysis.TableCalculations -gt 0) {
        Write-Host "  Table Calculations: $($Analysis.TableCalculations)" -ForegroundColor Magenta
    }

    Write-Host "`nFILTERS:" -ForegroundColor Yellow
    Write-Host "  Total Filters: $($Analysis.TotalFilters)" -ForegroundColor White
    if ($Detailed -and $Analysis.TotalFilters -gt 0) {
        Write-Host "    Categorical: $($Analysis.CategoricalFilters)" -ForegroundColor Gray
        Write-Host "    Quantitative: $($Analysis.QuantitativeFilters)" -ForegroundColor Gray
        Write-Host "    Date: $($Analysis.DateFilters)" -ForegroundColor Gray
        Write-Host "    Top N: $($Analysis.TopNFilters)" -ForegroundColor Gray
        if ($Analysis.ContextFilters -gt 0) {
            Write-Host "    Context Filters: $($Analysis.ContextFilters)" -ForegroundColor Magenta
        }
    }

    if ($Analysis.TotalActions -gt 0) {
        Write-Host "`nACTIONS:" -ForegroundColor Yellow
        Write-Host "  Total Actions: $($Analysis.TotalActions)" -ForegroundColor White
        if ($Detailed) {
            Write-Host "    Filter Actions: $($Analysis.FilterActions)" -ForegroundColor Gray
            Write-Host "    Highlight Actions: $($Analysis.HighlightActions)" -ForegroundColor Gray
            Write-Host "    URL Actions: $($Analysis.URLActions)" -ForegroundColor Gray
            Write-Host "    Parameter Actions: $($Analysis.ParameterActions)" -ForegroundColor Gray
        }
    }

    if ($Analysis.Dashboards -gt 0) {
        Write-Host "`nDASHBOARD ELEMENTS:" -ForegroundColor Yellow
        Write-Host "  Total Zones: $($Analysis.DashboardZones)" -ForegroundColor White
        Write-Host "  Parameter Controls: $($Analysis.ParameterControls)" -ForegroundColor White
        Write-Host "  Filter Controls: $($Analysis.FilterControls)" -ForegroundColor White
        if ($Detailed) {
            Write-Host "  Text Boxes: $($Analysis.TextBoxes)" -ForegroundColor Gray
            Write-Host "  Images: $($Analysis.Images)" -ForegroundColor Gray
            Write-Host "  Web Objects: $($Analysis.WebObjects)" -ForegroundColor Gray
        }
    }

    if ($Analysis.MarkTypes.Count -gt 0) {
        Write-Host "`nVISUAL TYPES:" -ForegroundColor Yellow
        Write-Host "  $($Analysis.MarkTypes -join ', ')" -ForegroundColor White
    }

    if ($Detailed) {
        Write-Host "`nADVANCED FEATURES:" -ForegroundColor Yellow
        Write-Host "  Custom SQL: $(if ($Analysis.HasCustomSQL) { 'Yes' } else { 'No' })" -ForegroundColor $(if ($Analysis.HasCustomSQL) { 'Cyan' } else { 'Gray' })
        Write-Host "  Aliases: $(if ($Analysis.HasAliases) { 'Yes' } else { 'No' })" -ForegroundColor Gray
        Write-Host "  Tooltips: $(if ($Analysis.HasTooltips) { 'Yes' } else { 'No' })" -ForegroundColor Gray
        Write-Host "  Max Calc Nesting: $($Analysis.MaxCalculationNesting)" -ForegroundColor $(if ($Analysis.MaxCalculationNesting -gt 5) { 'Yellow' } else { 'Gray' })
    }

    if ($Analysis.Warnings.Count -gt 0) {
        Write-Host "`nWARNINGS/NOTES:" -ForegroundColor Yellow
        foreach ($warning in $Analysis.Warnings) {
            Write-Host "  ! $warning" -ForegroundColor Yellow
        }
    }
}

# Main execution
Write-Host "`nTableau Workbook Analyzer" -ForegroundColor Cyan
Write-Host "========================`n" -ForegroundColor Cyan

$results = @()

# Get files to analyze
if (Test-Path $Path -PathType Container) {
    # Directory
    if ($Recursive) {
        $files = Get-ChildItem -Path $Path -Include "*.twb", "*.twbx" -Recurse
    }
    else {
        $files = Get-ChildItem -Path $Path -Include "*.twb", "*.twbx"
    }
}
elseif (Test-Path $Path -PathType Leaf) {
    # Single file
    $files = @(Get-Item $Path)
}
else {
    Write-Error "Path not found: $Path"
    exit 1
}

if ($files.Count -eq 0) {
    Write-Warning "No TWB or TWBX files found at: $Path"
    exit 0
}

Write-Host "Found $($files.Count) file(s) to analyze`n" -ForegroundColor Green

foreach ($file in $files) {
    $twbPath = $file.FullName
    $originalPath = $file.FullName
    $tempDir = $null

    # If TWBX, extract it first
    if ($file.Extension -eq '.twbx') {
        Write-Host "Extracting TWBX: $($file.Name)..." -ForegroundColor Gray
        $twbPath, $tempDir = Extract-TwbFromTwbx -TwbxPath $file.FullName

        if (-not $twbPath) {
            continue
        }
    }

    # Analyze the TWB file
    $analysis = Analyze-TwbFile -FilePath $twbPath -OriginalPath $originalPath

    if ($analysis) {
        $results += $analysis
        Display-Analysis -Analysis $analysis -Detailed:$Detailed
    }

    # Cleanup temp directory if we extracted a TWBX
    if ($tempDir -and (Test-Path $tempDir)) {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Summary
if ($results.Count -gt 1) {
    Write-Host "`n========================================"  -ForegroundColor Cyan
    Write-Host "SUMMARY" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Total Files Analyzed: $($results.Count)" -ForegroundColor White
    Write-Host "Average Complexity: $([Math]::Round(($results | Measure-Object -Property ComplexityScore -Average).Average, 1))" -ForegroundColor White
    Write-Host "Total Worksheets: $(($results | Measure-Object -Property Worksheets -Sum).Sum)" -ForegroundColor White
    Write-Host "Total Dashboards: $(($results | Measure-Object -Property Dashboards -Sum).Sum)" -ForegroundColor White
    Write-Host "Total Calculated Fields: $(($results | Measure-Object -Property CalculatedFields -Sum).Sum)" -ForegroundColor White

    Write-Host "`nComplexity Distribution:" -ForegroundColor Yellow
    $simple = ($results | Where-Object { $_.ComplexityScore -lt 30 }).Count
    $medium = ($results | Where-Object { $_.ComplexityScore -ge 30 -and $_.ComplexityScore -lt 60 }).Count
    $complex = ($results | Where-Object { $_.ComplexityScore -ge 60 }).Count

    Write-Host "  Simple (< 30): $simple files" -ForegroundColor Green
    Write-Host "  Medium (30-59): $medium files" -ForegroundColor Yellow
    Write-Host "  Complex (60+): $complex files" -ForegroundColor Red
}

# Export to CSV if requested
if ($OutputCsv) {
    $results | Select-Object FileName, FilePath, FileSize, Version, ComplexityScore, Worksheets, Dashboards, Stories, `
        DataSourcesWithData, Parameters, TotalFields, CalculatedFields, LODCalculations, TableCalculations, `
        TotalFilters, TotalActions, DashboardZones, ParameterControls, FilterControls | `
        Export-Csv -Path $OutputCsv -NoTypeInformation

    Write-Host "`nResults exported to: $OutputCsv" -ForegroundColor Green
}

Write-Host ""
