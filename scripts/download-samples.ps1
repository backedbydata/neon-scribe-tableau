# PowerShell script to download sample Tableau files from GitHub repositories

param(
    [string]$OutputDir = "..\samples",
    [switch]$Verbose
)

# Create output directory if it doesn't exist
$samplesPath = Join-Path $PSScriptRoot $OutputDir
if (-not (Test-Path $samplesPath)) {
    New-Item -ItemType Directory -Path $samplesPath -Force | Out-Null
    Write-Host "Created samples directory: $samplesPath" -ForegroundColor Green
}

# GitHub repositories with Tableau samples
$repos = @(
    @{
        Name = "vishnu-t-r/tableau_twbx"
        Description = "Simple chart examples"
        Url = "https://github.com/vishnu-t-r/tableau_twbx"
        Files = @("*.twbx", "*.twb")
    },
    @{
        Name = "amol-modi/Tableau"
        Description = "Academic examples with stories and dashboards"
        Url = "https://github.com/amol-modi/Tableau"
        Files = @("*.twb", "*.twbx")
    },
    @{
        Name = "aloth/tableau-book-resources"
        Description = "Professional examples from Visual Analytics book"
        Url = "https://github.com/aloth/tableau-book-resources"
        Files = @("*.twbx")
    },
    @{
        Name = "PacktPublishing/Tableau-10-Best-Practices"
        Description = "Best practices examples"
        Url = "https://github.com/PacktPublishing/Tableau-10-Best-Practices"
        Files = @("*.twb", "*.twbx", "*.xls", "*.xlsx")
    }
)

function Download-GitHubRepo {
    param(
        [string]$RepoUrl,
        [string]$RepoName,
        [string]$Description,
        [string[]]$FilePatterns
    )

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Repository: $RepoName" -ForegroundColor Cyan
    Write-Host "Description: $Description" -ForegroundColor Gray
    Write-Host "========================================" -ForegroundColor Cyan

    # Create subdirectory for this repo
    $repoDir = Join-Path $samplesPath ($RepoName -replace '/', '_')
    if (-not (Test-Path $repoDir)) {
        New-Item -ItemType Directory -Path $repoDir -Force | Out-Null
    }

    # Convert GitHub URL to API URL
    $apiUrl = $RepoUrl -replace 'github.com', 'api.github.com/repos'
    $zipUrl = "$RepoUrl/archive/refs/heads/main.zip"
    $masterZipUrl = "$RepoUrl/archive/refs/heads/master.zip"

    # Try to download the repository as zip
    $tempZip = Join-Path $env:TEMP "$($RepoName -replace '/', '_').zip"
    $downloaded = $false

    try {
        Write-Host "Attempting to download from main branch..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $zipUrl -OutFile $tempZip -ErrorAction Stop
        $downloaded = $true
        Write-Host "Downloaded successfully from main branch" -ForegroundColor Green
    }
    catch {
        try {
            Write-Host "Trying master branch..." -ForegroundColor Yellow
            Invoke-WebRequest -Uri $masterZipUrl -OutFile $tempZip -ErrorAction Stop
            $downloaded = $true
            Write-Host "Downloaded successfully from master branch" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to download repository: $_" -ForegroundColor Red
            return
        }
    }

    if ($downloaded) {
        # Extract zip file
        $tempExtract = Join-Path $env:TEMP ($RepoName -replace '/', '_')
        if (Test-Path $tempExtract) {
            Remove-Item $tempExtract -Recurse -Force
        }

        try {
            Expand-Archive -Path $tempZip -DestinationPath $tempExtract -Force
            Write-Host "Extracted repository" -ForegroundColor Green

            # Find and copy TWB/TWBX files
            $foundFiles = @()
            foreach ($pattern in $FilePatterns) {
                $files = Get-ChildItem -Path $tempExtract -Filter $pattern -Recurse -ErrorAction SilentlyContinue
                $foundFiles += $files
            }

            if ($foundFiles.Count -gt 0) {
                Write-Host "Found $($foundFiles.Count) matching files" -ForegroundColor Green

                foreach ($file in $foundFiles) {
                    $destPath = Join-Path $repoDir $file.Name
                    Copy-Item -Path $file.FullName -Destination $destPath -Force

                    if ($Verbose) {
                        Write-Host "  - Copied: $($file.Name)" -ForegroundColor Gray
                    }
                }

                Write-Host "Copied $($foundFiles.Count) files to $repoDir" -ForegroundColor Green
            }
            else {
                Write-Host "No matching files found in repository" -ForegroundColor Yellow
            }

            # Cleanup
            Remove-Item $tempExtract -Recurse -Force
            Remove-Item $tempZip -Force
        }
        catch {
            Write-Host "Error extracting or copying files: $_" -ForegroundColor Red
        }
    }
}

# Download from each repository
Write-Host "`nTableau Sample Files Downloader" -ForegroundColor Cyan
Write-Host "================================`n" -ForegroundColor Cyan

foreach ($repo in $repos) {
    Download-GitHubRepo -RepoUrl $repo.Url -RepoName $repo.Name -Description $repo.Description -FilePatterns $repo.Files
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Download Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Files saved to: $samplesPath" -ForegroundColor White

# Summary
$totalFiles = (Get-ChildItem -Path $samplesPath -Recurse -Include "*.twb", "*.twbx" -ErrorAction SilentlyContinue).Count
Write-Host "`nTotal TWB/TWBX files downloaded: $totalFiles" -ForegroundColor Green

# List directories created
Write-Host "`nDirectories created:" -ForegroundColor Yellow
Get-ChildItem -Path $samplesPath -Directory | ForEach-Object {
    $fileCount = (Get-ChildItem -Path $_.FullName -Include "*.twb", "*.twbx" -ErrorAction SilentlyContinue).Count
    Write-Host "  - $($_.Name): $fileCount files" -ForegroundColor Gray
}

Write-Host "`nYou can now analyze these files using the TWB analyzer script." -ForegroundColor Cyan
