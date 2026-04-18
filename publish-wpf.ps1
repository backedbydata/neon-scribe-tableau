# PowerShell script to publish the WPF application as a standalone executable

Write-Host "Publishing Neon Scribe for Tableau..." -ForegroundColor Cyan
Write-Host ""

# Kill any running instances
Write-Host "Checking for running instances..." -ForegroundColor Yellow
$process = Get-Process -Name "NeonScribe.Tableau.WPF" -ErrorAction SilentlyContinue
if ($process) {
    Write-Host "Found running instance. Stopping..." -ForegroundColor Yellow
    Stop-Process -Name "NeonScribe.Tableau.WPF" -Force
    Start-Sleep -Seconds 2
}

# Clean previous builds
Write-Host "Cleaning previous builds..." -ForegroundColor Yellow
if (Test-Path "publish") {
    Remove-Item -Path "publish" -Recurse -Force
}

# Publish as self-contained executable
Write-Host "Publishing application (this may take a minute)..." -ForegroundColor Yellow
dotnet publish src/NeonScribe.Tableau.WPF/NeonScribe.Tableau.WPF.csproj `
    --configuration Release `
    --runtime win-x64 `
    --self-contained true `
    --output publish `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "✓ Build successful!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Executable location:" -ForegroundColor Cyan
    Write-Host "  $PWD\publish\NeonScribe.Tableau.WPF.exe" -ForegroundColor White
    Write-Host ""
    Write-Host "To run the application:" -ForegroundColor Cyan
    Write-Host "  .\publish\NeonScribe.Tableau.WPF.exe" -ForegroundColor White
    Write-Host ""

    # Get file size
    $exePath = Join-Path $PWD "publish\NeonScribe.Tableau.WPF.exe"
    if (Test-Path $exePath) {
        $fileSize = (Get-Item $exePath).Length / 1MB
        Write-Host "Executable size: $([math]::Round($fileSize, 2)) MB" -ForegroundColor Gray
    }
} else {
    Write-Host ""
    Write-Host "✗ Build failed!" -ForegroundColor Red
    Write-Host "Please check the error messages above." -ForegroundColor Red
    exit 1
}
