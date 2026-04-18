# PowerShell script to build and run the WPF application

Write-Host "Building Neon Scribe for Tableau..." -ForegroundColor Cyan
Write-Host ""

# Kill any running instances
Write-Host "Checking for running instances..." -ForegroundColor Yellow
$process = Get-Process -Name "NeonScribe.Tableau.WPF" -ErrorAction SilentlyContinue
if ($process) {
    Write-Host "Found running instance. Stopping..." -ForegroundColor Yellow
    Stop-Process -Name "NeonScribe.Tableau.WPF" -Force
    Start-Sleep -Seconds 2
}

# Build the solution
Write-Host "Building solution..." -ForegroundColor Yellow
dotnet build NeonScribe.Tableau.slnx --configuration Debug

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "[SUCCESS] Build successful!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Starting application..." -ForegroundColor Cyan

    # Run the application
    $exePath = "src\NeonScribe.Tableau.WPF\bin\Debug\net10.0-windows\NeonScribe.Tableau.WPF.exe"
    if (Test-Path $exePath) {
        Start-Process $exePath
        Write-Host "Application started!" -ForegroundColor Green
    } else {
        Write-Host "[ERROR] Executable not found at: $exePath" -ForegroundColor Red
    }
} else {
    Write-Host ""
    Write-Host "[ERROR] Build failed!" -ForegroundColor Red
    Write-Host "Please check the error messages above." -ForegroundColor Red
    exit 1
}
