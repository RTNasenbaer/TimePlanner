# TimePlanner - Simple PowerShell Run Script
Write-Host "Starting TimePlanner..." -ForegroundColor Cyan

# Try to run directly - conda run will fail if environment doesn't exist
conda run -n zeitplan python main.py

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "Failed to start TimePlanner." -ForegroundColor Red
    Write-Host ""
    Write-Host "Possible issues:" -ForegroundColor Yellow
    Write-Host "1. Environment 'zeitplan' doesn't exist - run: .\setup.ps1" -ForegroundColor White
    Write-Host "2. main.py is missing or has errors" -ForegroundColor White
    Write-Host "3. Dependencies not installed properly" -ForegroundColor White
    Write-Host ""
    Write-Host "Available conda environments:" -ForegroundColor Yellow
    conda env list
}

Write-Host ""
Read-Host "Press Enter to close"