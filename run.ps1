# TimePlanner - PowerShell Run Script
Write-Host "====================================" -ForegroundColor Cyan
Write-Host "      Starting TimePlanner" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

# Check if conda is available
try {
    conda --version 2>$null | Out-Null
} catch {
    Write-Host "ERROR: Conda is not available." -ForegroundColor Red
    Write-Host "Please make sure Anaconda/Miniconda is installed." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if zeitplan environment exists
Write-Host "Checking for zeitplan environment..." -ForegroundColor Yellow
try {
    $envCheck = conda info --envs 2>$null | Select-String "zeitplan"
    if (-not $envCheck) {
        throw "Environment not found"
    }
    Write-Host "Found zeitplan environment!" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Environment 'zeitplan' not found." -ForegroundColor Red
    Write-Host "Available environments:" -ForegroundColor Yellow
    conda env list
    Write-Host ""
    Write-Host "Please run setup.ps1 first to create the environment." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Starting TimePlanner in zeitplan environment..." -ForegroundColor Green
Write-Host ""

# Run the application
conda run -n zeitplan python main.py

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "ERROR: Application failed to start." -ForegroundColor Red
    Write-Host "Make sure main.py exists and dependencies are installed." -ForegroundColor Yellow
}

Write-Host ""
Read-Host "Press Enter to close"