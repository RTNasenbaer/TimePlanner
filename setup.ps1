# TimePlanner - PowerShell Setup Script
Write-Host "====================================" -ForegroundColor Cyan
Write-Host "   TimePlanner - Automated Setup" -ForegroundColor Cyan  
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

# Check if conda is available
try {
    $condaVersion = conda --version 2>$null
    Write-Host "Found: $condaVersion" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Conda is not installed or not in PATH." -ForegroundColor Red
    Write-Host "Please install Miniconda or Anaconda first:" -ForegroundColor Yellow
    Write-Host "https://docs.conda.io/en/latest/miniconda.html" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# Remove existing environment
Write-Host "Removing existing 'zeitplan' environment if it exists..." -ForegroundColor Yellow
conda env remove -n zeitplan -y 2>$null
Write-Host ""

# Create environment
Write-Host "Creating conda environment 'zeitplan' with Python 3.11..." -ForegroundColor Yellow
conda create -n zeitplan python=3.11 -y
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to create conda environment." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host "Environment created successfully!" -ForegroundColor Green
Write-Host ""

# Install conda packages
Write-Host "Installing conda packages (pyqt, pyqtgraph, pandas, openpyxl)..." -ForegroundColor Yellow
conda install -n zeitplan pyqt pyqtgraph pandas openpyxl -y
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to install conda packages." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host "Conda packages installed successfully!" -ForegroundColor Green
Write-Host ""

# Install pip packages
Write-Host "Installing pip packages (PyQt5, qtawesome)..." -ForegroundColor Yellow
conda run -n zeitplan pip install PyQt5 qtawesome
if ($LASTEXITCODE -ne 0) {
    Write-Host "WARNING: Failed to install some pip packages. Trying requirements.txt..." -ForegroundColor Yellow
    conda run -n zeitplan pip install -r requirements.txt
}
Write-Host "Pip packages installed!" -ForegroundColor Green
Write-Host ""

Write-Host "====================================" -ForegroundColor Green
Write-Host "   Setup completed successfully!" -ForegroundColor Green
Write-Host "====================================" -ForegroundColor Green
Write-Host ""
Write-Host "To run TimePlanner:" -ForegroundColor Cyan
Write-Host "Option 1 - PowerShell: .\run.ps1" -ForegroundColor White
Write-Host "Option 2 - Manual: conda activate zeitplan && python main.py" -ForegroundColor White
Write-Host ""
Read-Host "Press Enter to finish"