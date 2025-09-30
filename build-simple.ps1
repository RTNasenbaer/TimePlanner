# TimePlanner - Simple Build Script (No Icon)
Write-Host "====================================" -ForegroundColor Cyan
Write-Host "   TimePlanner - Simple Build" -ForegroundColor Cyan  
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

# Check if conda is available
try {
    $condaVersion = conda --version 2>$null
    Write-Host "Found: $condaVersion" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Conda is not installed or not in PATH." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if zeitplan environment exists
Write-Host "Checking for zeitplan environment..." -ForegroundColor Yellow
try {
    $envCheck = conda info --envs 2>$null | Select-String "zeitplan"
    if (-not $envCheck) {
        Write-Host "Environment 'zeitplan' not found. Please run setup.ps1 first." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    Write-Host "Found zeitplan environment!" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Could not check conda environments." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# Install PyInstaller
Write-Host "Installing PyInstaller..." -ForegroundColor Yellow
conda run -n zeitplan pip install pyinstaller
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to install PyInstaller." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Clean previous builds
Write-Host "Cleaning previous builds..." -ForegroundColor Yellow
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "*.spec") { Remove-Item -Force "*.spec" }

Write-Host ""

# Build the executable - simple version
Write-Host "Building executable with PyInstaller (simple)..." -ForegroundColor Yellow
Write-Host "This may take several minutes..." -ForegroundColor Yellow

conda run -n zeitplan pyinstaller --onefile --windowed --name=TimePlanner --clean main.py

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: PyInstaller build failed." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "====================================" -ForegroundColor Green
Write-Host "   Build completed successfully!" -ForegroundColor Green
Write-Host "====================================" -ForegroundColor Green
Write-Host ""

# Check if executable was created
if (Test-Path "dist\TimePlanner.exe") {
    $exeSize = (Get-Item "dist\TimePlanner.exe").Length / 1MB
    Write-Host "Executable created: dist\TimePlanner.exe" -ForegroundColor Green
    Write-Host "Size: $([math]::Round($exeSize, 1)) MB" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Test the executable:" -ForegroundColor Cyan
    Write-Host ".\dist\TimePlanner.exe" -ForegroundColor White
} else {
    Write-Host "ERROR: TimePlanner.exe was not created." -ForegroundColor Red
}

Write-Host ""
Read-Host "Press Enter to finish"