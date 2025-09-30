# TimePlanner - Build Executable Script
Write-Host "====================================" -ForegroundColor Cyan
Write-Host "   TimePlanner - Build Executable" -ForegroundColor Cyan  
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

# Check if conda is available
try {
    $condaVersion = conda --version 2>$null
    Write-Host "Found: $condaVersion" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Conda is not installed or not in PATH." -ForegroundColor Red
    Write-Host "Please install Miniconda or Anaconda first." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if zeitplan environment exists
Write-Host "Checking for zeitplan environment..." -ForegroundColor Yellow
try {
    $envCheck = conda info --envs 2>$null | Select-String "zeitplan"
    if (-not $envCheck) {
        Write-Host "Environment 'zeitplan' not found. Creating it first..." -ForegroundColor Yellow
        Write-Host "Running setup..." -ForegroundColor Yellow
        .\setup.ps1
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: Failed to create environment." -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
    } else {
        Write-Host "Found zeitplan environment!" -ForegroundColor Green
    }
} catch {
    Write-Host "ERROR: Could not check conda environments." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# Update PyInstaller if needed
Write-Host "Installing/updating PyInstaller..." -ForegroundColor Yellow
conda run -n zeitplan pip install pyinstaller pillow
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to install PyInstaller." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# Clean previous builds
Write-Host "Cleaning previous builds..." -ForegroundColor Yellow
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "*.spec") { Remove-Item -Force "*.spec" }

Write-Host ""

# Create application icon (simple text-based icon)
Write-Host "Creating application resources..." -ForegroundColor Yellow
if (-not (Test-Path "icon.ico")) {
    Write-Host "No icon.ico found. Creating a simple one..." -ForegroundColor Yellow
    
    # Create a simple Python script to generate the icon
    $iconScript = @'
from PIL import Image, ImageDraw
import os

try:
    # Create a simple icon
    img = Image.new('RGBA', (256, 256), (70, 130, 180, 255))  # Steel blue background
    draw = ImageDraw.Draw(img)

    # Draw a clock-like design
    center = (128, 128)
    radius = 100

    # Outer circle
    draw.ellipse([center[0]-radius, center[1]-radius, center[0]+radius, center[1]+radius], 
                 outline=(255, 255, 255, 255), width=8)

    # Clock hands
    draw.line([center[0], center[1], center[0], center[1]-60], fill=(255, 255, 255, 255), width=8)  # Hour hand
    draw.line([center[0], center[1], center[0]+45, center[1]-30], fill=(255, 255, 255, 255), width=6)  # Minute hand

    # Center dot
    draw.ellipse([center[0]-8, center[1]-8, center[0]+8, center[1]+8], fill=(255, 255, 255, 255))

    # Save as ICO
    img.save('icon.ico', format='ICO', sizes=[(256,256), (128,128), (64,64), (32,32), (16,16)])
    print('Successfully created icon.ico')
except Exception as e:
    print(f'Error creating icon: {e}')
    print('Continuing without custom icon...')
'@
    
    $iconScript | Out-File -FilePath "create_icon.py" -Encoding UTF8
    conda run -n zeitplan python create_icon.py
    Remove-Item "create_icon.py" -Force
    
    # Verify icon was created
    if (-not (Test-Path "icon.ico")) {
        Write-Host "Icon creation failed. Building without custom icon..." -ForegroundColor Yellow
    }
}

Write-Host ""

# Build the executable
Write-Host "Building executable with PyInstaller..." -ForegroundColor Yellow
Write-Host "This may take several minutes..." -ForegroundColor Yellow

# Build PyInstaller arguments
$pyinstallerArgs = @(
    "--onefile",
    "--windowed",
    "--name=TimePlanner",
    "--hidden-import=pandas",
    "--hidden-import=openpyxl",
    "--hidden-import=PyQt5.QtWidgets",
    "--hidden-import=PyQt5.QtGui",
    "--hidden-import=PyQt5.QtCore",
    "--hidden-import=docx",
    "--hidden-import=docx.shared",
    "--hidden-import=colorsys",
    "--hidden-import=json",
    "--hidden-import=datetime",
    "--collect-all=PyQt5",
    "--clean"
)

# Add icon if it exists
if (Test-Path "icon.ico") {
    $pyinstallerArgs += "--icon=icon.ico"
    Write-Host "Using custom icon: icon.ico" -ForegroundColor Green
} else {
    Write-Host "No custom icon - using default" -ForegroundColor Yellow
}

# Add data files if they exist
if (Test-Path "*.xlsx") { 
    $pyinstallerArgs += "--add-data=*.xlsx;."
    Write-Host "Including Excel files" -ForegroundColor Green
}
if (Test-Path "*.png") { 
    $pyinstallerArgs += "--add-data=*.png;."
    Write-Host "Including PNG files" -ForegroundColor Green
}
if (Test-Path "*.docx") { 
    $pyinstallerArgs += "--add-data=*.docx;."
    Write-Host "Including DOCX template files" -ForegroundColor Green
}
if (Test-Path "lang") { 
    $pyinstallerArgs += "--add-data=lang;lang"
    Write-Host "Including language files" -ForegroundColor Green
} else {
    Write-Host "WARNING: lang folder not found - translations may not work" -ForegroundColor Yellow
}

# Add main.py
$pyinstallerArgs += "main.py"

Write-Host "PyInstaller arguments: $($pyinstallerArgs -join ' ')" -ForegroundColor Cyan
conda run -n zeitplan pyinstaller $pyinstallerArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: PyInstaller build failed." -ForegroundColor Red
    Write-Host "Check the output above for details." -ForegroundColor Yellow
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
    Write-Host "You can now:" -ForegroundColor Cyan
    Write-Host "1. Run: .\dist\TimePlanner.exe" -ForegroundColor White
    Write-Host "2. Copy dist\TimePlanner.exe anywhere and run it" -ForegroundColor White
    Write-Host "3. Create a desktop shortcut to TimePlanner.exe" -ForegroundColor White
    Write-Host ""
    
    # Test the executable
    $testRun = Read-Host "Do you want to test the executable now? (y/n)"
    if ($testRun -eq "y" -or $testRun -eq "Y") {
        Write-Host "Starting TimePlanner.exe..." -ForegroundColor Green
        Start-Process "dist\TimePlanner.exe"
    }
} else {
    Write-Host "ERROR: TimePlanner.exe was not created." -ForegroundColor Red
    Write-Host "Check the build output above for errors." -ForegroundColor Yellow
}

Write-Host ""
Read-Host "Press Enter to finish"