# TimePlanner - CI/CD Friendly Build Script
param(
    [string]$Version = "",
    [switch]$CI = $false,
    [switch]$Test = $false
)

Write-Host "====================================" -ForegroundColor Cyan
Write-Host "   TimePlanner - Build Executable" -ForegroundColor Cyan  
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

# Function to check command availability
function Test-Command {
    param([string]$Command)
    try {
        Get-Command $Command -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# Generate version info if not provided
if ([string]::IsNullOrEmpty($Version)) {
    $dateString = Get-Date -Format "yyyy.MM.dd"
    $timeString = Get-Date -Format "HHmm"
    $Version = "$dateString.$timeString"
    Write-Host "Generated version: $Version" -ForegroundColor Green
} else {
    Write-Host "Using provided version: $Version" -ForegroundColor Green
}

# Create/update version.py file
Write-Host "Creating version.py file..." -ForegroundColor Yellow
$versionContent = @"
"""
Version information for TimePlanner
This file is automatically updated by the build system.
"""

__version__ = "$Version"
VERSION = "$Version"
BUILD_DATE = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
GIT_COMMIT = ""

def get_version():
    """Get the current version string"""
    return __version__

def get_version_info():
    """Get detailed version information"""
    return {
        'version': __version__,
        'build_date': BUILD_DATE,
        'git_commit': GIT_COMMIT
    }

def get_version_string():
    """Get a formatted version string for display"""
    if BUILD_DATE:
        return f"TimePlanner v{__version__} (Built: {BUILD_DATE})"
    else:
        return f"TimePlanner v{__version__}"
"@

$versionContent | Out-File -FilePath "version.py" -Encoding UTF8

# Check Python
if (-not (Test-Command "python")) {
    Write-Host "ERROR: Python is not installed or not in PATH." -ForegroundColor Red
    if (-not $CI) {
        Read-Host "Press Enter to exit"
    }
    exit 1
}

Write-Host "Python version:" -ForegroundColor Green
python --version

# Install/upgrade required packages
Write-Host ""
Write-Host "Installing/updating Python packages..." -ForegroundColor Yellow
$packages = @(
    "PyQt5",
    "pandas", 
    "openpyxl", 
    "python-docx", 
    "pyinstaller", 
    "pillow"
)

foreach ($package in $packages) {
    Write-Host "Installing $package..." -ForegroundColor Yellow
    python -m pip install --upgrade $package
    if ($LASTEXITCODE -ne 0) {
        Write-Host "WARNING: Failed to install $package" -ForegroundColor Yellow
    }
}

Write-Host ""

# Clean previous builds
Write-Host "Cleaning previous builds..." -ForegroundColor Yellow
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }

# Create application icon if it doesn't exist
if (-not (Test-Path "icon.ico")) {
    Write-Host "Creating application icon..." -ForegroundColor Yellow
    $iconScript = @'
try:
    from PIL import Image, ImageDraw
    import os

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
    python create_icon.py
    Remove-Item "create_icon.py" -Force -ErrorAction SilentlyContinue
}

Write-Host ""

# Build executable name
$exeName = if ($Test) { "TimePlanner-test" } else { "TimePlanner-$Version" }

# Build the executable
Write-Host "Building executable: $exeName.exe" -ForegroundColor Yellow
Write-Host "This may take several minutes..." -ForegroundColor Yellow

# PyInstaller arguments
$pyinstallerArgs = @(
    "--onefile",
    "--windowed",
    "--name=$exeName",
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
}

# Add main.py
$pyinstallerArgs += "main.py"

# Run PyInstaller
Write-Host "PyInstaller command: pyinstaller $($pyinstallerArgs -join ' ')" -ForegroundColor Cyan
& pyinstaller $pyinstallerArgs

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: PyInstaller build failed." -ForegroundColor Red
    if (-not $CI) {
        Read-Host "Press Enter to exit"
    }
    exit 1
}

Write-Host ""

# Check if executable was created
$exePath = "dist\$exeName.exe"
if (Test-Path $exePath) {
    $exeSize = (Get-Item $exePath).Length / 1MB
    Write-Host "====================================" -ForegroundColor Green
    Write-Host "   Build completed successfully!" -ForegroundColor Green
    Write-Host "====================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Executable created: $exePath" -ForegroundColor Green
    Write-Host "Size: $([math]::Round($exeSize, 1)) MB" -ForegroundColor Cyan
    Write-Host "Version: $Version" -ForegroundColor Cyan
    
    if (-not $Test) {
        # Create portable distribution
        Write-Host ""
        Write-Host "Creating portable distribution..." -ForegroundColor Yellow
        $portableDir = "TimePlanner-Portable-$Version"
        if (Test-Path $portableDir) { Remove-Item -Recurse -Force $portableDir }
        New-Item -ItemType Directory -Path $portableDir | Out-Null
        
        # Copy executable (rename to simple name)
        Copy-Item $exePath "$portableDir\TimePlanner.exe"
        
        # Copy other files if they exist
        if (Test-Path "README.md") { Copy-Item "README.md" $portableDir }
        if (Test-Path "settings.json") { Copy-Item "settings.json" $portableDir }
        if (Test-Path "lang") { Copy-Item -Recurse "lang" $portableDir }
        
        # Create version info file
        $versionInfo = @"
TimePlanner Portable v$Version

This is a portable version of TimePlanner.
No installation required - just run TimePlanner.exe

Built on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
"@
        $versionInfo | Out-File -FilePath "$portableDir\VERSION.txt" -Encoding UTF8
        
        Write-Host "Portable distribution created: $portableDir" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "You can now:" -ForegroundColor Cyan
    Write-Host "1. Run: .\$exePath" -ForegroundColor White
    Write-Host "2. Copy the executable anywhere and run it" -ForegroundColor White
    
    if (-not $CI -and -not $Test) {
        # Test the executable
        $testRun = Read-Host "Do you want to test the executable now? (y/n)"
        if ($testRun -eq "y" -or $testRun -eq "Y") {
            Write-Host "Starting $exeName.exe..." -ForegroundColor Green
            Start-Process $exePath
        }
    }
} else {
    Write-Host "ERROR: $exeName.exe was not created." -ForegroundColor Red
    Write-Host "Check the build output above for errors." -ForegroundColor Yellow
    if (-not $CI) {
        Read-Host "Press Enter to exit"
    }
    exit 1
}

if (-not $CI) {
    Write-Host ""
    Read-Host "Press Enter to finish"
}

Write-Host ""
Write-Host "Build completed successfully!" -ForegroundColor Green