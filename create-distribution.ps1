# TimePlanner - Create Distribution Package
Write-Host "====================================" -ForegroundColor Cyan
Write-Host "  TimePlanner - Create Distribution" -ForegroundColor Cyan  
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

# Check if executable exists
if (-not (Test-Path "dist\TimePlanner.exe")) {
    Write-Host "ERROR: TimePlanner.exe not found in dist folder." -ForegroundColor Red
    Write-Host "Please run build-exe.ps1 first to create the executable." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Creating distribution package..." -ForegroundColor Yellow

# Create distribution folder
$distName = "TimePlanner-Portable"
if (Test-Path $distName) { 
    Remove-Item -Recurse -Force $distName 
}
New-Item -ItemType Directory -Name $distName | Out-Null

# Copy executable
Copy-Item "dist\TimePlanner.exe" "$distName\"

# Copy data files if they exist
if (Test-Path "*.xlsx") { Copy-Item "*.xlsx" "$distName\" }
if (Test-Path "*.png") { Copy-Item "*.png" "$distName\" }

# Create README
$readmeContent = @"
# TimePlanner - Portable Application

## How to Use
1. Double-click TimePlanner.exe to start the application
2. No installation required - this is a portable application
3. You can copy this folder anywhere and run it

## System Requirements
- Windows 10 or later
- No additional software installation required

## Troubleshooting
- If the application doesn't start, try running as administrator
- Make sure Windows Defender or antivirus isn't blocking the executable
- If you get a "Windows protected your PC" message, click "More info" then "Run anyway"

## Version Information
Built on: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Executable size: $([math]::Round((Get-Item "dist\TimePlanner.exe").Length / 1MB, 1)) MB

For support or updates, contact the developer.
"@

$readmeContent | Out-File -FilePath "$distName\README.txt" -Encoding UTF8

# Create ZIP package
Write-Host "Creating ZIP package..." -ForegroundColor Yellow
$zipName = "$distName.zip"
if (Test-Path $zipName) { Remove-Item $zipName }

# Use PowerShell's Compress-Archive
Compress-Archive -Path $distName -DestinationPath $zipName

Write-Host ""
Write-Host "====================================" -ForegroundColor Green
Write-Host "   Distribution package created!" -ForegroundColor Green
Write-Host "====================================" -ForegroundColor Green
Write-Host ""
Write-Host "Created files:" -ForegroundColor Cyan
Write-Host "- $distName\ (folder with portable app)" -ForegroundColor White
Write-Host "- $zipName (ZIP package for distribution)" -ForegroundColor White
Write-Host ""
Write-Host "You can now:" -ForegroundColor Cyan
Write-Host "1. Share the ZIP file with users" -ForegroundColor White
Write-Host "2. Extract and run anywhere on Windows" -ForegroundColor White
Write-Host "3. No installation required!" -ForegroundColor White

Write-Host ""
$openFolder = Read-Host "Open distribution folder? (y/n)"
if ($openFolder -eq "y" -or $openFolder -eq "Y") {
    Start-Process "explorer.exe" -ArgumentList (Get-Location).Path
}

Write-Host ""
Read-Host "Press Enter to finish"