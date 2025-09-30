# TimePlanner - All-in-One Build Script
Write-Host "====================================" -ForegroundColor Cyan
Write-Host "   TimePlanner - Complete Build" -ForegroundColor Cyan  
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "This script will:" -ForegroundColor Yellow
Write-Host "1. Set up the conda environment (if needed)" -ForegroundColor White
Write-Host "2. Build the standalone executable" -ForegroundColor White
Write-Host "3. Create a distribution package" -ForegroundColor White
Write-Host "4. Test the final application" -ForegroundColor White
Write-Host ""

$continue = Read-Host "Continue? (y/n)"
if ($continue -ne "y" -and $continue -ne "Y") {
    Write-Host "Build cancelled." -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "Step 1: Building executable..." -ForegroundColor Cyan
.\build-exe.ps1

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Build failed. Check the output above." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "Step 2: Creating distribution package..." -ForegroundColor Cyan
.\create-distribution.ps1

Write-Host ""
Write-Host "====================================" -ForegroundColor Green
Write-Host "   Complete build finished!" -ForegroundColor Green
Write-Host "====================================" -ForegroundColor Green
Write-Host ""
Write-Host "Your TimePlanner application is ready!" -ForegroundColor Green
Write-Host ""
Write-Host "Files created:" -ForegroundColor Cyan
Write-Host "✓ dist\TimePlanner.exe (standalone executable)" -ForegroundColor Green
Write-Host "✓ TimePlanner-Portable.zip (distribution package)" -ForegroundColor Green
Write-Host ""
Write-Host "What you can do now:" -ForegroundColor Cyan
Write-Host "• Double-click TimePlanner.exe to run locally" -ForegroundColor White
Write-Host "• Share TimePlanner-Portable.zip with others" -ForegroundColor White
Write-Host "• Recipients just extract and run - no install needed!" -ForegroundColor White

Write-Host ""
Read-Host "Press Enter to finish"