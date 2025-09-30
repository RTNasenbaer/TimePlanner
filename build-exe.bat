@echo off
echo ====================================
echo   TimePlanner - Build Executable
echo ====================================
echo.

REM Check conda
where conda >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: Conda not found. Please install Anaconda/Miniconda first.
    pause
    exit /b 1
)

echo Found conda, building executable...
echo.

REM Run the PowerShell script
powershell -ExecutionPolicy Bypass -File "build-exe.ps1"

echo.
pause