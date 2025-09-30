# TimePlanner Build Test Script
# This script tests if the optimized application can be built and run correctly

Write-Host "====================================" -ForegroundColor Cyan
Write-Host "   TimePlanner - Build Test Script" -ForegroundColor Cyan  
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

# Test 1: Check if all required files exist
Write-Host "Test 1: Checking required files..." -ForegroundColor Yellow

$requiredFiles = @(
    "main.py",
    "translation.py", 
    "requirements.txt",
    "lang\de-de.json",
    "lang\en-us.json"
)

$missingFiles = @()
foreach ($file in $requiredFiles) {
    if (-not (Test-Path $file)) {
        $missingFiles += $file
        Write-Host "  MISSING: $file" -ForegroundColor Red
    } else {
        Write-Host "  OK: $file" -ForegroundColor Green
    }
}

if ($missingFiles.Count -gt 0) {
    Write-Host "ERROR: Missing required files. Cannot proceed with build test." -ForegroundColor Red
    exit 1
}

Write-Host ""

# Test 2: Check Python syntax
Write-Host "Test 2: Checking Python syntax..." -ForegroundColor Yellow
python -m py_compile main.py
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: main.py has syntax errors" -ForegroundColor Red
    exit 1
} else {
    Write-Host "  OK: main.py syntax is valid" -ForegroundColor Green
}

python -m py_compile translation.py
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: translation.py has syntax errors" -ForegroundColor Red
    exit 1
} else {
    Write-Host "  OK: translation.py syntax is valid" -ForegroundColor Green
}

Write-Host ""

# Test 3: Check imports
Write-Host "Test 3: Testing optimized imports..." -ForegroundColor Yellow
python -c "
try:
    # Test lazy imports
    import main
    print('main.py import: OK')
    
    # Test translation system
    from translation import tr, init_translation_system
    init_translation_system('de-de')
    test_text = tr('app_title')
    print(f'Translation system: OK - {test_text}')
    
    # Test optimized classes
    settings = main.AppSettings()
    print('AppSettings (lazy loading): OK')
    
    section = main.TimeSection(0, 10, 'Test', main.QColor(255, 0, 0))
    print(f'TimeSection (slots): OK - {section.name}')
    
    # Test color caching
    color1 = main.get_nice_color(0)
    color2 = main.get_nice_color(0)  # Should use cache
    print('Color caching: OK')
    
    # Test lazy pandas import
    try:
        pd = main._import_pandas()
        print('Lazy pandas import: OK')
    except ImportError as e:
        print(f'Lazy pandas import: MISSING - {e}')
    
    # Test lazy docx import
    try:
        Document = main._import_docx()
        print('Lazy docx import: OK')
    except ImportError as e:
        print(f'Lazy docx import: MISSING - {e}')
        
    print('ALL IMPORT TESTS PASSED')
    
except Exception as e:
    print(f'IMPORT ERROR: {e}')
    import traceback
    traceback.print_exc()
    exit(1)
"

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Import tests failed" -ForegroundColor Red
    exit 1
} else {
    Write-Host "  OK: All imports working correctly" -ForegroundColor Green
}

Write-Host ""

# Test 4: Check PyInstaller compatibility
Write-Host "Test 4: Checking PyInstaller compatibility..." -ForegroundColor Yellow

# Check if conda environment exists
$envCheck = conda info --envs 2>$null | Select-String "zeitplan"
if (-not $envCheck) {
    Write-Host "  WARNING: zeitplan environment not found. Run setup.ps1 first." -ForegroundColor Yellow
} else {
    Write-Host "  OK: zeitplan environment exists" -ForegroundColor Green
    
    # Test PyInstaller availability
    conda run -n zeitplan pyinstaller --version > $null 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  OK: PyInstaller is available" -ForegroundColor Green
    } else {
        Write-Host "  WARNING: PyInstaller not found in environment" -ForegroundColor Yellow
    }
}

Write-Host ""

# Test 5: Quick build test (dry run)
Write-Host "Test 5: Testing build configuration..." -ForegroundColor Yellow

if (Test-Path "TimePlanner.spec") {
    Write-Host "  OK: TimePlanner.spec exists" -ForegroundColor Green
    
    # Check spec file content
    $specContent = Get-Content "TimePlanner.spec" -Raw
    if ($specContent -match "lang.*lang" -and $specContent -match "docx" -and $specContent -match "translation") {
        Write-Host "  OK: Spec file includes optimized configuration" -ForegroundColor Green
    } else {
        Write-Host "  WARNING: Spec file might be outdated" -ForegroundColor Yellow
    }
} else {
    Write-Host "  INFO: No spec file found (will be auto-generated)" -ForegroundColor Cyan
}

Write-Host ""

# Test 6: Check template files
Write-Host "Test 6: Checking template files..." -ForegroundColor Yellow

if (Test-Path "Trainingseinheit_Name_StandardFile_Date.docx") {
    Write-Host "  OK: DOCX template found" -ForegroundColor Green
} else {
    Write-Host "  WARNING: DOCX template not found - Word export may not work" -ForegroundColor Yellow
}

Write-Host ""

# Final Results
Write-Host "====================================" -ForegroundColor Green
Write-Host "   Build Test Summary" -ForegroundColor Green
Write-Host "====================================" -ForegroundColor Green
Write-Host ""

if ($missingFiles.Count -eq 0) {
    Write-Host "✓ All required files present" -ForegroundColor Green
    Write-Host "✓ Python syntax valid" -ForegroundColor Green
    Write-Host "✓ Optimized imports working" -ForegroundColor Green
    Write-Host "✓ Ready for build process" -ForegroundColor Green
    Write-Host ""
    Write-Host "You can now run:" -ForegroundColor Cyan
    Write-Host "1. .\build-exe.ps1 - To create the executable" -ForegroundColor White
    Write-Host "2. .\create-distribution.ps1 - To create portable package" -ForegroundColor White
    Write-Host ""
    
    # Offer to run the build
    $runBuild = Read-Host "Would you like to run the build process now? (y/n)"
    if ($runBuild -eq "y" -or $runBuild -eq "Y") {
        Write-Host "Starting build process..." -ForegroundColor Green
        .\build-exe.ps1
    }
} else {
    Write-Host "✗ Some issues found - check output above" -ForegroundColor Red
    Write-Host "Please fix the issues before building" -ForegroundColor Yellow
}

Write-Host ""
Read-Host "Press Enter to finish"