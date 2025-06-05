# Test Script for PowerShell Productivity Suite
# This script tests if the basic setup works after our fixes

Write-Host "Testing PowerShell Productivity Suite Setup..." -ForegroundColor Cyan
Write-Host "=" * 50 -ForegroundColor Gray

$moduleRoot = Split-Path $MyInvocation.MyCommand.Path -Parent

# Test 1: Check if files exist
Write-Host "`n1. Checking required files..." -ForegroundColor Yellow
$requiredFiles = @(
    "helper.ps1",
    "core-data.ps1", 
    "main.ps1",
    "stubs.ps1"
)

$allFilesExist = $true
foreach ($file in $requiredFiles) {
    $path = Join-Path $moduleRoot $file
    if (Test-Path $path) {
        Write-Host "  ✅ $file" -ForegroundColor Green
    } else {
        Write-Host "  ❌ $file" -ForegroundColor Red
        $allFilesExist = $false
    }
}

if (-not $allFilesExist) {
    Write-Host "`n❌ Some required files are missing. Cannot continue." -ForegroundColor Red
    exit 1
}

# Test 2: Try loading stubs
Write-Host "`n2. Testing stubs.ps1..." -ForegroundColor Yellow
try {
    . "$moduleRoot\stubs.ps1"
    Write-Host "  ✅ Stubs loaded successfully" -ForegroundColor Green
} catch {
    Write-Host "  ❌ Error loading stubs: $_" -ForegroundColor Red
    exit 1
}

# Test 3: Try loading helper.ps1
Write-Host "`n3. Testing helper.ps1..." -ForegroundColor Yellow
try {
    . "$moduleRoot\helper.ps1"
    Write-Host "  ✅ Helper loaded successfully" -ForegroundColor Green
} catch {
    Write-Host "  ❌ Error loading helper: $_" -ForegroundColor Red
    exit 1
}

# Test 4: Try loading core-data.ps1
Write-Host "`n4. Testing core-data.ps1..." -ForegroundColor Yellow
try {
    . "$moduleRoot\core-data.ps1"
    Write-Host "  ✅ Core-data loaded successfully" -ForegroundColor Green
} catch {
    Write-Host "  ❌ Error loading core-data: $_" -ForegroundColor Red
    Write-Host "      This might be the parser error. Try running fix-parser-error.ps1 first." -ForegroundColor Yellow
    exit 1
}

# Test 5: Test key functions
Write-Host "`n5. Testing key functions..." -ForegroundColor Yellow

$testFunctions = @(
    "Get-DefaultSettings",
    "Write-Header", 
    "Write-Success",
    "Show-Dashboard",
    "Initialize-ThemeSystem"
)

foreach ($func in $testFunctions) {
    try {
        if (Get-Command $func -ErrorAction Stop) {
            Write-Host "  ✅ $func" -ForegroundColor Green
        }
    } catch {
        Write-Host "  ❌ $func not found" -ForegroundColor Red
    }
}

# Test 6: Test data initialization
Write-Host "`n6. Testing data initialization..." -ForegroundColor Yellow
try {
    if ($script:Data -and $script:Data.Settings) {
        Write-Host "  ✅ Data structure initialized" -ForegroundColor Green
        Write-Host "    - Projects: $($script:Data.Projects.Count)" -ForegroundColor Gray
        Write-Host "    - Tasks: $($script:Data.Tasks.Count)" -ForegroundColor Gray
        Write-Host "    - Settings loaded: $($script:Data.Settings.Keys.Count) keys" -ForegroundColor Gray
    } else {
        Write-Host "  ❌ Data structure not properly initialized" -ForegroundColor Red
    }
} catch {
    Write-Host "  ❌ Error checking data structure: $_" -ForegroundColor Red
}

Write-Host "`n" + "=" * 50 -ForegroundColor Gray
Write-Host "✅ Basic setup test completed!" -ForegroundColor Green
Write-Host ""
Write-Host "If all tests passed, you can now run:" -ForegroundColor Cyan
Write-Host "  .\main.ps1" -ForegroundColor White
Write-Host ""
Write-Host "If core-data.ps1 failed, run this first:" -ForegroundColor Yellow
Write-Host "  .\fix-parser-error.ps1" -ForegroundColor White
