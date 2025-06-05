# Targeted Fix for PMC Terminal Syntax Errors
# This script applies specific fixes for the identified syntax errors

Write-Host "PMC Terminal - Targeted Syntax Fix" -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor DarkGray

# Fix 1: core-data.ps1 line 605 - Sort-Object syntax error
$coreDataPath = "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-data.ps1"

if (Test-Path $coreDataPath) {
    Write-Host "`nFixing core-data.ps1..." -ForegroundColor Yellow
    
    # Create backup
    $backupPath = $coreDataPath + ".backup." + (Get-Date -Format "yyyyMMdd_HHmmss")
    Copy-Item $coreDataPath $backupPath
    Write-Host "Backup created: $(Split-Path $backupPath -Leaf)" -ForegroundColor Green
    
    # Read content
    $content = Get-Content $coreDataPath -Raw
    
    # Fix the Sort-Object expression on line ~605
    # The issue is in the Search-CommandSnippets function
    # Original problematic pattern:
    # Sort-Object UseCount -Descending, @{Expression = {if($_.LastUsed){try{[datetime]$_.LastUsed}catch{[datetime]::MinValue}}else{[datetime]::MinValue}}; Descending = $true}, Description
    
    # The fix: properly format the Sort-Object with hashtables
    $problemPattern = 'Sort-Object UseCount -Descending, @\{Expression = \{if\(\$_\.LastUsed\)\{try\{\[datetime\]\$_\.LastUsed\}catch\{\[datetime\]::MinValue\}\}else\{\[datetime\]::MinValue\}\}; Descending = \$true\}, Description'
    $fixedPattern = 'Sort-Object @{Expression = {$_.UseCount}; Descending = $true}, @{Expression = {if($_.LastUsed){try{[datetime]$_.LastUsed}catch{[datetime]::MinValue}}else{[datetime]::MinValue}}; Descending = $true}, Description'
    
    $content = $content -replace $problemPattern, $fixedPattern
    
    # Alternative fix if the above doesn't match exactly
    if ($content -match 'Sort-Object UseCount -Descending,') {
        Write-Host "Applying alternative Sort-Object fix..." -ForegroundColor DarkCyan
        # More flexible pattern matching
        $content = $content -replace 'Sort-Object\s+UseCount\s+-Descending,\s*@\{Expression\s*=\s*\{[^}]+\}[^}]*\},\s*Description', 
                                    'Sort-Object @{Expression = {$_.UseCount}; Descending = $true}, @{Expression = {if($_.LastUsed){try{[datetime]$_.LastUsed}catch{[datetime]::MinValue}}else{[datetime]::MinValue}}; Descending = $true}, Description'
    }
    
    # Save the fixed content
    $content | Set-Content $coreDataPath -Encoding UTF8
    Write-Host "core-data.ps1 updated" -ForegroundColor Green
}

# Fix 2: Variable reference errors with $[Math]::
$filesToFixMath = @(
    "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\helper.ps1",
    "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-time.ps1",
    "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\ui.ps1"
)

foreach ($filePath in $filesToFixMath) {
    if (Test-Path $filePath) {
        $fileName = Split-Path $filePath -Leaf
        Write-Host "`nFixing $fileName..." -ForegroundColor Yellow
        
        # Create backup
        $backupPath = $filePath + ".backup." + (Get-Date -Format "yyyyMMdd_HHmmss")
        Copy-Item $filePath $backupPath
        Write-Host "Backup created: $(Split-Path $backupPath -Leaf)" -ForegroundColor Green
        
        # Read and fix content
        $content = Get-Content $filePath -Raw
        $originalContent = $content
        
        # Fix $[Math]:: to [Math]::
        $content = $content -replace '\$\[Math\]::', '[Math]::'
        
        # Fix $([Math]:: to ([Math]::
        $content = $content -replace '\$\(\[Math\]::', '([Math]::'
        
        if ($content -ne $originalContent) {
            $content | Set-Content $filePath -Encoding UTF8
            Write-Host "$fileName updated" -ForegroundColor Green
        } else {
            Write-Host "$fileName - no changes needed" -ForegroundColor DarkGray
        }
    }
}

Write-Host "`n" + ("=" * 60) -ForegroundColor DarkGray
Write-Host "Fixes applied!" -ForegroundColor Cyan

# Validate the fixes
Write-Host "`nValidating syntax..." -ForegroundColor Yellow

$allFiles = @(
    "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-data.ps1",
    "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\helper.ps1",
    "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-time.ps1",
    "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\main.ps1"
)

$hasErrors = $false
foreach ($filePath in $allFiles) {
    if (Test-Path $filePath) {
        $fileName = Split-Path $filePath -Leaf
        Write-Host "`nChecking $fileName..." -ForegroundColor DarkCyan
        
        $content = Get-Content $filePath -Raw
        $tokens = $null
        $errors = $null
        $null = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$errors)
        
        if ($errors.Count -eq 0) {
            Write-Host "✓ No syntax errors" -ForegroundColor Green
        } else {
            Write-Host "✗ Found $($errors.Count) error(s)" -ForegroundColor Red
            $hasErrors = $true
            foreach ($error in $errors) {
                Write-Host "  Line $($error.Token.StartLine): $($error.Message)" -ForegroundColor Red
            }
        }
    }
}

if (-not $hasErrors) {
    Write-Host "`n✓ All syntax errors have been fixed!" -ForegroundColor Green
    Write-Host "`nYou can now run: . .\test-setup.ps1" -ForegroundColor Yellow
} else {
    Write-Host "`n⚠ Some syntax errors remain. Please run .\diagnose-errors.ps1 for details." -ForegroundColor Yellow
}
