# Fix Duplicate Content in core-data.ps1
# The file appears to have been duplicated multiple times

Write-Host "PMC Terminal - Fix Duplicate Content Issue" -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor DarkGray

$coreDataPath = "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-data.ps1"

if (-not (Test-Path $coreDataPath)) {
    Write-Host "Error: core-data.ps1 not found!" -ForegroundColor Red
    exit 1
}

# Create backup
$backupPath = $coreDataPath + ".backup.duplicated." + (Get-Date -Format "yyyyMMdd_HHmmss")
Copy-Item $coreDataPath $backupPath
Write-Host "Backup created: $(Split-Path $backupPath -Leaf)" -ForegroundColor Green

# Read the file
$content = Get-Content $coreDataPath -Raw
$lines = Get-Content $coreDataPath
Write-Host "Original file size: $($lines.Count) lines" -ForegroundColor Yellow

# Look for duplicate module headers
$moduleHeaders = @()
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match '^# Core Data Management Module$') {
        $moduleHeaders += $i + 1  # Line number (1-based)
    }
}

Write-Host "Found $($moduleHeaders.Count) module headers at lines: $($moduleHeaders -join ', ')" -ForegroundColor Yellow

if ($moduleHeaders.Count -gt 1) {
    Write-Host "`nFile appears to be duplicated. Extracting first complete copy..." -ForegroundColor Yellow
    
    # Find the end of the first module (before the second header)
    $firstHeaderLine = $moduleHeaders[0] - 1  # Convert to 0-based
    $secondHeaderLine = $moduleHeaders[1] - 1  # Convert to 0-based
    
    # Extract the first complete module
    $firstModule = $lines[0..($secondHeaderLine - 1)]
    
    # Check if the extracted module looks complete
    $functionCount = ($firstModule | Where-Object { $_ -match '^function\s+global:' }).Count
    Write-Host "First module contains $functionCount function definitions" -ForegroundColor Green
    
    # Look for key functions we expect
    $expectedFunctions = @(
        "Get-DefaultSettings",
        "Show-ProjectsAndTemplates", 
        "Add-Project",
        "Add-TodoTask",
        "Search-CommandSnippets"
    )
    
    $foundFunctions = @()
    foreach ($func in $expectedFunctions) {
        if ($firstModule -join "`n" -match "function\s+global:$func") {
            $foundFunctions += $func
        }
    }
    
    Write-Host "Found $($foundFunctions.Count) of $($expectedFunctions.Count) expected functions" -ForegroundColor Green
    
    if ($foundFunctions.Count -eq $expectedFunctions.Count) {
        Write-Host "`nFirst module appears complete. Extracting..." -ForegroundColor Green
        
        # Fix the Sort-Object issue in line 605 before saving
        Write-Host "`nFixing Sort-Object syntax at line 605..." -ForegroundColor DarkCyan
        
        # Find and fix the incomplete Sort-Object
        for ($i = 0; $i -lt $firstModule.Count; $i++) {
            if ($firstModule[$i] -match 'Sort-Object @\{Expression = \{$' -and $i -gt 0) {
                if ($firstModule[$i-1] -match 'Get-CommandSnippet.*\|$') {
                    Write-Host "Found incomplete Sort-Object at line $($i+1)" -ForegroundColor Yellow
                    $firstModule[$i] = '                Sort-Object @{Expression = {$_.UseCount}; Descending = $true}, @{Expression = {if($_.LastUsed){try{[datetime]$_.LastUsed}catch{[datetime]::MinValue}}else{[datetime]::MinValue}}; Descending = $true}, Description'
                }
            }
            
            # Fix Sort-Object UseCount patterns
            if ($firstModule[$i] -match 'Sort-Object\s+UseCount\s+-Descending,') {
                $firstModule[$i] = $firstModule[$i] -replace 'Sort-Object\s+UseCount\s+-Descending,\s+Description',
                                                             'Sort-Object @{Expression = {$_.UseCount}; Descending = $true}, Description'
            }
        }
        
        # Save the deduplicated and fixed content
        $fixedPath = $coreDataPath + ".fixed"
        $firstModule -join "`n" | Set-Content $fixedPath -Encoding UTF8
        
        # Validate the fixed file
        Write-Host "`nValidating fixed file..." -ForegroundColor Yellow
        $fixedContent = Get-Content $fixedPath -Raw
        $tokens = $null
        $errors = $null
        $null = [System.Management.Automation.PSParser]::Tokenize($fixedContent, [ref]$errors)
        
        if ($errors.Count -eq 0) {
            Write-Host "✓ Fixed file has no syntax errors!" -ForegroundColor Green
            
            # Replace the original with the fixed version
            Move-Item $fixedPath $coreDataPath -Force
            Write-Host "✓ Original file replaced with deduplicated version" -ForegroundColor Green
            Write-Host "`nReduced from $($lines.Count) lines to $($firstModule.Count) lines" -ForegroundColor Green
        } else {
            Write-Host "✗ Fixed file still has $($errors.Count) error(s)" -ForegroundColor Red
            Write-Host "Errors:" -ForegroundColor Red
            $errors | Select-Object -First 5 | ForEach-Object {
                Write-Host "  Line $($_.Token.StartLine): $($_.Message)" -ForegroundColor Red
            }
            
            Write-Host "`nFixed file saved as: $fixedPath" -ForegroundColor Yellow
            Write-Host "Review and manually copy if needed" -ForegroundColor Yellow
        }
    } else {
        Write-Host "`nFirst module appears incomplete. Missing functions:" -ForegroundColor Red
        $expectedFunctions | Where-Object { $_ -notin $foundFunctions } | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Red
        }
    }
} else {
    Write-Host "`nNo duplication detected, but file still has errors." -ForegroundColor Yellow
    Write-Host "Running standard fixes..." -ForegroundColor Yellow
    
    # Run the comprehensive fix script
    & ".\fix-core-data-comprehensive.ps1"
}

Write-Host "`n" + ("=" * 60) -ForegroundColor DarkGray
Write-Host "Process completed!" -ForegroundColor Cyan
Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "1. Run: . .\test-setup.ps1" -ForegroundColor White
Write-Host "2. Test the commands that were failing" -ForegroundColor White
Write-Host "3. If errors persist, restore from backup: Copy-Item '$backupPath' '$coreDataPath'" -ForegroundColor White
