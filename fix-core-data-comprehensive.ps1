# Comprehensive Fix for PMC Terminal core-data.ps1 Corruption
# This script fixes the multiple syntax errors found in the diagnostic

Write-Host "PMC Terminal - Comprehensive Syntax Fix" -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor DarkGray

$coreDataPath = "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-data.ps1"

if (-not (Test-Path $coreDataPath)) {
    Write-Host "Error: core-data.ps1 not found!" -ForegroundColor Red
    exit 1
}

# Create backup
$backupPath = $coreDataPath + ".backup." + (Get-Date -Format "yyyyMMdd_HHmmss")
Copy-Item $coreDataPath $backupPath
Write-Host "Backup created: $(Split-Path $backupPath -Leaf)" -ForegroundColor Green

# Read all lines
$lines = Get-Content $coreDataPath
Write-Host "Total lines in file: $($lines.Count)" -ForegroundColor Yellow

# Fix 1: Remove the incomplete line 605 and any duplicate header
Write-Host "`nFix 1: Removing incomplete Sort-Object at line 605..." -ForegroundColor DarkCyan

# Check what's actually at line 605 (array index 604)
if ($lines.Count -gt 604) {
    Write-Host "Line 605 content: $($lines[604])" -ForegroundColor DarkGray
    
    # If line 605 is the incomplete Sort-Object, we need to fix it
    if ($lines[604] -match 'Sort-Object @\{Expression = \{$') {
        Write-Host "Found incomplete Sort-Object statement" -ForegroundColor Yellow
        
        # Check if this is part of Search-CommandSnippets function
        # Look for the previous line that starts the pipeline
        if ($lines[603] -match 'Get-CommandSnippet.*\|$') {
            Write-Host "This appears to be in Search-CommandSnippets function" -ForegroundColor DarkCyan
            
            # Replace the incomplete line with the correct Sort-Object
            $lines[604] = '                Sort-Object @{Expression = {$_.UseCount}; Descending = $true}, @{Expression = {if($_.LastUsed){try{[datetime]$_.LastUsed}catch{[datetime]::MinValue}}else{[datetime]::MinValue}}; Descending = $true}, Description'
            
            # Remove the duplicate header comment if it exists on the next line
            if ($lines[605] -match '^# Core Data Management Module') {
                Write-Host "Removing duplicate header at line 606" -ForegroundColor Yellow
                $lines = $lines[0..604] + $lines[606..($lines.Count-1)]
            }
        }
    }
}

# Fix 2: Fix all Sort-Object expressions with missing/incorrect syntax
Write-Host "`nFix 2: Fixing Sort-Object expressions..." -ForegroundColor DarkCyan
$fixCount = 0

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]
    
    # Fix Sort-Object UseCount -Descending patterns
    if ($line -match 'Sort-Object\s+UseCount\s+-Descending,\s+@\{Expression') {
        $lines[$i] = $line -replace 'Sort-Object\s+UseCount\s+-Descending,\s+@\{Expression\s*=\s*\{[^}]+\}[^}]*\},\s*Description',
                                    'Sort-Object @{Expression = {$_.UseCount}; Descending = $true}, @{Expression = {if($_.LastUsed){try{[datetime]$_.LastUsed}catch{[datetime]::MinValue}}else{[datetime]::MinValue}}; Descending = $true}, Description'
        $fixCount++
    }
    
    # Fix simple Sort-Object UseCount -Descending, Description
    if ($line -match 'Sort-Object\s+UseCount\s+-Descending,\s+Description') {
        $lines[$i] = $line -replace 'Sort-Object\s+UseCount\s+-Descending,\s+Description',
                                    'Sort-Object @{Expression = {$_.UseCount}; Descending = $true}, Description'
        $fixCount++
    }
    
    # Fix Sort-Object with nested hashtables (Priority sorting)
    if ($line -match 'Sort-Object\s+@\{Expression=\{@\{"Critical"=1;"High"=2;"Medium"=3;"Low"=4\}\[\$_\.(Name|Priority)\]\}\}') {
        $lines[$i] = $line -replace 'Sort-Object\s+@\{Expression=\{(@\{"Critical"=1;"High"=2;"Medium"=3;"Low"=4\})\[\$_\.(Name|Priority)\]\}\}',
                                    'Sort-Object @{Expression={$1[$_.$2]}}' 
        $fixCount++
    }
}

Write-Host "Fixed $fixCount Sort-Object expressions" -ForegroundColor Green

# Fix 3: Fix variable reference errors with colons
Write-Host "`nFix 3: Fixing variable reference errors..." -ForegroundColor DarkCyan
$fixCount = 0

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]
    
    # Fix patterns like Write-Host "  $categoryName: stuff"
    # The issue is the colon after the variable in string interpolation
    if ($line -match 'Write-Host\s+"[^"]*\$\w+:\s*[^"]*"') {
        # This pattern is actually valid in PowerShell, but might be causing issues
        # Let's wrap the variable in $() to be safe
        $lines[$i] = $line -replace '\$(\w+):', '$($1):'
        $fixCount++
    }
}

Write-Host "Fixed $fixCount variable reference patterns" -ForegroundColor Green

# Fix 4: Specific fixes for known problematic patterns
Write-Host "`nFix 4: Applying specific pattern fixes..." -ForegroundColor DarkCyan

# Fix the specific pattern at line 1430
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match 'Write-Host\s+"\s+\$categoryName:\s+\$\(\$catGroup\.Count\)\s+snippet\(s\)"') {
        $lines[$i] = '        Write-Host "  $($categoryName): $($catGroup.Count) snippet(s)"'
        Write-Host "Fixed line $($i+1): categoryName display" -ForegroundColor Green
    }
}

# Save the fixed content
Write-Host "`nSaving fixed content..." -ForegroundColor Yellow
$lines -join "`n" | Set-Content $coreDataPath -Encoding UTF8

# Validate the fix
Write-Host "`nValidating syntax..." -ForegroundColor Yellow
$content = Get-Content $coreDataPath -Raw
$tokens = $null
$errors = $null
$null = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$errors)

if ($errors.Count -eq 0) {
    Write-Host "✓ All syntax errors have been fixed!" -ForegroundColor Green
    Write-Host "`nYou can now run: . .\test-setup.ps1" -ForegroundColor Yellow
} else {
    Write-Host "✗ $($errors.Count) error(s) remain:" -ForegroundColor Red
    
    # Show first 5 errors for debugging
    $errors | Select-Object -First 5 | ForEach-Object {
        Write-Host "  Line $($_.Token.StartLine): $($_.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nThe file may be too corrupted. Consider:" -ForegroundColor Yellow
    Write-Host "1. Restoring from backup: Copy-Item '$backupPath' '$coreDataPath'" -ForegroundColor White
    Write-Host "2. Or check if you have a clean version in Git" -ForegroundColor White
}

# Also fix the other files with Sort-Object issues
Write-Host "`n`nFixing other files..." -ForegroundColor Cyan

# Fix core-time.ps1
$coreTimePath = "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-time.ps1"
if (Test-Path $coreTimePath) {
    Write-Host "`nFixing core-time.ps1..." -ForegroundColor Yellow
    $content = Get-Content $coreTimePath -Raw
    
    # Fix the Sort-Object pattern with nested hashtable
    $content = $content -replace 'Sort-Object\s+@\{Expression=\{@\{"Critical"=1;"High"=2;"Medium"=3;"Low"=4\}\[\$_\.Priority\]\}\}',
                                'Sort-Object @{Expression={$priorityOrder[$_.Priority]}}'
    
    # Add priority order definition if needed
    if ($content -notmatch '\$priorityOrder\s*=') {
        $insertPoint = $content.IndexOf('foreach ($taskItem in $group.Group | Sort-Object')
        if ($insertPoint -gt 0) {
            $lineStart = $content.LastIndexOf("`n", $insertPoint) + 1
            $indent = " " * 12  # Adjust based on context
            $priorityDef = "$indent`$priorityOrder = @{`"Critical`"=1;`"High`"=2;`"Medium`"=3;`"Low`"=4}`n$indent"
            $content = $content.Insert($lineStart, $priorityDef)
        }
    }
    
    $content | Set-Content $coreTimePath -Encoding UTF8
    Write-Host "core-time.ps1 fixed" -ForegroundColor Green
}

# Fix main.ps1
$mainPath = "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\main.ps1"
if (Test-Path $mainPath) {
    Write-Host "`nFixing main.ps1..." -ForegroundColor Yellow
    $lines = Get-Content $mainPath
    
    for ($i = 0; $i -lt $lines.Count; $i++) {
        # Fix Sort-Object patterns with date expressions
        if ($lines[$i] -match 'Sort-Object\s+@\{Expression=\{if\(\[string\]::IsNullOrEmpty') {
            # This looks correct, might just need brace balancing
            # Count braces
            $openBraces = ([regex]::Matches($lines[$i], '\{')).Count
            $closeBraces = ([regex]::Matches($lines[$i], '\}')).Count
            
            if ($openBraces -ne $closeBraces) {
                Write-Host "Line $($i+1) has unbalanced braces: $openBraces open, $closeBraces close" -ForegroundColor Yellow
                # The pattern should end with }}}, not }}}}, 
                $lines[$i] = $lines[$i] -replace '\}\}\}\},', '}}},'
            }
        }
    }
    
    $lines -join "`n" | Set-Content $mainPath -Encoding UTF8
    Write-Host "main.ps1 fixed" -ForegroundColor Green
}

Write-Host "`n" + ("=" * 60) -ForegroundColor DarkGray
Write-Host "Fix process completed!" -ForegroundColor Cyan
