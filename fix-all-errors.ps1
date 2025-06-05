# Fix PowerShell Syntax Errors in PMC Terminal Modular Files
# This script fixes the identified syntax errors

Write-Host "PMC Terminal Modular - Syntax Error Fix Script" -ForegroundColor Cyan
Write-Host "=" * 50 -ForegroundColor DarkGray

# Define files to fix
$filesToFix = @{
    "core-data.ps1" = @{
        Path = "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-data.ps1"
        Fixes = @(
            @{
                Description = "Fix Sort-Object syntax error at line 605"
                # The issue is likely a missing closing brace in the Expression block
                Pattern = 'Sort-Object UseCount -Descending, @\{Expression = \{if\('
                Replace = 'Sort-Object UseCount -Descending, @{Expression = {if('
            },
            @{
                Description = "Fix complex Sort-Object expression"
                # Looking for pattern that might be missing closing braces
                Pattern = '@\{Expression = \{if\(\$_\.LastUsed\)\{try\{\[datetime\]\$_\.LastUsed\}catch\{\[datetime\]::MinValue\}\}else\{\[datetime\]::MinValue\}\}; Descending = \$true\}'
                Replace = '@{Expression = {if($_.LastUsed){try{[datetime]$_.LastUsed}catch{[datetime]::MinValue}}else{[datetime]::MinValue}}; Descending = $true}'
            }
        )
    }
    "helper.ps1" = @{
        Path = "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\helper.ps1"
        Fixes = @(
            @{
                Description = "Fix $[Math]::Floor syntax error"
                Pattern = '\$\[Math\]::'
                Replace = '[Math]::'
            }
        )
    }
    "core-time.ps1" = @{
        Path = "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-time.ps1"
        Fixes = @(
            @{
                Description = "Fix $[Math]::Floor syntax error"
                Pattern = '\$\[Math\]::'
                Replace = '[Math]::'
            }
        )
    }
}

# Process each file
foreach ($fileName in $filesToFix.Keys) {
    $fileInfo = $filesToFix[$fileName]
    $filePath = $fileInfo.Path
    
    Write-Host "`nProcessing: $fileName" -ForegroundColor Yellow
    
    if (Test-Path $filePath) {
        # Create backup
        $backupPath = $filePath + ".backup." + (Get-Date -Format "yyyyMMdd_HHmmss")
        Copy-Item $filePath $backupPath
        Write-Host "  ✓ Backup created: $(Split-Path $backupPath -Leaf)" -ForegroundColor Green
        
        # Read content
        $content = Get-Content $filePath -Raw
        $originalContent = $content
        
        # Apply fixes
        foreach ($fix in $fileInfo.Fixes) {
            Write-Host "  → Applying: $($fix.Description)" -ForegroundColor DarkCyan
            $content = $content -replace $fix.Pattern, $fix.Replace
        }
        
        # Additional fix for Sort-Object in core-data.ps1
        if ($fileName -eq "core-data.ps1") {
            # Fix the specific line that seems to be problematic
            # The error mentions line 605, let's fix Sort-Object expressions more broadly
            
            # Fix 1: Ensure proper closing of Expression blocks in Sort-Object
            $sortPattern1 = 'Sort-Object UseCount -Descending, @\{Expression\s*=\s*\{[^}]+\}\s*;\s*Descending\s*=\s*\$true\}, Description'
            $sortReplace1 = 'Sort-Object @{Expression = {$_.UseCount}; Descending = $true}, @{Expression = {if($_.LastUsed){try{[datetime]$_.LastUsed}catch{[datetime]::MinValue}}else{[datetime]::MinValue}}; Descending = $true}, Description'
            
            if ($content -match 'Sort-Object UseCount -Descending,') {
                Write-Host "  → Fixing Sort-Object expression syntax" -ForegroundColor DarkCyan
                # This is a more targeted fix for the specific pattern
                $content = $content -replace 'Sort-Object UseCount -Descending, @\{Expression = \{if\(\$_\.LastUsed\)\{try\{[^}]+\}catch\{[^}]+\}\}else\{[^}]+\}\}; Descending = \$true\}, Description', 
                                            'Sort-Object @{Expression = {$_.UseCount}; Descending = $true}, @{Expression = {if($_.LastUsed){try{[datetime]$_.LastUsed}catch{[datetime]::MinValue}}else{[datetime]::MinValue}}; Descending = $true}, Description'
            }
        }
        
        # Write fixed content if changes were made
        if ($content -ne $originalContent) {
            $content | Set-Content $filePath -Encoding UTF8
            Write-Host "  ✓ File updated successfully" -ForegroundColor Green
            
            # Test syntax
            try {
                $tokens = $null
                $errors = $null
                $null = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$errors)
                
                if ($errors.Count -eq 0) {
                    Write-Host "  ✓ Syntax validation passed!" -ForegroundColor Green
                } else {
                    Write-Host "  ⚠ Syntax errors detected:" -ForegroundColor Yellow
                    foreach ($error in $errors) {
                        Write-Host "    - Line $($error.Token.StartLine): $($error.Message)" -ForegroundColor Red
                    }
                }
            }
            catch {
                Write-Host "  ✗ Syntax validation failed: $_" -ForegroundColor Red
            }
        } else {
            Write-Host "  ℹ No changes needed" -ForegroundColor DarkGray
        }
    } else {
        Write-Host "  ✗ File not found!" -ForegroundColor Red
    }
}

Write-Host "`n" + ("=" * 50) -ForegroundColor DarkGray
Write-Host "Fix process completed!" -ForegroundColor Cyan
Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "1. Run: . .\test-setup.ps1" -ForegroundColor White
Write-Host "2. Test the application functionality" -ForegroundColor White
Write-Host "3. If errors persist, check the backup files" -ForegroundColor White
