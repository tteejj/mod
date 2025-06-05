# Diagnostic Script to Find Exact Error Lines
# This script will help identify the exact syntax errors in the PMC Terminal files

Write-Host "PMC Terminal - Syntax Error Diagnostic Tool" -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor DarkGray

$files = @(
    "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-data.ps1",
    "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\helper.ps1",
    "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-time.ps1",
    "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\main.ps1"
)

foreach ($filePath in $files) {
    if (Test-Path $filePath) {
        $fileName = Split-Path $filePath -Leaf
        Write-Host "`n`nAnalyzing: $fileName" -ForegroundColor Yellow
        Write-Host "-" * 60 -ForegroundColor DarkGray
        
        $lines = Get-Content $filePath
        $lineNum = 0
        
        # Check each line for common syntax issues
        foreach ($line in $lines) {
            $lineNum++
            $issues = @()
            
            # Check for $[Math]:: pattern
            if ($line -match '\$\[Math\]::') {
                $issues += "Invalid syntax: `$[Math]:: should be [Math]::"
            }
            
            # Check for problematic Sort-Object patterns
            if ($line -match 'Sort-Object.*UseCount.*-Descending.*@\{' -and $lineNum -ge 600 -and $lineNum -le 610) {
                $issues += "Potential Sort-Object syntax issue"
            }
            
            # Check for unbalanced braces in @{Expression blocks
            if ($line -match '@\{Expression\s*=\s*\{') {
                $openBraces = ([regex]::Matches($line, '\{')).Count
                $closeBraces = ([regex]::Matches($line, '\}')).Count
                if ($openBraces -ne $closeBraces) {
                    $issues += "Unbalanced braces in @{Expression block ($openBraces open, $closeBraces close)"
                }
            }
            
            # Report issues
            if ($issues.Count -gt 0) {
                Write-Host "`nLine $lineNum`: $($line.Trim())" -ForegroundColor Red
                foreach ($issue in $issues) {
                    Write-Host "  → $issue" -ForegroundColor Yellow
                }
            }
            
            # Special attention to lines mentioned in errors
            if ($fileName -eq "core-data.ps1" -and $lineNum -eq 605) {
                Write-Host "`nLine 605 (ERROR LINE):" -ForegroundColor Magenta
                Write-Host $line -ForegroundColor White
                Write-Host "Previous line (604): $($lines[$lineNum-2])" -ForegroundColor DarkGray
                Write-Host "Next line (606): $($lines[$lineNum])" -ForegroundColor DarkGray
            }
            
            if ($lineNum -eq 299) {
                if ($line -match '\$\[Math\]::') {
                    Write-Host "`nLine 299 (ERROR LINE in $fileName):" -ForegroundColor Magenta
                    Write-Host $line -ForegroundColor White
                }
            }
        }
        
        # Use PowerShell parser to check for syntax errors
        Write-Host "`n`nRunning PowerShell syntax parser..." -ForegroundColor Cyan
        $content = Get-Content $filePath -Raw
        $tokens = $null
        $errors = $null
        $null = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$errors)
        
        if ($errors.Count -gt 0) {
            Write-Host "Parser found $($errors.Count) error(s):" -ForegroundColor Red
            foreach ($error in $errors) {
                Write-Host "  Line $($error.Token.StartLine), Column $($error.Token.StartColumn): $($error.Message)" -ForegroundColor Red
                # Get the actual line content
                if ($error.Token.StartLine -le $lines.Count) {
                    Write-Host "  Content: $($lines[$error.Token.StartLine - 1].Trim())" -ForegroundColor Yellow
                }
            }
        } else {
            Write-Host "No syntax errors detected by parser" -ForegroundColor Green
        }
    } else {
        Write-Host "`nFile not found: $filePath" -ForegroundColor Red
    }
}

Write-Host "`n`n" + ("=" * 60) -ForegroundColor DarkGray
Write-Host "Diagnostic scan completed!" -ForegroundColor Cyan

# Additional check for function definitions
Write-Host "`n`nChecking for function definitions..." -ForegroundColor Yellow
$functionsToCheck = @("Add-TodoTask", "Show-ProjectsAndTemplates", "Add-Project")

foreach ($funcName in $functionsToCheck) {
    $found = $false
    foreach ($filePath in $files) {
        if (Test-Path $filePath) {
            $content = Get-Content $filePath -Raw
            if ($content -match "function\s+(global:)?$funcName") {
                $fileName = Split-Path $filePath -Leaf
                Write-Host "✓ $funcName found in $fileName" -ForegroundColor Green
                $found = $true
                break
            }
        }
    }
    if (-not $found) {
        Write-Host "✗ $funcName NOT FOUND in any file" -ForegroundColor Red
    }
}
