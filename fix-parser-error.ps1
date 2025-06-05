# Fix for core-data.ps1 parser error at line 605
# This script contains fixes for common PowerShell Sort-Object syntax errors

# Common issue: Missing semicolon in Sort-Object @{Expression=...} blocks
# Search and replace pattern for malformed Sort-Object expressions

# This PowerShell script will fix the syntax error
$filePath = "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-data.ps1"

if (Test-Path $filePath) {
    Write-Host "Reading core-data.ps1 to fix parser error..." -ForegroundColor Yellow
    
    $content = Get-Content $filePath -Raw
    
    # Common fixes for Sort-Object syntax errors:
    
    # Fix 1: Missing semicolon between Sort-Object expressions
    $content = $content -replace 'Sort-Object UseCount -Descending, @{Expression = \{if\(', 'Sort-Object UseCount -Descending, @{Expression = {if('
    
    # Fix 2: Malformed @{Expression blocks (missing closing braces)
    $content = $content -replace '@{Expression = \{if\(\$_\.LastUsed\)\{try\{\[datetime\]\$_\.LastUsed\}catch\{\[datetime\]::MinValue\}\}else\{\[datetime\]::MinValue\}; Descending = \$true\}', '@{Expression = {if($_.LastUsed){try{[datetime]$_.LastUsed}catch{[datetime]::MinValue}}else{[datetime]::MinValue}}; Descending = $true}'
    
    # Fix 3: Missing closing braces in complex expressions
    $content = $content -replace '@{Expression = \{if\(\$_\.LastUsed\)\{try\{\[datetime\]::Parse\(\$_\.LastUsed\)\}catch\{\[datetime\]::MinValue\}\}else\{\[datetime\]::MinValue\}\}', '@{Expression = {if($_.LastUsed){try{[datetime]::Parse($_.LastUsed)}catch{[datetime]::MinValue}}else{[datetime]::MinValue}}}'
    
    # Create backup
    $backupPath = $filePath + ".backup." + (Get-Date -Format "yyyyMMdd_HHmmss")
    Copy-Item $filePath $backupPath
    Write-Host "Backup created: $backupPath" -ForegroundColor Green
    
    # Apply fix
    $content | Set-Content $filePath -Encoding UTF8
    Write-Host "Applied syntax fixes to core-data.ps1" -ForegroundColor Green
    
    # Test syntax by loading the file
    try {
        $null = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$null)
        Write-Host "✅ Syntax validation passed!" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Syntax errors still remain: $_" -ForegroundColor Red
        Write-Host "Restoring backup..." -ForegroundColor Yellow
        Copy-Item $backupPath $filePath
    }
}
else {
    Write-Host "File not found: $filePath" -ForegroundColor Red
}
