# PMC Terminal Modular - Syntax Error Fixes

## Issues Identified

### 1. Parser Error in core-data.ps1 (Line 605)
- **Error**: Missing argument in parameter list
- **Location**: Search-CommandSnippets function, Sort-Object command
- **Issue**: Malformed hashtable expression in Sort-Object
- **Fix**: Properly format the Sort-Object with complete hashtable syntax

### 2. Variable Reference Error
- **Error**: Variable reference is not valid. ':' was not followed by a valid variable name
- **Location**: Line 299 (likely in helper.ps1 or core-time.ps1)
- **Issue**: Using `$[Math]::Floor` instead of `[Math]::Floor`
- **Fix**: Remove the $ before [Math]

### 3. Missing Functions
- **Functions**: Add-TodoTask, Show-ProjectsAndTemplates, Add-Project
- **Issue**: These functions ARE defined in core-data.ps1, but the file isn't loading due to parser errors
- **Fix**: Once syntax errors are fixed, functions will be available

## How to Fix

### Option 1: Quick Fix (Recommended)
```powershell
# Run the targeted fix script
.\apply-fixes.ps1

# Then reload the modules
. .\test-setup.ps1
```

### Option 2: Diagnose First
```powershell
# First, diagnose the exact issues
.\diagnose-errors.ps1

# Then apply fixes
.\apply-fixes.ps1

# Finally, reload modules
. .\test-setup.ps1
```

### Option 3: Comprehensive Fix
```powershell
# Run the comprehensive fix script
.\fix-all-errors.ps1

# Then reload modules
. .\test-setup.ps1
```

## Manual Fixes (if scripts don't work)

### Fix 1: core-data.ps1 Line 605
Find this line:
```powershell
Sort-Object UseCount -Descending, @{Expression = {if($_.LastUsed){try{[datetime]$_.LastUsed}catch{[datetime]::MinValue}}else{[datetime]::MinValue}}; Descending = $true}, Description
```

Replace with:
```powershell
Sort-Object @{Expression = {$_.UseCount}; Descending = $true}, @{Expression = {if($_.LastUsed){try{[datetime]$_.LastUsed}catch{[datetime]::MinValue}}else{[datetime]::MinValue}}; Descending = $true}, Description
```

### Fix 2: Variable Reference Errors
Find any instances of:
```powershell
$[Math]::Floor
$[Math]::Round
$[Math]::Ceiling
```

Replace with:
```powershell
[Math]::Floor
[Math]::Round
[Math]::Ceiling
```

## Testing After Fixes

1. Run `test-setup.ps1` to reload all modules
2. Try the commands that were failing:
   - Press 'a' to test Add-TodoTask
   - Press 'm' to test time entry (uses Show-ProjectsAndTemplates)
   - Press '4' then '1' to test Add-Project

## Backup Files

All fix scripts create timestamped backups before making changes:
- `*.backup.YYYYMMDD_HHMMSS`

If something goes wrong, you can restore from these backups.

## Prevention

To prevent these issues in the future:
1. Always use proper hashtable syntax in Sort-Object: `@{Expression = {...}; Descending = $true}`
2. Never use $ before type accelerators like [Math], [DateTime], etc.
3. Test each module file individually with: `[System.Management.Automation.PSParser]::Tokenize($content, [ref]$null)`
