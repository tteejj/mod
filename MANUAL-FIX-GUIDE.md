# PMC Terminal - Manual Fix Instructions

## Problem Summary

Your `core-data.ps1` file has been corrupted with:
1. **Massive duplication** - The file is over 8,800 lines (should be ~2,000)
2. **Incomplete Sort-Object statement** at line 605
3. **Repeated syntax errors** due to the duplication

## Quick Fix Steps

### Option 1: Automated Fix (Recommended)
```powershell
# This will extract the first complete copy and fix syntax errors
.\fix-duplicates.ps1

# Then reload modules
. .\test-setup.ps1
```

### Option 2: Git Restore (If you have a clean version)
```powershell
# Check git status
git status

# If core-data.ps1 is modified, restore it
git checkout -- core-data.ps1

# Then reload
. .\test-setup.ps1
```

### Option 3: Manual Emergency Fix
If the automated scripts don't work, here's what to do manually:

1. **Open core-data.ps1 in a text editor**

2. **Find line 605** (the incomplete Sort-Object)
   - Look for: `Sort-Object @{Expression = {`
   - This line is incomplete and causing the parser to fail

3. **Replace line 605 with:**
   ```powershell
                Sort-Object @{Expression = {$_.UseCount}; Descending = $true}, @{Expression = {if($_.LastUsed){try{[datetime]$_.LastUsed}catch{[datetime]::MinValue}}else{[datetime]::MinValue}}; Descending = $true}, Description
   ```

4. **Find and remove duplicate content**
   - Search for "# Core Data Management Module"
   - If you find it more than once, you have duplicates
   - Keep only the FIRST occurrence and everything up to the SECOND occurrence
   - Delete everything after the second "# Core Data Management Module"

5. **Save the file**

### Option 4: Use a Backup
Check if you have any backups:
```powershell
# List all backup files
Get-ChildItem *.backup.* | Sort-Object LastWriteTime -Descending

# If you find a good backup from before the corruption
Copy-Item "core-data.ps1.backup.TIMESTAMP" core-data.ps1
```

## Verification

After fixing, verify with:
```powershell
# Check syntax
$content = Get-Content core-data.ps1 -Raw
$errors = $null
$null = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$errors)
if ($errors.Count -eq 0) {
    Write-Host "✓ No syntax errors!" -ForegroundColor Green
} else {
    Write-Host "✗ Still has $($errors.Count) errors" -ForegroundColor Red
}

# Check file size (should be around 2000-3000 lines, not 8800)
(Get-Content core-data.ps1).Count
```

## Prevention

To prevent this in the future:
1. Always use version control (git)
2. Make regular backups before major edits
3. Use `-Encoding UTF8` when saving PowerShell files
4. Test syntax after each edit with the parser

## If All Else Fails

If none of these solutions work, you may need to:
1. Get a fresh copy of the file from the original source/repository
2. Or recreate the file from a known good version

The key functions that MUST be in core-data.ps1:
- Get-DefaultSettings
- Show-ProjectsAndTemplates
- Add-Project
- Add-TodoTask
- Search-CommandSnippets
- Complete-Task
- And all other task/project management functions
