# PowerShell Productivity Suite - Fix Guide

This document explains how to fix the PowerShell initialization errors you're experiencing.

## Issues Identified

1. **File loading order problems** - main.ps1 was trying to load files with incorrect names
2. **Missing functions** - Many expected functions were not defined
3. **PowerShell 5.1 compatibility issues** - Syntax problems and method calls
4. **Parser error in core-data.ps1** - Malformed Sort-Object expression around line 605

## Quick Fix Steps

### Step 1: Run the Setup Test
First, test the current state:
```powershell
.\test-setup.ps1
```

### Step 2: Fix Parser Error (if needed)
If core-data.ps1 fails to load, run:
```powershell
.\fix-parser-error.ps1
```

### Step 3: Test Again
Run the test again to verify:
```powershell
.\test-setup.ps1
```

### Step 4: Launch the Application
If all tests pass:
```powershell
.\main.ps1
```

## Files Created/Modified

### New Files
- **stubs.ps1** - Contains missing function implementations and PowerShell 5.1 compatibility stubs
- **fix-parser-error.ps1** - Automatically fixes common Sort-Object syntax errors
- **test-setup.ps1** - Tests the setup to identify remaining issues

### Modified Files  
- **main.ps1** - Fixed file loading order and compatibility issues

## What the Stubs Provide

The `stubs.ps1` file provides working implementations for:

- **Theme System**: `Initialize-ThemeSystem`, `Get-ThemeProperty`, `Apply-PSStyle`
- **Console Output**: `Write-Header`, `Write-Success`, `Write-Info`, `Write-Warning`, `Write-Error`
- **Dashboard & UI**: `Show-Dashboard`, `Show-Help`, `Draw-ProgressBar`
- **Time Management**: `Add-ManualTimeEntry`, `Start-Timer`, `Stop-Timer` (stubs)
- **Reports**: `Show-WeekReport`, `Export-FormattedTimesheet` (stubs)
- **UI Utilities**: `Format-TableUnicode`, `Show-Calendar`

## Key Fixes Applied

### 1. File Loading Order
Changed from:
```powershell
. "$script:ModuleRoot\helper2.ps1"       # Wrong filename
. "$script:ModuleRoot\core-data2.ps1"    # Wrong filename
```

To:
```powershell
. "$script:ModuleRoot\stubs.ps1"         # Load stubs first
. "$script:ModuleRoot\helper.ps1"        # Correct filename
. "$script:ModuleRoot\core-data.ps1"     # Correct filename
```

### 2. PowerShell 5.1 Compatibility
- Fixed `.HasKey()` → `.ContainsKey()`
- Added fallback implementations for missing PSStyle features
- Provided compatibility shims for newer PowerShell features

### 3. Missing Functions
The stubs provide working implementations that:
- Display appropriate messages about missing functionality
- Maintain the expected function signatures
- Allow the application to start and show menus
- Provide basic functionality for core features

## Expected Behavior After Fixes

After applying these fixes, you should be able to:

1. ✅ Start the application without parser errors
2. ✅ See the main dashboard
3. ✅ Navigate through menus
4. ✅ Use basic functions (though some will show "not implemented" messages)
5. ✅ Add/view projects and tasks (basic functionality)

## Next Steps

Once the application starts successfully, you can:

1. **Implement missing modules**: Create full implementations for `core-time.ps1`, `theme.ps1`, `ui.ps1`
2. **Replace stubs**: Gradually replace stub functions with full implementations
3. **Add features**: Extend functionality as needed
4. **Improve compatibility**: Add more PowerShell 5.1 specific optimizations

## Troubleshooting

### If test-setup.ps1 shows errors:
1. Check that all files exist in the modular directory
2. Verify PowerShell execution policy allows script execution
3. Run `fix-parser-error.ps1` if core-data.ps1 fails to load

### If main.ps1 still has issues:
1. Check the console output for specific error messages
2. Verify that `$script:Data` is properly initialized
3. Ensure all required functions are available

### If functions show "not implemented":
This is expected! The stubs are placeholders. Create the actual implementations in the appropriate module files as needed.

## File Structure
```
modular/
├── main.ps1              # Main entry point (modified)
├── helper.ps1            # Utility functions (existing)
├── core-data.ps1         # Data management (existing, may need parser fix)
├── stubs.ps1             # Missing function stubs (new)
├── fix-parser-error.ps1  # Parser error fix script (new)
├── test-setup.ps1        # Setup verification script (new)
├── theme.ps1             # Theme system (existing, optional)
├── ui.ps1                # UI functions (existing, optional)
└── core-time.ps1         # Time tracking (existing, optional)
```

## Support

If you continue to experience issues after following this guide:

1. Run `test-setup.ps1` and share the output
2. Check for any remaining error messages in the console
3. Verify your PowerShell version: `$PSVersionTable.PSVersion`

The fixes provided should resolve the major initialization issues and get your productivity suite running!
