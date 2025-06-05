# Missing Function Stubs for PowerShell 5.1 Compatibility
# Add these to your core-data2.ps1, theme2.ps1, or create separate stub files

#region Theme System Stubs (if theme2.ps1 is missing)

function global:Initialize-ThemeSystem {
    # Stub implementation - initialize basic theme system
    if (-not $script:CurrentTheme) {
        $script:CurrentTheme = @{
            Name = "Legacy"
            Palette = @{
                PrimaryFG = "White"
                SecondaryFG = "Gray"
                AccentFG = "Cyan"
                SuccessFG = "Green"
                WarningFG = "Yellow"
                ErrorFG = "Red"
                InfoFG = "Blue"
                SubtleFG = "DarkGray"
            }
            DataTable = @{
                Header = @{ FG = "Yellow"; BG = $null }
                DataRow = @{ FG = "White"; BG = $null }
                AltRow = @{ FG = "Gray"; BG = $null }
            }
        }
    }
}

function global:Get-ThemeProperty {
    param([string]$PropertyPath)
    
    if (-not $script:CurrentTheme) { Initialize-ThemeSystem }
    
    $parts = $PropertyPath -split '\.'
    $current = $script:CurrentTheme
    
    foreach ($part in $parts) {
        if ($current -is [hashtable] -and $current.ContainsKey($part)) {
            $current = $current[$part]
        } else {
            return "White" # Default fallback color
        }
    }
    
    return if ($current) { $current } else { "White" }
}

function global:Apply-PSStyle {
    param(
        [string]$Text,
        [string]$FG = $null,
        [string]$BG = $null,
        [switch]$Bold
    )
    
    # Simple fallback for PowerShell 5.1 - just return the text
    # In a full implementation, you might use ANSI escape codes
    return $Text
}

function global:Get-BorderStyleChars {
    param([string]$Style = "Single")
    
    # Unicode box drawing characters
    switch ($Style) {
        "Double" {
            return @{
                TopLeft = "╔"; TopRight = "╗"; BottomLeft = "╚"; BottomRight = "╝"
                Horizontal = "═"; Vertical = "║"
                TTop = "╦"; TBottom = "╩"; TLeft = "╠"; TRight = "╣"
                Cross = "╬"
            }
        }
        default { # Single
            return @{
                TopLeft = "┌"; TopRight = "┐"; BottomLeft = "└"; BottomRight = "┘"
                Horizontal = "─"; Vertical = "│"
                TTop = "┬"; TBottom = "┴"; TLeft = "├"; TRight = "┤"
                Cross = "┼"
            }
        }
    }
}

function global:Edit-ThemeSettings {
    Write-Header "Theme Settings"
    Write-Host "Theme settings configuration requires full theme2.ps1 implementation."
    Write-Host "Press Enter to continue..."
    Read-Host
}

#endregion

#region Console Output Stubs

function global:Write-Header {
    param([string]$Text)
    
    Clear-Host
    Write-Host ""
    Write-Host "═" * ($Text.Length + 4) -ForegroundColor Cyan
    Write-Host "  $Text  " -ForegroundColor Yellow
    Write-Host "═" * ($Text.Length + 4) -ForegroundColor Cyan
    Write-Host ""
}

function global:Write-Success {
    param([string]$Message)
    Write-Host "✅ $Message" -ForegroundColor Green
}

function global:Write-Info {
    param([string]$Message)
    Write-Host "ℹ️  $Message" -ForegroundColor Blue
}

function global:Write-Warning {
    param([string]$Message)
    Write-Host "⚠️  $Message" -ForegroundColor Yellow
}

function global:Write-Error {
    param([string]$Message)
    Write-Host "❌ $Message" -ForegroundColor Red
}

#endregion

#region Dashboard and UI Stubs

function global:Show-Dashboard {
    Clear-Host
    Write-Header "Productivity Suite Dashboard"
    
    # Show current time and date
    $now = Get-Date
    Write-Host "📅 Today: $($now.ToString('dddd, MMMM dd, yyyy'))" -ForegroundColor Cyan
    Write-Host "🕐 Current Time: $($now.ToString('HH:mm:ss'))" -ForegroundColor Cyan
    
    # Show basic stats if data exists
    if ($script:Data) {
        if ($script:Data.ActiveTimers -and $script:Data.ActiveTimers.Count -gt 0) {
            Write-Host ""
            Write-Host "⏰ Active Timers: $($script:Data.ActiveTimers.Count)" -ForegroundColor Red
        }
        
        if ($script:Data.Tasks) {
            $activeTasks = ($script:Data.Tasks | Where-Object { -not $_.Completed -and ($_.IsCommand -ne $true) }).Count
            $completedTasks = ($script:Data.Tasks | Where-Object { $_.Completed -and ($_.IsCommand -ne $true) }).Count
            Write-Host "📋 Tasks: $activeTasks active, $completedTasks completed" -ForegroundColor Yellow
        }
        
        if ($script:Data.Projects) {
            Write-Host "🏗️  Projects: $($script:Data.Projects.Count)" -ForegroundColor Magenta
        }
    }
    
    Write-Host ""
    Write-Host "Main Menu:" -ForegroundColor Yellow
    Write-Host "[M] Manual Time Entry    [S] Start Timer    [A] Add Task"
    Write-Host "[V] View Active Timers   [T] Today's View   [W] Week Report"
    Write-Host "[P] Project Details      [H] Help"
    Write-Host ""
    Write-Host "Menu Categories:" -ForegroundColor Yellow
    Write-Host "[1] Time Management      [2] Task Management"
    Write-Host "[3] Reports & Analytics  [4] Projects & Clients"
    Write-Host "[5] Tools & Utilities    [6] Settings & Config"
    Write-Host ""
    Write-Host "[Q] Quit" -ForegroundColor Red
    Write-Host ""
    Write-Host "💡 Tip: Use '+' shortcuts (e.g., +time, +task, +help)" -ForegroundColor DarkGray
}

function global:Show-Help {
    Write-Header "Help & Usage Guide"
    
    Write-Host "QUICK ACTIONS (use '+' prefix):" -ForegroundColor Yellow
    Write-Host "  +time, +m, +9  : Manual time entry"
    Write-Host "  +timer, +s     : Start timer"
    Write-Host "  +stop          : Stop timer"
    Write-Host "  +task, +a      : Add new task"
    Write-Host "  +qa            : Quick add task"
    Write-Host "  +help, +h      : This help screen"
    Write-Host ""
    
    Write-Host "MAIN MENU OPTIONS:" -ForegroundColor Yellow
    Write-Host "  [M] Manual Time Entry  - Log time for completed work"
    Write-Host "  [S] Start Timer        - Begin tracking time"
    Write-Host "  [A] Add Task          - Create new todo item"
    Write-Host "  [V] View Timers       - See active timers"
    Write-Host "  [T] Today's View      - Today's summary"
    Write-Host "  [W] Week Report       - Weekly time summary"
    Write-Host "  [P] Project Details   - View/manage projects"
    Write-Host ""
    
    Write-Host "MENU CATEGORIES:" -ForegroundColor Yellow
    Write-Host "  [1] Time Management   - All time tracking features"
    Write-Host "  [2] Task Management   - Todo lists and task tracking"
    Write-Host "  [3] Reports           - Analytics and exports"
    Write-Host "  [4] Projects          - Project and client management"
    Write-Host "  [5] Tools            - Utilities and snippets"
    Write-Host "  [6] Settings         - Configuration options"
    Write-Host ""
    
    Write-Host "For more detailed help on specific features, explore the menu categories."
    Write-Host ""
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

#endregion

#region Time Management Stubs (if core-time2.ps1 is missing)

function global:Add-ManualTimeEntry {
    Write-Header "Manual Time Entry"
    Write-Host "This feature requires the core-time2.ps1 module to be implemented."
    Write-Host "Please create or locate the missing core-time2.ps1 file."
}

function global:Start-Timer {
    param([string]$ProjectKeyParam, [string]$TaskIdParam)
    Write-Header "Start Timer"
    Write-Host "This feature requires the core-time2.ps1 module to be implemented."
    Write-Host "Params received: ProjectKey=$ProjectKeyParam, TaskId=$TaskIdParam"
}

function global:Stop-Timer {
    Write-Header "Stop Timer"
    Write-Host "This feature requires the core-time2.ps1 module to be implemented."
}

function global:Show-ActiveTimers {
    Write-Header "Active Timers"
    if ($script:Data.ActiveTimers -and $script:Data.ActiveTimers.Count -gt 0) {
        Write-Host "Active timers found, but display requires core-time2.ps1 implementation."
    } else {
        Write-Host "No active timers running." -ForegroundColor Green
    }
}

function global:Quick-TimeEntry {
    param([string]$InputString)
    Write-Host "Quick time entry requires core-time2.ps1 implementation."
    Write-Host "Input received: $InputString"
}

function global:Edit-TimeEntry {
    Write-Header "Edit Time Entry"
    Write-Host "This feature requires core-time2.ps1 implementation."
}

function global:Delete-TimeEntry {
    Write-Header "Delete Time Entry"
    Write-Host "This feature requires core-time2.ps1 implementation."
}

function global:Show-TodayTimeLog {
    Write-Header "Today's Time Log"
    Write-Host "This feature requires core-time2.ps1 implementation."
}

function global:Edit-TimeTrackingSettings {
    Write-Header "Time Tracking Settings"
    Write-Host "This feature requires core-time2.ps1 implementation."
}

function global:Stop-SingleTimer {
    param([string]$Key, [switch]$Silent)
    Write-Host "Stop single timer requires core-time2.ps1 implementation."
}

#endregion

#region Report Stubs

function global:Show-WeekReport {
    Write-Header "Week Report"
    Write-Host "Week report requires reporting module implementation."
    Write-Host "Current week: $($script:Data.CurrentWeek.ToString('yyyy-MM-dd'))"
}

function global:Show-ExtendedReport {
    Write-Header "Extended Week Report"
    Write-Host "Extended report requires implementation."
}

function global:Show-MonthSummary {
    Write-Header "Month Summary"
    Write-Host "Month summary requires implementation."
}

function global:Show-TaskAnalytics {
    Write-Header "Task Analytics"
    Write-Host "Task analytics requires implementation."
}

function global:Show-TimeAnalytics {
    Write-Header "Time Analytics"
    Write-Host "Time analytics requires implementation."
}

function global:Export-FormattedTimesheet {
    Write-Header "Export Timesheet"
    Write-Host "Timesheet export requires implementation."
}

#endregion

#region UI Utility Stubs

function global:Draw-ProgressBar {
    param([int]$Percent, [int]$Width = 20)
    
    $filled = [Math]::Floor($Width * $Percent / 100)
    $empty = $Width - $filled
    
    $bar = "█" * $filled + "░" * $empty
    Write-Host "[$bar] $Percent%" -NoNewline -ForegroundColor Green
}

function global:Format-TableUnicode {
    param(
        [array]$Data,
        [array]$Columns,
        [string]$Title = "",
        [string]$BorderStyle = "Single"
    )
    
    # Simple table implementation
    if ($Title) {
        Write-Host "`n$Title" -ForegroundColor Yellow
        Write-Host ("=" * $Title.Length) -ForegroundColor Yellow
    }
    
    # For now, just use PowerShell's built-in Format-Table
    $Data | Format-Table -Property ($Columns | ForEach-Object { $_.Name }) -AutoSize
}

function global:Show-Calendar {
    Write-Header "Calendar View"
    
    $now = Get-Date
    Write-Host "Current Month: $($now.ToString('MMMM yyyy'))" -ForegroundColor Cyan
    Write-Host ""
    
    # Simple calendar stub - just show current date
    Write-Host "📅 Today: $($now.ToString('dddd, MMMM dd, yyyy'))" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Full calendar implementation requires additional development."
    Write-Host ""
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

#endregion

#region Ensure Required Functions Exist

# Some functions that might be called from other modules
if (-not (Get-Command -Name "Show-BudgetWarning" -ErrorAction SilentlyContinue)) {
    function global:Show-BudgetWarning {
        param([string]$ProjectKey)
        # Stub implementation
        Write-Host "Budget warning system requires implementation." -ForegroundColor Yellow
    }
}

#endregion
