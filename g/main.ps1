# Unified Productivity Suite v5.0 - Main Entry Point
# PowerShell 7.5+ Required

#region Module Loading

# Get script directory
$script:ModuleRoot = $PSScriptRoot

# Dot source modules in dependency order
. "$script:ModuleRoot\theme.ps1"
. "$script:ModuleRoot\helper.ps1"
. "$script:ModuleRoot\ui.ps1"
. "$script:ModuleRoot\core-data.ps1"
. "$script:ModuleRoot\core-time.ps1"

# Initialize systems
Initialize-ThemeSystem
Load-UnifiedData

#endregion

#region Quick Action System

# Quick action map for +key shortcuts
$script:QuickActionMap = @{
    # Time actions
    '9' = { Add-ManualTimeEntry; return $true }
    'm' = { Add-ManualTimeEntry; return $true }
    'time' = { Add-ManualTimeEntry; return $true }
    's' = { Start-Timer; return $true }
    'timer' = { Start-Timer; return $true }
    'stop' = { Stop-Timer; return $true }
    
    # Task actions
    'a' = { Add-TodoTask; return $true }
    'task' = { Add-TodoTask; return $true }
    'qa' = { 
        $input = Read-Host "Quick add task"
        Quick-AddTask -Input $input
        return $true
    }
    
    # View actions
    'v' = { 
        Show-ActiveTimers
        Write-Host "`nPress any key to continue..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return $true
    }
    't' = { Show-TasksViewMenu; return $true }
    'today' = { Show-TodayView; return $true }
    'w' = { 
        Show-WeekReport
        Write-Host "`nPress Enter to continue..."
        Read-Host
        return $true
    }
    'week' = { 
        Show-WeekReport
        Write-Host "`nPress Enter to continue..."
        Read-Host
        return $true
    }
    
    # Project actions
    'p' = { 
        Show-ProjectDetail
        Write-Host "`nPress Enter to continue..."
        Read-Host
        return $true
    }
    'projects' = { Show-ProjectsMenu; return $true }
    
    # Command snippets
    'c' = { Manage-CommandSnippets; return $true }
    'cmd' = { Manage-CommandSnippets; return $true }
    'snippets' = { Manage-CommandSnippets; return $true }
    
    # Reports
    'r' = { Show-ReportsMenu; return $true }
    'export' = { Export-FormattedTimesheet; return $true }
    'timesheet' = { Export-FormattedTimesheet; return $true }
    
    # Help
    'h' = { Show-Help; return $true }
    'help' = { Show-Help; return $true }
    '?' = { Show-QuickActionHelp; return $true }
}

function Process-QuickAction {
    param([string]$Key)
    
    $action = $script:QuickActionMap[$Key.ToLower()]
    if ($action) {
        return & $action
    }
    
    # Check for partial matches
    $matches = $script:QuickActionMap.Keys | Where-Object { $_ -like "$Key*" }
    if ($matches.Count -eq 1) {
        $action = $script:QuickActionMap[$matches[0]]
        return & $action
    } elseif ($matches.Count -gt 1) {
        Write-Warning "Ambiguous quick action. Matches: $($matches -join ', ')"
        return $true
    }
    
    return $false
}

function Show-QuickActionHelp {
    Write-Header "Quick Actions Help"
    Write-Host "Use +key from any prompt to trigger quick actions:" -ForegroundColor Gray
    Write-Host ""
    
    Write-Host "Time Management:" -ForegroundColor Yellow
    Write-Host "  +9, +m, +time     Manual time entry"
    Write-Host "  +s, +timer        Start timer"
    Write-Host "  +stop             Stop timer"
    Write-Host ""
    
    Write-Host "Task Management:" -ForegroundColor Yellow
    Write-Host "  +a, +task         Add task"
    Write-Host "  +qa               Quick add task"
    Write-Host "  +t                Today's tasks"
    Write-Host ""
    
    Write-Host "Views & Reports:" -ForegroundColor Yellow
    Write-Host "  +v                View active timers"
    Write-Host "  +w, +week         Week report"
    Write-Host "  +today            Today view"
    Write-Host "  +timesheet        Export formatted timesheet"
    Write-Host ""
    
    Write-Host "Other:" -ForegroundColor Yellow
    Write-Host "  +p                Project details"
    Write-Host "  +c, +cmd          Command snippets"
    Write-Host "  +h, +help         Main help"
    Write-Host "  +?                This quick action help"
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

#endregion

#region Menu Structure

$script:MenuStructure = @{
    "Time Management" = @{
        Header = "Time Management"
        Options = @(
            @{Key="1"; Label="Manual Time Entry (Preferred)"; Action={Add-ManualTimeEntry}}
            @{Key="2"; Label="Start Timer"; Action={Start-Timer}}
            @{Key="3"; Label="Stop Timer"; Action={Stop-Timer}}
            @{Key="4"; Label="View Active Timers"; Action={
                Show-ActiveTimers
                Write-Host "`nPress any key to continue..."
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }}
            @{Key="5"; Label="Quick Time Entry"; Action={
                Write-Host "Enter: PROJECT HOURS [DESCRIPTION]" -ForegroundColor Gray
                $input = Read-Host "Quick entry"
                Quick-TimeEntry $input
            }}
            @{Key="6"; Label="Edit Time Entry"; Action={Edit-TimeEntry}}
            @{Key="7"; Label="Delete Time Entry"; Action={Delete-TimeEntry}}
            @{Key="8"; Label="Today's Time Log"; Action={Show-TodayTimeLog}}
            @{Key="9"; Label="Export Formatted Timesheet"; Action={Export-FormattedTimesheet}}
        )
    }
    
    "Task Management" = @{
        Header = "Task Management"
        Action = {Show-TaskManagementMenu}  # Special handling for task menu
    }
    
    "Reports & Analytics" = @{
        Header = "Reports & Analytics"
        Options = @(
            @{Key="1"; Label="Week Report (Tab-delimited)"; Action={Show-WeekReport}}
            @{Key="2"; Label="Extended Week Report"; Action={Show-ExtendedReport}}
            @{Key="3"; Label="Month Summary"; Action={Show-MonthSummary}}
            @{Key="4"; Label="Project Summary"; Action={Show-ProjectSummary}}
            @{Key="5"; Label="Task Analytics"; Action={Show-TaskAnalytics}}
            @{Key="6"; Label="Time Analytics"; Action={Show-TimeAnalytics}}
            @{Key="7"; Label="Export Data"; Action={Export-AllData}}
            @{Key="8"; Label="Formatted Timesheet"; Action={Export-FormattedTimesheet}}
            @{Key="9"; Label="Change Report Week"; Action={Change-ReportWeek}}
        )
    }
    
    "Projects & Clients" = @{
        Header = "Projects & Clients"
        Options = @(
            @{Key="1"; Label="Add Project"; Action={Add-Project}}
            @{Key="2"; Label="Import from Excel"; Action={Import-ProjectFromExcel}}
            @{Key="3"; Label="View Project Details"; Action={Show-ProjectDetail}}
            @{Key="4"; Label="Edit Project"; Action={Edit-Project}}
            @{Key="5"; Label="Configure Excel Form"; Action={Configure-ExcelForm}}
            @{Key="6"; Label="Batch Import Projects"; Action={Batch-ImportProjects}}
            @{Key="7"; Label="Export Projects"; Action={Export-Projects}}
        )
    }
    
    "Tools & Utilities" = @{
        Header = "Tools & Utilities"
        Options = @(
            @{Key="1"; Label="Command Snippets"; Action={Manage-CommandSnippets}}
            @{Key="2"; Label="Excel Copy Jobs"; Action={Show-ExcelIntegrationMenu}}
            @{Key="3"; Label="Quick Actions Help"; Action={Show-QuickActionHelp}}
            @{Key="4"; Label="Backup Data"; Action={
                Backup-Data
                Write-Host "`nPress Enter to continue..."
                Read-Host
            }}
            @{Key="5"; Label="Test Excel Connection"; Action={Test-ExcelConnection}}
        )
    }
    
    "Settings & Config" = @{
        Header = "Settings & Configuration"
        Options = @(
            @{Key="1"; Label="Time Tracking Settings"; Action={Edit-TimeTrackingSettings}}
            @{Key="2"; Label="Task Settings"; Action={Edit-TaskSettings}}
            @{Key="3"; Label="Excel Form Configuration"; Action={Configure-ExcelForm}}
            @{Key="4"; Label="Theme Settings"; Action={Edit-ThemeSettings}}
            @{Key="5"; Label="Command Snippet Settings"; Action={Edit-CommandSnippetSettings}}
            @{Key="6"; Label="Export All Data"; Action={Export-AllData}}
            @{Key="7"; Label="Import Data"; Action={Import-Data}}
            @{Key="8"; Label="Reset to Defaults"; Action={Reset-ToDefaults}}
        )
    }
}

#endregion

#region Main Functions

function Show-Menu {
    param($MenuConfig)
    
    Write-Header $MenuConfig.Header
    
    if ($MenuConfig.Options) {
        foreach ($option in $MenuConfig.Options) {
            Write-Host "[$($option.Key)] $($option.Label)"
        }
        Write-Host "`n[B] Back to Dashboard"
        
        $choice = Read-Host "`nChoice"
        
        if ($choice -eq 'B' -or $choice -eq 'b') {
            return $true
        }
        
        $selected = $MenuConfig.Options | Where-Object { $_.Key -eq $choice }
        if ($selected) {
            & $selected.Action
            if ($choice -ne "B" -and $choice -ne "b") {
                Write-Host "`nPress Enter to continue..."
                Read-Host
            }
        }
        
        return $false
    }
}

function Show-MainMenu {
    while ($true) {
        Show-Dashboard
        
        Write-Host "`nCommand: " -NoNewline -ForegroundColor Yellow
        $choice = Read-Host
        
        # Handle quick actions
        if ($choice -match '^\+(.+)$') {
            if (Process-QuickAction $matches[1]) {
                continue
            } else {
                Write-Warning "Unknown quick action: +$($matches[1]). Use +? for help."
                Start-Sleep -Seconds 1
                continue
            }
        }
        
        # Handle direct commands
        switch ($choice.ToUpper()) {
            # Quick keys (dashboard shortcuts)
            "M" { Add-ManualTimeEntry }
            "S" { Start-Timer }
            "A" { Add-TodoTask }
            "V" {
                Show-ActiveTimers
                Write-Host "`nPress any key to continue..."
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "T" { Show-TodayView }
            "W" {
                Show-WeekReport
                Write-Host "`nPress Enter to continue..."
                Read-Host
            }
            "P" {
                Show-ProjectDetail
                Write-Host "`nPress Enter to continue..."
                Read-Host
            }
            "H" { Show-Help }
            
            # Menu numbers
            "1" {
                while ($true) {
                    if (Show-Menu $script:MenuStructure["Time Management"]) { break }
                }
            }
            "2" { Show-TaskManagementMenu }
            "3" {
                while ($true) {
                    if (Show-Menu $script:MenuStructure["Reports & Analytics"]) { break }
                }
            }
            "4" {
                while ($true) {
                    Show-ProjectsAndTemplates
                    Write-Host ""
                    if (Show-Menu $script:MenuStructure["Projects & Clients"]) { break }
                }
            }
            "5" {
                while ($true) {
                    if (Show-Menu $script:MenuStructure["Tools & Utilities"]) { break }
                }
            }
            "6" {
                while ($true) {
                    Show-CurrentSettings
                    if (Show-Menu $script:MenuStructure["Settings & Config"]) { break }
                }
            }
            
            # Quit
            "Q" {
                # Save any active timers info before quitting
                if ($script:Data.ActiveTimers.Count -gt 0) {
                    Write-Warning "You have $($script:Data.ActiveTimers.Count) active timer(s) running!"
                    Write-Host "Stop all timers before quitting? (Y/N)"
                    $stopTimers = Read-Host
                    if ($stopTimers -eq 'Y' -or $stopTimers -eq 'y') {
                        foreach ($key in @($script:Data.ActiveTimers.Keys)) {
                            Stop-SingleTimer -Key $key
                        }
                    }
                }
                
                Save-UnifiedData
                Write-Host "`n👋 Thanks for using Unified Productivity Suite!" -ForegroundColor Cyan
                Write-Host "Stay productive! 🚀" -ForegroundColor Yellow
                return
            }
            
            default {
                # Check for quick commands
                if ($choice -match '^q\s+(.+)') {
                    Quick-TimeEntry $choice.Substring(2)
                }
                elseif ($choice -match '^qa\s+(.+)') {
                    Quick-AddTask -Input $choice.Substring(3)
                }
                elseif (-not [string]::IsNullOrEmpty($choice)) {
                    Write-Warning "Unknown command. Press [H] for help or +? for quick actions."
                    Start-Sleep -Seconds 1
                }
            }
        }
    }
}

function Show-TodayView {
    Clear-Host
    Write-Header "Today's Overview"
    
    $today = [DateTime]::Today
    $todayStr = $today.ToString("yyyy-MM-dd")
    
    # Today's time entries
    $todayEntries = $script:Data.TimeEntries | Where-Object { $_.Date -eq $todayStr }
    $todayHours = ($todayEntries | Measure-Object -Property Hours -Sum).Sum
    $todayHours = if ($todayHours) { [Math]::Round($todayHours, 2) } else { 0 }
    
    Write-Host "📅 " -NoNewline
    Write-Host $today.ToString("dddd, MMMM dd, yyyy") -ForegroundColor Cyan
    Write-Host ""
    
    # Time summary
    Write-Host "⏱️  TIME LOGGED: " -NoNewline -ForegroundColor Yellow
    Write-Host "$todayHours hours" -NoNewline
    $targetHours = $script:Data.Settings.HoursPerDay
    $percent = if ($targetHours -gt 0) { [Math]::Round(($todayHours / $targetHours) * 100, 0) } else { 0 }
    Write-Host " ($percent% of $targetHours hour target)" -ForegroundColor Gray
    
    # Active timers
    if ($script:Data.ActiveTimers.Count -gt 0) {
        Write-Host "`n⏰ ACTIVE TIMERS:" -ForegroundColor Red
        foreach ($timer in $script:Data.ActiveTimers.GetEnumerator()) {
            $elapsed = (Get-Date) - [DateTime]$timer.Value.StartTime
            $project = Get-ProjectOrTemplate $timer.Value.ProjectKey
            Write-Host "   → $($project.Name): $([Math]::Floor($elapsed.TotalHours)):$($elapsed.ToString('mm\:ss'))" -ForegroundColor Cyan
        }
    }
    
    # Today's tasks
    $todayTasks = $script:Data.Tasks | Where-Object {
        (-not $_.Completed) -and
        ($_.DueDate -and ([DateTime]::Parse($_.DueDate).Date -eq $today))
    }
    
    $overdueTasks = $script:Data.Tasks | Where-Object {
        (-not $_.Completed) -and
        ($_.DueDate -and ([DateTime]::Parse($_.DueDate) -lt $today))
    }
    
    if ($overdueTasks.Count -gt 0) {
        Write-Host "`n⚠️  OVERDUE TASKS ($($overdueTasks.Count)):" -ForegroundColor Red
        foreach ($task in $overdueTasks | Sort-Object DueDate, Priority | Select-Object -First 5) {
            Show-TaskItemCompact $task
        }
        if ($overdueTasks.Count -gt 5) {
            Write-Host "   ... and $($overdueTasks.Count - 5) more" -ForegroundColor DarkGray
        }
    }
    
    if ($todayTasks.Count -gt 0) {
        Write-Host "`n📋 TODAY'S TASKS ($($todayTasks.Count)):" -ForegroundColor Yellow
        foreach ($task in $todayTasks | Sort-Object Priority) {
            Show-TaskItemCompact $task
        }
    } elseif ($overdueTasks.Count -eq 0) {
        Write-Host "`n✅ No tasks due today!" -ForegroundColor Green
    }
    
    # In progress tasks
    $inProgressTasks = $script:Data.Tasks | Where-Object {
        (-not $_.Completed) -and ($_.Progress -gt 0) -and ($_.Progress -lt 100)
    }
    
    if ($inProgressTasks.Count -gt 0) {
        Write-Host "`n🔄 IN PROGRESS ($($inProgressTasks.Count)):" -ForegroundColor Blue
        foreach ($task in $inProgressTasks | Sort-Object -Descending Progress | Select-Object -First 3) {
            Show-TaskItemCompact $task
            $progressBar = "[" + ("█" * [math]::Floor($task.Progress / 10)) + ("░" * (10 - [math]::Floor($task.Progress / 10))) + "]"
            Write-Host "      $progressBar $($task.Progress)%" -ForegroundColor Green
        }
    }
    
    # Recent command snippets
    $recentCommands = Get-RecentCommandSnippets -Count 3
    if ($recentCommands.Count -gt 0) {
        Write-Host "`n💡 RECENT COMMANDS:" -ForegroundColor Magenta
        foreach ($cmd in $recentCommands) {
            Write-Host "   [$($cmd.Id.Substring(0,6))] $($cmd.Description)" -ForegroundColor White
            if ($cmd.Hotkey) {
                Write-Host "         Hotkey: $($cmd.Hotkey)" -ForegroundColor Gray
            }
        }
    }
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Show-TaskItemCompact {
    param($Task)
    
    $priorityInfo = Get-PriorityInfo $Task.Priority
    Write-Host "   $($priorityInfo.Icon) " -NoNewline
    
    if ($Task.Completed) {
        Write-Host "[$($Task.Id.Substring(0,6))] $($Task.Description)" -ForegroundColor DarkGray
    } else {
        $status = Get-TaskStatus $Task
        $color = switch ($status) {
            "Overdue" { "Red" }
            "Due Today" { "Yellow" }
            "Due Soon" { "Cyan" }
            "In Progress" { "Blue" }
            default { "White" }
        }
        Write-Host "[$($Task.Id.Substring(0,6))] $($Task.Description)" -ForegroundColor $color
    }
    
    if ($Task.ProjectKey) {
        $project = Get-ProjectOrTemplate $Task.ProjectKey
        if ($project) {
            Write-Host "      Project: $($project.Name)" -ForegroundColor Gray
        }
    }
}

function Change-ReportWeek {
    Write-Host "[P]revious, [N]ext, [T]his week, or enter date (YYYY-MM-DD): " -NoNewline
    $nav = Read-Host
    
    switch ($nav.ToUpper()) {
        'P' { $script:Data.CurrentWeek = $script:Data.CurrentWeek.AddDays(-7) }
        'N' { $script:Data.CurrentWeek = $script:Data.CurrentWeek.AddDays(7) }
        'T' { $script:Data.CurrentWeek = Get-WeekStart }
        default {
            try {
                $date = [DateTime]::Parse($nav)
                $script:Data.CurrentWeek = Get-WeekStart $date
            } catch {
                Write-Error "Invalid date format"
                return
            }
        }
    }
    Save-UnifiedData
    Write-Success "Report week changed to: $($script:Data.CurrentWeek.ToString('yyyy-MM-dd'))"
}

function Show-CurrentSettings {
    Write-Host "Current Settings:" -ForegroundColor Yellow
    Write-Host "  Default Rate:        `$$($script:Data.Settings.DefaultRate)/hour"
    Write-Host "  Hours per Day:       $($script:Data.Settings.HoursPerDay)"
    Write-Host "  Days per Week:       $($script:Data.Settings.DaysPerWeek)"
    Write-Host "  Default Priority:    $($script:Data.Settings.DefaultPriority)"
    Write-Host "  Default Category:    $($script:Data.Settings.DefaultCategory)"
    Write-Host "  Show Completed Days: $($script:Data.Settings.ShowCompletedDays)"
    Write-Host "  Auto-Archive Days:   $($script:Data.Settings.AutoArchiveDays)"
    if ($script:Data.Settings.CommandSnippets) {
        Write-Host "  Snippet Hotkeys:     $(if ($script:Data.Settings.CommandSnippets.EnableHotkeys) { 'Enabled' } else { 'Disabled' })"
        Write-Host "  Auto-Copy Snippets:  $(if ($script:Data.Settings.CommandSnippets.AutoCopyToClipboard) { 'Yes' } else { 'No' })"
    }
}

function Start-UnifiedProductivitySuite {
    Write-Host "Unified Productivity Suite v5.0" -ForegroundColor Cyan
    Write-Host "Initializing..." -ForegroundColor Gray
    
    # Show quick action tip on first run
    if (-not $script:Data.Settings.QuickActionTipShown) {
        Write-Host "`nTIP: Use +key shortcuts from any prompt for quick actions!" -ForegroundColor Yellow
        Write-Host "     Try +? to see all available quick actions." -ForegroundColor Gray
        $script:Data.Settings.QuickActionTipShown = $true
        Save-UnifiedData
        Start-Sleep -Seconds 2
    }
    
    Show-MainMenu
}

#endregion

# Entry point
Start-UnifiedProductivitySuite