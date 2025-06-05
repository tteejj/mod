# UI Components Module
# Tables, menus, borders, and display functions

#region Table Formatting

function Format-TableUnicode {
    param(
        [Parameter(ValueFromPipeline)]
        [object[]]$Data,
        
        [hashtable[]]$Columns,
        
        [string]$BorderStyle = "Single",
        [string]$Title = "",
        [switch]$NoHeader,
        [switch]$Wrap,
        [int]$MaxWidth = 0,
        [hashtable]$RowColors = @{},
        [scriptblock]$RowColorCondition
    )
    
    begin {
        $allData = @()
        $border = Get-BorderStyleChars -Style $BorderStyle
    }
    
    process {
        $allData += $Data
    }
    
    end {
        if ($allData.Count -eq 0) {
            Write-Host "No data to display" -ForegroundColor Gray
            return
        }
        
        # Auto-generate columns if not provided
        if (-not $Columns) {
            $props = $allData[0].PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' }
            $Columns = $props | ForEach-Object {
                @{
                    Name = $_.Name
                    Title = $_.Name
                    Width = 0  # Auto-size
                }
            }
        }
        
        # Calculate column widths
        foreach ($col in $Columns) {
            if ($col.Width -eq 0) {
                # Auto-size column
                $maxLen = $col.Title.Length
                foreach ($item in $allData) {
                    $value = Get-PropertyValue $item $col.Name
                    $len = $value.ToString().Length
                    if ($len -gt $maxLen) { $maxLen = $len }
                }
                $col.Width = [Math]::Min($maxLen + 2, 50)  # Cap at 50 chars
            }
        }
        
        # Calculate total width
        $totalWidth = ($Columns | Measure-Object -Property Width -Sum).Sum + ($Columns.Count + 1)
        
        if ($MaxWidth -gt 0 -and $totalWidth -gt $MaxWidth) {
            # Adjust widths proportionally
            $factor = $MaxWidth / $totalWidth
            foreach ($col in $Columns) {
                $col.Width = [Math]::Max([int]($col.Width * $factor), 5)
            }
            $totalWidth = $MaxWidth
        }
        
        # Draw top border
        Write-Host $border.TopLeft -NoNewline
        for ($i = 0; $i -lt $Columns.Count; $i++) {
            Write-Host ($border.Horizontal * $Columns[$i].Width) -NoNewline
            if ($i -lt $Columns.Count - 1) {
                Write-Host $border.TTop -NoNewline
            }
        }
        Write-Host $border.TopRight
        
        # Draw title if provided
        if ($Title) {
            Write-Host $border.Vertical -NoNewline
            $titlePadded = " $Title ".PadRight($totalWidth - 2)
            Write-Host $titlePadded -NoNewline -ForegroundColor Cyan
            Write-Host $border.Vertical
            
            # Title separator
            Write-Host $border.TLeft -NoNewline
            for ($i = 0; $i -lt $Columns.Count; $i++) {
                Write-Host ($border.Horizontal * $Columns[$i].Width) -NoNewline
                if ($i -lt $Columns.Count - 1) {
                    Write-Host $border.Cross -NoNewline
                }
            }
            Write-Host $border.TRight
        }
        
        # Draw header
        if (-not $NoHeader) {
            Write-Host $border.Vertical -NoNewline
            foreach ($col in $Columns) {
                $headerText = Format-TableCell -Text $col.Title -Width $col.Width -Align Center
                Write-Host $headerText -NoNewline -ForegroundColor $(Get-ThemeProperty "DataTable.Header.FG")
                Write-Host $border.Vertical -NoNewline
            }
            Write-Host
            
            # Header separator
            Write-Host $border.TLeft -NoNewline
            for ($i = 0; $i -lt $Columns.Count; $i++) {
                Write-Host ($border.Horizontal * $Columns[$i].Width) -NoNewline
                if ($i -lt $Columns.Count - 1) {
                    Write-Host $border.Cross -NoNewline
                }
            }
            Write-Host $border.TRight
        }
        
        # Draw data rows
        $rowIndex = 0
        foreach ($item in $allData) {
            Write-Host $border.Vertical -NoNewline
            
            # Determine row color
            $rowColor = $null
            if ($RowColorCondition) {
                $result = & $RowColorCondition $item
                if ($result) { $rowColor = $result }
            } elseif ($rowIndex % 2 -eq 1) {
                $rowColor = "DarkGray"
            }
            
            foreach ($col in $Columns) {
                $value = Get-PropertyValue $item $col.Name
                $cellText = Format-TableCell -Text $value -Width $col.Width -Align $col.Align
                
                if ($rowColor) {
                    Write-Host $cellText -NoNewline -ForegroundColor $rowColor
                } else {
                    Write-Host $cellText -NoNewline
                }
                Write-Host $border.Vertical -NoNewline
            }
            Write-Host
            $rowIndex++
        }
        
        # Draw bottom border
        Write-Host $border.BottomLeft -NoNewline
        for ($i = 0; $i -lt $Columns.Count; $i++) {
            Write-Host ($border.Horizontal * $Columns[$i].Width) -NoNewline
            if ($i -lt $Columns.Count - 1) {
                Write-Host $border.TBottom -NoNewline
            }
        }
        Write-Host $border.BottomRight
    }
}

function Format-TableCell {
    param(
        [string]$Text,
        [int]$Width,
        [string]$Align = "Left"
    )
    
    if ($Text.Length -gt $Width - 2) {
        $Text = $Text.Substring(0, $Width - 3) + "…"
    }
    
    $padded = switch ($Align) {
        "Center" { $Text.PadLeft(($Width + $Text.Length) / 2).PadRight($Width) }
        "Right" { $Text.PadLeft($Width - 1) + " " }
        default { " " + $Text.PadRight($Width - 1) }
    }
    
    return $padded
}

function Get-PropertyValue {
    param($Object, $PropertyName)
    
    if ($PropertyName -contains ".") {
        $parts = $PropertyName -split '\.'
        $current = $Object
        foreach ($part in $parts) {
            $current = $current.$part
            if ($null -eq $current) { return "" }
        }
        return $current
    }
    
    $value = $Object.$PropertyName
    if ($null -eq $value) { return "" }
    return $value.ToString()
}

#endregion

#region Dashboard Display

function Show-Dashboard {
    Clear-Host
    
    # Use Format-TableUnicode for the header
    $headerData = @([PSCustomObject]@{
        Title = "UNIFIED PRODUCTIVITY SUITE v5.0"
        Subtitle = "All-in-One Command Center"
    })
    
    Write-Host @"
╔═══════════════════════════════════════════════════════════╗
║          UNIFIED PRODUCTIVITY SUITE v5.0                  ║
║               All-in-One Command Center                   ║
╚═══════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan

    # Quick stats
    $activeTimers = $script:Data.ActiveTimers.Count
    $activeTasks = ($script:Data.Tasks | Where-Object { -not $_.Completed }).Count
    $todayHours = ($script:Data.TimeEntries | Where-Object { $_.Date -eq (Get-Date).ToString("yyyy-MM-dd") } | Measure-Object -Property Hours -Sum).Sum
    $todayHours = if ($todayHours) { [Math]::Round($todayHours, 2) } else { 0 }
    
    Write-Host "`n📊 CURRENT STATUS" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════════════════" -ForegroundColor DarkGray
    
    Write-Host "  📅 Today: " -NoNewline
    Write-Host (Get-Date).ToString("dddd, MMMM dd, yyyy") -ForegroundColor White
    
    Write-Host "  ⏱️  Today's Hours: " -NoNewline
    if ($todayHours -gt 0) {
        Write-Host "$todayHours" -ForegroundColor Green
    } else {
        Write-Host "None logged" -ForegroundColor Gray
    }
    
    Write-Host "  ⏰ Active Timers: " -NoNewline
    if ($activeTimers -gt 0) {
        Write-Host "$activeTimers running" -ForegroundColor Red
        
        # Show active timer details
        foreach ($timer in $script:Data.ActiveTimers.GetEnumerator() | Select-Object -First 2) {
            $elapsed = (Get-Date) - [DateTime]$timer.Value.StartTime
            $project = Get-ProjectOrTemplate $timer.Value.ProjectKey
            Write-Host "     → $($project.Name): $([Math]::Floor($elapsed.TotalHours)):$($elapsed.ToString('mm\:ss'))" -ForegroundColor DarkCyan
        }
        if ($script:Data.ActiveTimers.Count -gt 2) {
            Write-Host "     → ... and $($script:Data.ActiveTimers.Count - 2) more" -ForegroundColor DarkGray
        }
    } else {
        Write-Host "None" -ForegroundColor Green
    }
    
    Write-Host "  ✅ Active Tasks: " -NoNewline
    if ($activeTasks -gt 0) {
        Write-Host "$activeTasks" -ForegroundColor Yellow
        
        # Show urgent tasks
        $overdue = $script:Data.Tasks | Where-Object {
            $_.DueDate -and ([datetime]::Parse($_.DueDate) -lt [datetime]::Today) -and -not $_.Completed
        }
        $dueToday = $script:Data.Tasks | Where-Object {
            $_.DueDate -and ([datetime]::Parse($_.DueDate).Date -eq [datetime]::Today) -and -not $_.Completed
        }
        
        if ($overdue.Count -gt 0) {
            Write-Host "     ⚠️  $($overdue.Count) overdue!" -ForegroundColor Red
        }
        if ($dueToday.Count -gt 0) {
            Write-Host "     📅 $($dueToday.Count) due today" -ForegroundColor Yellow
        }
    } else {
        Write-Host "None - inbox zero! 🎉" -ForegroundColor Green
    }
    
    Write-Host "  📁 Active Projects: " -NoNewline
    $activeProjects = ($script:Data.Projects.Values | Where-Object { $_.Status -eq "Active" }).Count
    Write-Host $activeProjects -ForegroundColor Cyan
    
    # Check for command snippets
    $commandCount = ($script:Data.Tasks | Where-Object { $_.IsCommand -eq $true }).Count
    if ($commandCount -gt 0) {
        Write-Host "  💡 Command Snippets: " -NoNewline
        Write-Host $commandCount -ForegroundColor Magenta
    }
    
    # Week summary
    Write-Host "`n📈 WEEK SUMMARY" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════════════════" -ForegroundColor DarkGray
    
    $weekStart = Get-WeekStart
    $weekEntries = $script:Data.TimeEntries | Where-Object {
        $entryDate = [DateTime]::Parse($_.Date)
        $entryDate -ge $weekStart -and $entryDate -lt $weekStart.AddDays(7)
    }
    
    $weekHours = ($weekEntries | Measure-Object -Property Hours -Sum).Sum
    $weekHours = if ($weekHours) { [Math]::Round($weekHours, 2) } else { 0 }
    
    Write-Host "  Week of: $($weekStart.ToString('MMM dd, yyyy'))"
    Write-Host "  Total Hours: $weekHours / $($script:Data.Settings.HoursPerDay * $script:Data.Settings.DaysPerWeek) target"
    
    # Progress bar for week
    $targetHours = $script:Data.Settings.HoursPerDay * $script:Data.Settings.DaysPerWeek
    $weekProgress = if ($targetHours -gt 0) { [Math]::Min(100, [Math]::Round(($weekHours / $targetHours) * 100, 0)) } else { 0 }
    $progressBar = "[" + ("█" * [math]::Floor($weekProgress / 10)) + ("░" * (10 - [math]::Floor($weekProgress / 10))) + "]"
    Write-Host "  Progress: $progressBar $weekProgress%" -ForegroundColor $(if ($weekProgress -ge 80) { "Green" } elseif ($weekProgress -ge 60) { "Yellow" } else { "Red" })
    
    # Quick actions
    Write-Host "`n⚡ QUICK ACTIONS" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════════════════" -ForegroundColor DarkGray
    Write-Host "  [M] Manual Time Entry    [S] Start Timer      [+key] Quick Actions"
    Write-Host "  [A] Add Task            [V] View Active Timers"
    Write-Host "  [T] Today's Tasks       [W] Week Report"
    Write-Host "  [P] Projects            [H] Help"
    
    Write-Host "`n🔧 FULL MENU OPTIONS" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════════════════" -ForegroundColor DarkGray
    Write-Host "  [1] Time Management     [4] Projects & Clients"
    Write-Host "  [2] Task Management     [5] Tools & Utilities"
    Write-Host "  [3] Reports & Analytics [6] Settings & Config"
    Write-Host "`n  [Q] Quit"
}

#endregion

#region Calendar Display

function Show-Calendar {
    param(
        [DateTime]$Month = (Get-Date),
        [DateTime[]]$HighlightDates = @()
    )
    
    Write-Header "Calendar - $($Month.ToString('MMMM yyyy'))"
    
    $firstDay = Get-Date $Month -Day 1
    $lastDay = $firstDay.AddMonths(1).AddDays(-1)
    $startOffset = [int]$firstDay.DayOfWeek
    
    # Header
    Write-Host "  Sun  Mon  Tue  Wed  Thu  Fri  Sat" -ForegroundColor Cyan
    Write-Host "  " + ("─" * 35) -ForegroundColor DarkGray
    
    # Calculate task counts per day
    $tasksByDate = @{}
    $script:Data.Tasks | Where-Object { $_.DueDate } | ForEach-Object {
        $date = [DateTime]::Parse($_.DueDate).Date
        if ($date.Month -eq $Month.Month -and $date.Year -eq $Month.Year) {
            if (-not $tasksByDate.ContainsKey($date)) {
                $tasksByDate[$date] = 0
            }
            $tasksByDate[$date]++
        }
    }
    
    # Days
    Write-Host -NoNewline "  "
    
    # Empty cells before first day
    for ($i = 0; $i -lt $startOffset; $i++) {
        Write-Host -NoNewline "     "
    }
    
    # Days of month
    for ($day = 1; $day -le $lastDay.Day; $day++) {
        $currentDate = Get-Date -Year $Month.Year -Month $Month.Month -Day $day
        $dayOfWeek = [int]$currentDate.DayOfWeek
        
        # Format day
        $dayStr = $day.ToString().PadLeft(3)
        
        # Determine color
        $color = "White"
        if ($currentDate.Date -eq [DateTime]::Today) {
            $color = "Green"
        } elseif ($tasksByDate.ContainsKey($currentDate.Date)) {
            $count = $tasksByDate[$currentDate.Date]
            if ($count -ge 3) {
                $color = "Red"
            } elseif ($count -ge 1) {
                $color = "Yellow"
            }
        } elseif ($dayOfWeek -eq 0 -or $dayOfWeek -eq 6) {
            $color = "DarkGray"
        }
        
        Write-Host -NoNewline $dayStr -ForegroundColor $color
        
        # Task indicator
        if ($tasksByDate.ContainsKey($currentDate.Date)) {
            Write-Host -NoNewline "*" -ForegroundColor Cyan
        } else {
            Write-Host -NoNewline " "
        }
        
        # Space or newline
        if ($dayOfWeek -eq 6) {
            Write-Host
            if ($day -lt $lastDay.Day) {
                Write-Host -NoNewline "  "
            }
        } else {
            Write-Host -NoNewline " "
        }
    }
    
    Write-Host "`n"
    
    # Legend
    Write-Host "  Legend: " -NoNewline -ForegroundColor Gray
    Write-Host "Today" -NoNewline -ForegroundColor Green
    Write-Host " | " -NoNewline -ForegroundColor Gray
    Write-Host "Tasks*" -NoNewline -ForegroundColor Cyan
    Write-Host " | " -NoNewline -ForegroundColor Gray
    Write-Host "Busy" -NoNewline -ForegroundColor Red
    
    # Navigation
    Write-Host "`n  [P]revious | [N]ext | [T]oday | [Y]ear view"
    $nav = Read-Host "  Navigation"
    
    switch ($nav.ToUpper()) {
        "P" { Show-Calendar -Month $Month.AddMonths(-1) }
        "N" { Show-Calendar -Month $Month.AddMonths(1) }
        "T" { Show-Calendar -Month (Get-Date) }
        "Y" { Show-YearCalendar -Year $Month.Year }
    }
}

function Show-YearCalendar {
    param([int]$Year = (Get-Date).Year)
    
    Write-Header "Calendar - $Year"
    
    # Display 3 months per row
    for ($row = 0; $row -lt 4; $row++) {
        $months = @()
        for ($col = 0; $col -lt 3; $col++) {
            $monthNum = $row * 3 + $col + 1
            if ($monthNum -le 12) {
                $months += Get-Date -Year $Year -Month $monthNum -Day 1
            }
        }
        
        # Month headers
        Write-Host
        foreach ($month in $months) {
            Write-Host ("  " + $month.ToString("MMMM").PadRight(20)) -NoNewline -ForegroundColor Cyan
            Write-Host "  " -NoNewline
        }
        Write-Host
        
        # Day headers
        foreach ($month in $months) {
            Write-Host "  Su Mo Tu We Th Fr Sa  " -NoNewline -ForegroundColor DarkCyan
        }
        Write-Host
        
        # Calculate max weeks needed
        $maxWeeks = 6
        
        # Display days
        for ($week = 0; $week -lt $maxWeeks; $week++) {
            foreach ($month in $months) {
                Write-Host "  " -NoNewline
                
                $firstDay = Get-Date $month -Day 1
                $lastDay = $firstDay.AddMonths(1).AddDays(-1)
                $startOffset = [int]$firstDay.DayOfWeek
                
                for ($dayOfWeek = 0; $dayOfWeek -lt 7; $dayOfWeek++) {
                    $dayNum = $week * 7 + $dayOfWeek - $startOffset + 1
                    
                    if ($dayNum -ge 1 -and $dayNum -le $lastDay.Day) {
                        $currentDate = Get-Date -Year $Year -Month $month.Month -Day $dayNum
                        
                        if ($currentDate.Date -eq [DateTime]::Today) {
                            Write-Host $dayNum.ToString().PadLeft(2) -NoNewline -ForegroundColor Green
                        } else {
                            Write-Host $dayNum.ToString().PadLeft(2) -NoNewline
                        }
                    } else {
                        Write-Host "  " -NoNewline
                    }
                    
                    if ($dayOfWeek -lt 6) {
                        Write-Host " " -NoNewline
                    }
                }
                Write-Host "  " -NoNewline
            }
            Write-Host
        }
    }
    
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

#endregion

#region Progress Bar

function Draw-ProgressBar {
    param(
        [int]$Percent,
        [int]$Width = 20,
        [string]$FillChar = "█",
        [string]$EmptyChar = "░",
        [string]$ForegroundColor = "Green",
        [string]$BackgroundColor = "DarkGray"
    )
    
    $filled = [Math]::Floor($Width * ($Percent / 100))
    $empty = $Width - $filled
    
    Write-Host "[" -NoNewline
    Write-Host ($FillChar * $filled) -NoNewline -ForegroundColor $ForegroundColor
    Write-Host ($EmptyChar * $empty) -NoNewline -ForegroundColor $BackgroundColor
    Write-Host "] $Percent%" -NoNewline
}

#endregion

#region Menu Display

function Show-MenuSelection {
    param(
        [string]$Title,
        [string[]]$Options,
        [string]$Prompt = "Select option",
        [switch]$AllowMultiple,
        [switch]$ReturnIndex
    )
    
    Write-Header $Title
    
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Host "[$($i + 1)] $($Options[$i])"
    }
    
    if ($AllowMultiple) {
        Write-Host "`nEnter numbers separated by commas (e.g., 1,3,5)"
        Write-Host "Or enter 'all' to select all, 'none' to cancel"
    } else {
        Write-Host "`n[0] Cancel"
    }
    
    $selection = Read-Host "`n$Prompt"
    
    if ($AllowMultiple) {
        if ($selection -eq 'all') {
            if ($ReturnIndex) {
                return 0..($Options.Count - 1)
            } else {
                return $Options
            }
        } elseif ($selection -eq 'none' -or [string]::IsNullOrWhiteSpace($selection)) {
            return @()
        }
        
        $indices = $selection -split ',' | ForEach-Object {
            $num = $_.Trim()
            if ($num -match '^\d+$') {
                $idx = [int]$num - 1
                if ($idx -ge 0 -and $idx -lt $Options.Count) {
                    if ($ReturnIndex) { $idx } else { $Options[$idx] }
                }
            }
        }
        return $indices
    } else {
        if ($selection -eq '0' -or [string]::IsNullOrWhiteSpace($selection)) {
            return $null
        }
        
        if ($selection -match '^\d+$') {
            $idx = [int]$selection - 1
            if ($idx -ge 0 -and $idx -lt $Options.Count) {
                if ($ReturnIndex) { return $idx } else { return $Options[$idx] }
            }
        }
        
        return $null
    }
}

#endregion

#region Help Display

function Show-Help {
    Clear-Host
    Write-Header "Help & Documentation"
    
    Write-Host @"
UNIFIED PRODUCTIVITY SUITE v5.0
===============================

This integrated suite combines time tracking, task management, project
management, Excel integration, and command snippets into a seamless 
productivity system.

QUICK ACTIONS (use +key from any prompt):
----------------------------------------
+9, +m, +time     Manual time entry
+s, +timer        Start timer
+stop             Stop timer
+a, +task         Add task
+qa               Quick add task
+t                Today's tasks
+v                View active timers
+w, +week         Week report
+today            Today view
+timesheet        Export formatted timesheet
+p                Project details
+c, +cmd          Command snippets
+h, +help         Main help
+?                Quick action help

TIME TRACKING:
-------------
- Manual entry is the preferred method for accurate time logging
- Timers are available for real-time tracking
- Link time entries to specific tasks for detailed tracking
- Budget warnings alert you when projects approach limits
- Quick entry format: PROJECT HOURS [DESCRIPTION]
- Export formatted timesheets for external systems

TASK MANAGEMENT:
---------------
- Smart sorting prioritizes urgent and overdue tasks
- Multiple views: List, Kanban, Timeline, Project
- Subtasks for breaking down complex work
- Progress tracking with automatic calculation from subtasks
- Quick add syntax: qa DESCRIPTION #category @tags !priority due:date

PROJECT MANAGEMENT:
------------------
- Import projects from Excel forms with configurable mapping
- Track budgets, billing rates, and project status
- Automatic task and time statistics per project
- Excel copy jobs for repetitive data transfers

COMMAND SNIPPETS:
----------------
- Store frequently used commands and code snippets
- Organize by category with tags for easy searching
- Optional hotkeys for quick access
- Auto-copy to clipboard feature
- Search and filter capabilities

EXCEL INTEGRATION:
-----------------
- Configure form mappings for consistent data import
- Create reusable copy configurations
- Batch import multiple projects
- Export data in multiple formats

FORMATTED TIMESHEET:
-------------------
The formatted timesheet export creates a CSV file with:
- ID1 and formatted ID2 columns
- Daily hours for the week (Mon-Fri)
- Suitable for import into external time tracking systems

KEYBOARD SHORTCUTS:
------------------
From task view:
- qa <text>  : Quick add task
- c <id>     : Complete task
- e <id>     : Edit task
- d <id>     : Delete task

DATA STORAGE:
------------
All data is stored in: $script:DataPath
Backups are automatic and kept for 30 days

TIPS:
----
1. Use manual time entry for accurate logging after completing work
2. Link tasks to projects for better organization
3. Set up Excel copy jobs for repetitive data entry
4. Use the dashboard for a quick overview of your day
5. Archive completed tasks regularly to keep views clean
6. Store common commands as snippets for quick access
7. Use +key shortcuts to navigate quickly without menus

"@ -ForegroundColor White
    
    Write-Host "`nPress any key to return..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

#endregion