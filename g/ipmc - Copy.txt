
# Unified Productivity Suite v4.0
# Complete Integration of All Modules with Full Functionality Preserved
# PowerShell 7.5+ Required

#region Configuration and Initialization

$script:DataPath = Join-Path $env:USERPROFILE ".ProductivitySuite"
$script:UnifiedDataFile = Join-Path $script:DataPath "unified_data.json"
$script:ConfigFile = Join-Path $script:DataPath "config.json"
$script:BackupPath = Join-Path $script:DataPath "backups"
$script:ExcelCopyConfigFile = Join-Path $script:DataPath "excelcopy_configs.json"

# Ensure directories exist
@($script:DataPath, $script:BackupPath) | ForEach-Object {
    if (-not (Test-Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
    }
}

# Unified Data Model - Preserving all original functionality
$script:Data = @{
    Projects = @{}      # Master project repository with full TimeTracker template support
    Tasks = @()         # Full TodoTracker task model with subtasks
    TimeEntries = @()   # All time entries with manual and timer support
    ActiveTimers = @{}  # Currently running timers
    ArchivedTasks = @() # TodoTracker archive
    ExcelCopyJobs = @{} # Saved Excel copy configurations
    CurrentWeek = Get-Date -Hour 0 -Minute 0 -Second 0
    Settings = @{
        # Time Tracker Settings
        DefaultRate = 100
        Currency = "USD"
        HoursPerDay = 8
        DaysPerWeek = 5
        TimeTrackerTemplates = @{
            "ADMIN" = @{
                Id1 = "100"
                Id2 = "ADM"
                Name = "Administrative Tasks"
                Client = "Internal"
                Department = "Operations"
                BillingType = "Non-Billable"
                Status = "Active"
                Budget = 0
                Rate = 0
                Notes = "General administrative tasks"
            }
            "MEETING" = @{
                Id1 = "101"
                Id2 = "MTG"
                Name = "Meetings & Calls"
                Client = "Internal"
                Department = "Various"
                BillingType = "Non-Billable"
                Status = "Active"
                Budget = 0
                Rate = 0
                Notes = "Team meetings and calls"
            }
            "TRAINING" = @{
                Id1 = "102"
                Id2 = "TRN"
                Name = "Training & Learning"
                Client = "Internal"
                Department = "HR"
                BillingType = "Non-Billable"
                Status = "Active"
                Budget = 0
                Rate = 0
                Notes = "Professional development"
            }
            "BREAK" = @{
                Id1 = "103"
                Id2 = "BRK"
                Name = "Breaks & Personal"
                Client = "Internal"
                Department = "Personal"
                BillingType = "Non-Billable"
                Status = "Active"
                Budget = 0
                Rate = 0
                Notes = "Breaks and personal time"
            }
        }
        # Todo Tracker Settings
        DefaultPriority = "Medium"
        DefaultCategory = "General"
        ShowCompletedDays = 7
        EnableTimeTracking = $true
        AutoArchiveDays = 30
        # Excel Integration Settings
        ExcelFormConfig = @{
            WorksheetName = "Project Info"
            StandardFields = @{
                "Id1" = @{ LabelCell = "A5"; ValueCell = "B5"; Label = "Project ID" }
                "Id2" = @{ LabelCell = "A6"; ValueCell = "B6"; Label = "Task Code" }
                "Name" = @{ LabelCell = "A7"; ValueCell = "B7"; Label = "Project Name" }
                "FullName" = @{ LabelCell = "A8"; ValueCell = "B8"; Label = "Full Description" }
                "AssignedDate" = @{ LabelCell = "A9"; ValueCell = "B9"; Label = "Start Date" }
                "DueDate" = @{ LabelCell = "A10"; ValueCell = "B10"; Label = "End Date" }
                "Manager" = @{ LabelCell = "A11"; ValueCell = "B11"; Label = "Project Manager" }
                "Budget" = @{ LabelCell = "A12"; ValueCell = "B12"; Label = "Budget" }
                "Status" = @{ LabelCell = "A13"; ValueCell = "B13"; Label = "Status" }
                "Priority" = @{ LabelCell = "A14"; ValueCell = "B14"; Label = "Priority" }
                "Department" = @{ LabelCell = "A15"; ValueCell = "B15"; Label = "Department" }
                "Client" = @{ LabelCell = "A16"; ValueCell = "B16"; Label = "Client" }
                "BillingType" = @{ LabelCell = "A17"; ValueCell = "B17"; Label = "Billing Type" }
                "Rate" = @{ LabelCell = "A18"; ValueCell = "B18"; Label = "Hourly Rate" }
            }
        }
        # UI Theme
        Theme = @{
            Header = "Cyan"
            Success = "Green"
            Warning = "Yellow"
            Error = "Red"
            Info = "Blue"
            Accent = "Magenta"
            Subtle = "DarkGray"
        }
    }
}

#endregion

#region Core Utility Functions

function Load-UnifiedData {
    try {
        if (Test-Path $script:UnifiedDataFile) {
            $loaded = Get-Content $script:UnifiedDataFile | ConvertFrom-Json -AsHashtable
           
            # Deep merge to preserve structure
            foreach ($key in $loaded.Keys) {
                if ($key -eq "Settings" -and $script:Data.ContainsKey($key)) {
                    # Merge settings carefully to preserve defaults
                    foreach ($settingKey in $loaded.Settings.Keys) {
                        $script:Data.Settings[$settingKey] = $loaded.Settings[$settingKey]
                    }
                } else {
                    $script:Data[$key] = $loaded[$key]
                }
            }
           
            # Ensure CurrentWeek is a DateTime
            if ($script:Data.CurrentWeek -is [string]) {
                $script:Data.CurrentWeek = [DateTime]::Parse($script:Data.CurrentWeek)
            }
        }
    } catch {
        Write-Warning "Could not load data, starting fresh: $_"
    }
}

function Save-UnifiedData {
    try {
        # Auto-backup
        if ((Get-Random -Maximum 10) -eq 0 -or -not (Test-Path $script:UnifiedDataFile)) {
            Backup-Data -Silent
        }
       
        $script:Data | ConvertTo-Json -Depth 10 | Set-Content $script:UnifiedDataFile
    } catch {
        Write-Error "Failed to save data: $_"
    }
}

function Backup-Data {
    param([switch]$Silent)
   
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupFile = Join-Path $script:BackupPath "backup_$timestamp.json"
   
    try {
        $script:Data | ConvertTo-Json -Depth 10 | Set-Content $backupFile
        if (-not $Silent) {
            Write-Success "Backup created: $backupFile"
        }
       
        # Keep only last 30 backups
        Get-ChildItem $script:BackupPath -Filter "backup_*.json" |
            Sort-Object CreationTime -Descending |
            Select-Object -Skip 30 |
            Remove-Item -Force
    } catch {
        Write-Error "Backup failed: $_"
    }
}

function Write-Header {
    param([string]$Text)
    Write-Host "`n$Text" -ForegroundColor $script:Data.Settings.Theme.Header
    Write-Host ("=" * $Text.Length) -ForegroundColor DarkCyan
}

function Write-Success {
    param([string]$Text)
    Write-Host "✓ $Text" -ForegroundColor $script:Data.Settings.Theme.Success
}

function Write-Warning {
    param([string]$Text)
    Write-Host "⚠ $Text" -ForegroundColor $script:Data.Settings.Theme.Warning
}

function Write-Error {
    param([string]$Text)
    Write-Host "✗ $Text" -ForegroundColor $script:Data.Settings.Theme.Error
}

function Write-Info {
    param([string]$Text)
    Write-Host "ℹ $Text" -ForegroundColor $script:Data.Settings.Theme.Info
}

function New-TodoId {
    return [System.Guid]::NewGuid().ToString().Substring(0, 8)
}

function Format-Id2 {
    param([string]$Id2)
   
    if ($Id2.Length -gt 9) {
        $Id2 = $Id2.Substring(0, 9)
    }
   
    $paddingNeeded = 12 - 2 - $Id2.Length
    $zeros = "0" * $paddingNeeded
   
    return "V${zeros}${Id2}S"
}

function Get-WeekStart {
    param([DateTime]$Date = (Get-Date))
   
    $daysFromMonday = [int]$Date.DayOfWeek
    if ($daysFromMonday -eq 0) { $daysFromMonday = 7 }
    $monday = $Date.AddDays(1 - $daysFromMonday)
   
    return Get-Date $monday -Hour 0 -Minute 0 -Second 0
}

function Get-WeekDates {
    param([DateTime]$WeekStart)
   
    return @(0..4 | ForEach-Object { $WeekStart.AddDays($_) })
}

function Format-TodoDate {
    param($DateString)
    if ([string]::IsNullOrEmpty($DateString)) { return "" }
    try {
        $date = [datetime]::Parse($DateString)
        $today = [datetime]::Today
        $diff = ($date - $today).Days
       
        $dateStr = $date.ToString("MMM dd")
        if ($diff -eq 0) { return "Today" }
        elseif ($diff -eq 1) { return "Tomorrow" }
        elseif ($diff -eq -1) { return "Yesterday" }
        elseif ($diff -gt 0 -and $diff -le 7) { return "$dateStr (in $diff days)" }
        elseif ($diff -lt 0) {
            $absDiff = [Math]::Abs($diff)
            return "$dateStr ($absDiff days ago)"
        }
        else { return $dateStr }
    }
    catch { return $DateString }
}

function Get-PriorityInfo {
    param($Priority)
    switch ($Priority) {
        "Critical" { return @{ Color = "Magenta"; Icon = "🔥" } }
        "High" { return @{ Color = "Red"; Icon = "🔴" } }
        "Medium" { return @{ Color = "Yellow"; Icon = "🟡" } }
        "Low" { return @{ Color = "Green"; Icon = "🟢" } }
        default { return @{ Color = "Gray"; Icon = "⚪" } }
    }
}

function Get-ProjectOrTemplate {
    param([string]$Key)
   
    if ($script:Data.Projects.ContainsKey($Key)) {
        return $script:Data.Projects[$Key]
    } elseif ($script:Data.Settings.TimeTrackerTemplates.ContainsKey($Key)) {
        return $script:Data.Settings.TimeTrackerTemplates[$Key]
    }
   
    return $null
}

#endregion

#region Time Tracking Functions - Full TimeTracker.ps1 functionality

function Add-ManualTimeEntry {
    Write-Header "Manual Time Entry"
   
    # Show projects and templates
    Show-ProjectsAndTemplates
   
    $projectKey = Read-Host "`nProject/Template Key"
    $project = Get-ProjectOrTemplate $projectKey
   
    if (-not $project) {
        Write-Error "Project not found"
        return
    }
   
    # Date entry with smart parsing
    $dateStr = Read-Host "Date (YYYY-MM-DD, 'today', 'yesterday', or press Enter for today)"
    if ([string]::IsNullOrWhiteSpace($dateStr)) {
        $date = (Get-Date).Date
    } else {
        try {
            $date = switch ($dateStr.ToLower()) {
                'today' { [DateTime]::Today }
                'yesterday' { [DateTime]::Today.AddDays(-1) }
                default { [DateTime]::Parse($dateStr) }
            }
        } catch {
            Write-Error "Invalid date format"
            return
        }
    }
   
    Write-Host "Enter time as hours (e.g., 2.5) or time range (e.g., 09:00-11:30)"
    $timeInput = Read-Host "Time"
   
    $hours = 0
    $startTime = ""
    $endTime = ""
   
    # Parse time input
    if ($timeInput -match '(\d{1,2}:\d{2})-(\d{1,2}:\d{2})') {
        try {
            $start = [DateTime]::Parse("$date $($Matches[1])")
            $end = [DateTime]::Parse("$date $($Matches[2])")
           
            if ($end -lt $start) {
                $end = $end.AddDays(1)
            }
           
            $hours = ($end - $start).TotalHours
            $startTime = $Matches[1]
            $endTime = $Matches[2]
           
            Write-Info "Calculated hours: $([Math]::Round($hours, 2))"
        } catch {
            Write-Error "Invalid time format"
            return
        }
    } else {
        try {
            $hours = [double]$timeInput
        } catch {
            Write-Error "Invalid hours format"
            return
        }
    }
   
    $description = Read-Host "Description (optional)"
   
    # Link to task?
    Write-Host "`nLink to a task? (Y/N)"
    $linkTask = Read-Host
    $taskId = $null
   
    if ($linkTask -eq 'Y' -or $linkTask -eq 'y') {
        $projectTasks = $script:Data.Tasks | Where-Object { $_.ProjectKey -eq $projectKey -and -not $_.Completed }
        if ($projectTasks.Count -gt 0) {
            Write-Host "`nActive tasks for this project:"
            foreach ($task in $projectTasks) {
                Write-Host "  [$($task.Id.Substring(0,6))] $($task.Description)"
            }
            $taskId = Read-Host "`nTask ID (partial ok)"
           
            $matchedTask = $script:Data.Tasks | Where-Object { $_.Id -like "$taskId*" } | Select-Object -First 1
            if ($matchedTask) {
                $taskId = $matchedTask.Id
                # Update task time spent
                $matchedTask.TimeSpent = [Math]::Round($matchedTask.TimeSpent + $hours, 2)
            } else {
                Write-Warning "Task not found, proceeding without task link"
                $taskId = $null
            }
        }
    }
   
    # Create entry
    $entry = @{
        Id = New-TodoId
        ProjectKey = $projectKey
        TaskId = $taskId
        Date = $date.ToString("yyyy-MM-dd")
        Hours = [Math]::Round($hours, 2)
        Description = $description
        StartTime = $startTime
        EndTime = $endTime
        EnteredAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }
   
    $script:Data.TimeEntries += $entry
    Update-ProjectStatistics -ProjectKey $projectKey
    Save-UnifiedData
   
    Write-Success "Time entry added: $($entry.Hours) hours for $($project.Name) on $($date.ToString('yyyy-MM-dd'))"
   
    # Check budget warning
    Show-BudgetWarning -ProjectKey $projectKey
}

function Start-Timer {
    param(
        [string]$ProjectKey,
        [string]$TaskId,
        [string]$Description
    )
   
    if (-not $ProjectKey -and -not $TaskId) {
        Write-Header "Start Timer"
       
        Write-Host "[P] Timer for Project/Template"
        Write-Host "[T] Timer for Task"
        $choice = Read-Host "Choice"
       
        if ($choice -eq 'T' -or $choice -eq 't') {
            # Show active tasks
            $activeTasks = $script:Data.Tasks | Where-Object { -not $_.Completed }
            if ($activeTasks.Count -eq 0) {
                Write-Warning "No active tasks available"
                return
            }
           
            # Group by project
            $grouped = $activeTasks | Group-Object ProjectKey
            foreach ($group in $grouped) {
                $projectName = if ($group.Name) {
                    $proj = Get-ProjectOrTemplate $group.Name
                    if ($proj) { $proj.Name } else { "No Project" }
                } else {
                    "No Project"
                }
               
                Write-Host "`n$projectName" -ForegroundColor $script:Data.Settings.Theme.Accent
                foreach ($task in $group.Group | Sort-Object Priority) {
                    $priorityInfo = Get-PriorityInfo $task.Priority
                    Write-Host "  $($priorityInfo.Icon) [$($task.Id.Substring(0,6))] $($task.Description)"
                }
            }
           
            $TaskId = Read-Host "`nTask ID (partial ok)"
            $task = $script:Data.Tasks | Where-Object { $_.Id -like "$TaskId*" } | Select-Object -First 1
           
            if (-not $task) {
                Write-Error "Task not found"
                return
            }
           
            $TaskId = $task.Id
            $ProjectKey = $task.ProjectKey
        } else {
            Show-ProjectsAndTemplates
            $ProjectKey = Read-Host "`nProject/Template Key"
           
            if (-not (Get-ProjectOrTemplate $ProjectKey)) {
                Write-Error "Project not found"
                return
            }
        }
    }
   
    if (-not $Description) {
        $Description = Read-Host "Description (optional)"
    }
   
    # Check for existing timer on same project/task
    $existingKey = if ($TaskId) { $TaskId } else { $ProjectKey }
    if ($script:Data.ActiveTimers.ContainsKey($existingKey)) {
        Write-Warning "Timer already running for this item!"
        return
    }
   
    # Create timer
    $timer = @{
        StartTime = Get-Date
        ProjectKey = $ProjectKey
        TaskId = $TaskId
        Description = $Description
    }
   
    $script:Data.ActiveTimers[$existingKey] = $timer
    Save-UnifiedData
   
    Write-Success "Timer started!"
   
    if ($TaskId) {
        $task = $script:Data.Tasks | Where-Object { $_.Id -eq $TaskId }
        Write-Host "Task: $($task.Description)" -ForegroundColor Gray
    }
   
    $project = Get-ProjectOrTemplate $ProjectKey
    Write-Host "Project: $($project.Name)" -ForegroundColor Gray
    Write-Host "Started at: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Gray
}

function Stop-Timer {
    param([string]$TimerKey)
   
    if ($script:Data.ActiveTimers.Count -eq 0) {
        Write-Warning "No active timers"
        return
    }
   
    if (-not $TimerKey) {
        Show-ActiveTimers
        $TimerKey = Read-Host "`nStop timer (ID/Key, or 'all' to stop all)"
    }
   
    if ($TimerKey -eq 'all') {
        $count = $script:Data.ActiveTimers.Count
        foreach ($key in @($script:Data.ActiveTimers.Keys)) {
            Stop-SingleTimer -Key $key -Silent
        }
        Write-Success "Stopped $count timer(s)"
        Save-UnifiedData
        return
    }
   
    # Find matching timer
    $matchedKey = $null
    foreach ($key in $script:Data.ActiveTimers.Keys) {
        if ($key -like "$TimerKey*") {
            $matchedKey = $key
            break
        }
    }
   
    if (-not $matchedKey) {
        Write-Error "Timer not found"
        return
    }
   
    Stop-SingleTimer -Key $matchedKey
    Save-UnifiedData
}

function Stop-SingleTimer {
    param(
        [string]$Key,
        [switch]$Silent
    )
   
    $timer = $script:Data.ActiveTimers[$Key]
    if (-not $timer) { return }
   
    $endTime = Get-Date
    $duration = ($endTime - [DateTime]$timer.StartTime).TotalHours
   
    # Create time entry
    $entry = @{
        Id = New-TodoId
        ProjectKey = $timer.ProjectKey
        TaskId = $timer.TaskId
        Date = $endTime.Date.ToString("yyyy-MM-dd")
        Hours = [Math]::Round($duration, 2)
        StartTime = ([DateTime]$timer.StartTime).ToString("HH:mm")
        EndTime = $endTime.ToString("HH:mm")
        Description = $timer.Description
        EnteredAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }
   
    $script:Data.TimeEntries += $entry
   
    # Update task time if applicable
    if ($timer.TaskId) {
        $task = $script:Data.Tasks | Where-Object { $_.Id -eq $timer.TaskId }
        if ($task) {
            $task.TimeSpent = [Math]::Round($task.TimeSpent + $duration, 2)
        }
    }
   
    # Update project stats
    Update-ProjectStatistics -ProjectKey $timer.ProjectKey
   
    # Remove timer
    $script:Data.ActiveTimers.Remove($Key)
   
    if (-not $Silent) {
        $project = Get-ProjectOrTemplate $timer.ProjectKey
        Write-Success "Timer stopped for: $($project.Name)"
        Write-Host "Duration: $([Math]::Round($duration, 2)) hours ($($entry.StartTime) - $($entry.EndTime))"
       
        # Check budget warning
        Show-BudgetWarning -ProjectKey $timer.ProjectKey
    }
}

function Show-ActiveTimers {
    Write-Header "Active Timers"
   
    if ($script:Data.ActiveTimers.Count -eq 0) {
        Write-Host "No active timers" -ForegroundColor Gray
        return
    }
   
    $totalElapsed = [TimeSpan]::Zero
   
    foreach ($timer in $script:Data.ActiveTimers.GetEnumerator()) {
        $elapsed = (Get-Date) - [DateTime]$timer.Value.StartTime
        $totalElapsed += $elapsed
        $project = Get-ProjectOrTemplate $timer.Value.ProjectKey
       
        Write-Host "`n[$($timer.Key.Substring(0,6))] " -NoNewline -ForegroundColor Yellow
       
        if ($timer.Value.TaskId) {
            $task = $script:Data.Tasks | Where-Object { $_.Id -eq $timer.Value.TaskId }
            Write-Host "Task: $($task.Description)"
        }
       
        Write-Host "Project: $($project.Name) ($($project.Client))"
        Write-Host "  Started: $($timer.Value.StartTime.ToString('HH:mm:ss'))" -ForegroundColor Gray
        Write-Host "  Elapsed: $([Math]::Floor($elapsed.TotalHours)):$($elapsed.ToString('mm\:ss'))" -ForegroundColor Cyan
       
        if ($timer.Value.Description) {
            Write-Host "  Note: $($timer.Value.Description)" -ForegroundColor Gray
        }
       
        if ($project.BillingType -eq "Billable" -and $project.Rate -gt 0) {
            $value = $elapsed.TotalHours * $project.Rate
            Write-Host "  Value: `$$([Math]::Round($value, 2))" -ForegroundColor Green
        }
    }
   
    if ($script:Data.ActiveTimers.Count -gt 1) {
        Write-Host "`nTotal Time: $([Math]::Floor($totalElapsed.TotalHours)):$($totalElapsed.ToString('mm\:ss'))" -ForegroundColor Cyan
    }
}

function Quick-TimeEntry {
    param([string]$Input)
   
    # Format: PROJECT HOURS [DESCRIPTION]
    $parts = $Input -split ' ', 3
    if ($parts.Count -lt 2) {
        Write-Error "Format: PROJECT HOURS [DESCRIPTION]"
        return
    }
   
    $projectKey = $parts[0]
    $hours = try { [double]$parts[1] } catch { 0 }
    $description = if ($parts.Count -eq 3) { $parts[2] } else { "" }
   
    if ($hours -eq 0) {
        Write-Error "Invalid hours format"
        return
    }
   
    $project = Get-ProjectOrTemplate $projectKey
    if (-not $project) {
        Write-Error "Unknown project: $projectKey"
        return
    }
   
    $entry = @{
        Id = New-TodoId
        ProjectKey = $projectKey
        TaskId = $null
        Date = (Get-Date).Date.ToString("yyyy-MM-dd")
        Hours = [Math]::Round($hours, 2)
        Description = $description
        StartTime = ""
        EndTime = ""
        EnteredAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }
   
    $script:Data.TimeEntries += $entry
    Update-ProjectStatistics -ProjectKey $projectKey
    Save-UnifiedData
   
    Write-Success "Quick entry: $hours hours for $($project.Name)"
}

function Show-BudgetWarning {
    param([string]$ProjectKey)
   
    $project = Get-ProjectOrTemplate $ProjectKey
    if (-not $project -or $project.BillingType -eq "Non-Billable" -or -not $project.Budget -or $project.Budget -eq 0) {
        return
    }
   
    Update-ProjectStatistics -ProjectKey $ProjectKey
    $percentUsed = ($project.TotalHours / $project.Budget) * 100
   
    if ($percentUsed -ge 90) {
        Write-Warning "Budget alert: $([Math]::Round($percentUsed, 1))% used ($([Math]::Round($project.Budget - $project.TotalHours, 2)) hours remaining)"
    } elseif ($percentUsed -ge 75) {
        Write-Warning "Budget notice: $([Math]::Round($percentUsed, 1))% used"
    }
}

function Update-ProjectStatistics {
    param([string]$ProjectKey)
   
    $project = $script:Data.Projects[$ProjectKey]
    if (-not $project) { return }
   
    # Calculate total hours
    $projectEntries = $script:Data.TimeEntries | Where-Object { $_.ProjectKey -eq $ProjectKey }
    $project.TotalHours = [Math]::Round(($projectEntries | Measure-Object -Property Hours -Sum).Sum, 2)
   
    # Update task counts
    $projectTasks = $script:Data.Tasks | Where-Object { $_.ProjectKey -eq $ProjectKey }
    $project.CompletedTasks = ($projectTasks | Where-Object { $_.Completed }).Count
    $project.ActiveTasks = ($projectTasks | Where-Object { -not $_.Completed }).Count
}

#endregion

#region Task Management Functions - Full TodoTracker.ps1 functionality

function Add-TodoTask {
    Write-Header "Add New Task"
   
    # Description
    $description = Read-Host "`nTask description"
    if ([string]::IsNullOrEmpty($description)) {
        Write-Error "Task description cannot be empty!"
        return
    }
   
    # Priority
    Write-Host "`nPriority: [C]ritical, [H]igh, [M]edium, [L]ow (default: $($script:Data.Settings.DefaultPriority))" -ForegroundColor Gray
    $priorityInput = Read-Host "Priority"
    $priority = switch ($priorityInput.ToUpper()) {
        "C" { "Critical" }
        "H" { "High" }
        "L" { "Low" }
        "M" { "Medium" }
        default { $script:Data.Settings.DefaultPriority }
    }
   
    # Category
    $existingCategories = $script:Data.Tasks | Select-Object -ExpandProperty Category -Unique | Where-Object { $_ }
    if ($existingCategories) {
        Write-Host "`nExisting categories: $($existingCategories -join ', ')" -ForegroundColor DarkCyan
    }
    $category = Read-Host "Category (default: $($script:Data.Settings.DefaultCategory))"
    if ([string]::IsNullOrEmpty($category)) {
        $category = $script:Data.Settings.DefaultCategory
    }
   
    # Project
    Write-Host "`nLink to project? (Y/N)"
    $linkProject = Read-Host
    $projectKey = $null
   
    if ($linkProject -eq 'Y' -or $linkProject -eq 'y') {
        Show-ProjectsAndTemplates -Simple
        $projectKey = Read-Host "`nProject key"
       
        if ($projectKey -and -not (Get-ProjectOrTemplate $projectKey)) {
            Write-Warning "Project not found, task will be created without project link"
            $projectKey = $null
        }
    }
   
    # Dates
    Write-Host "`nStart date (optional): Enter date, 'today', 'tomorrow', or '+X' for X days from now" -ForegroundColor Gray
    $startDateInput = Read-Host "Start date"
    $startDate = $null
    if ($startDateInput) {
        try {
            $startDate = switch -Regex ($startDateInput.ToLower()) {
                '^today$' { [datetime]::Today }
                '^tomorrow$' { [datetime]::Today.AddDays(1) }
                '^\+(\d+)$' { [datetime]::Today.AddDays([int]$Matches[1]) }
                default { [datetime]::Parse($startDateInput) }
            }
            $startDate = $startDate.ToString("yyyy-MM-dd")
        }
        catch {
            Write-Warning "Invalid date format. Start date not set."
            $startDate = $null
        }
    }
   
    Write-Host "`nDue date (optional): Enter date, 'today', 'tomorrow', or '+X' for X days from now" -ForegroundColor Gray
    $dueDateInput = Read-Host "Due date"
    $dueDate = $null
    if ($dueDateInput) {
        try {
            $dueDate = switch -Regex ($dueDateInput.ToLower()) {
                '^today$' { [datetime]::Today }
                '^tomorrow$' { [datetime]::Today.AddDays(1) }
                '^\+(\d+)$' { [datetime]::Today.AddDays([int]$Matches[1]) }
                default { [datetime]::Parse($dueDateInput) }
            }
            $dueDate = $dueDate.ToString("yyyy-MM-dd")
        }
        catch {
            Write-Warning "Invalid date format. Due date not set."
            $dueDate = $null
        }
    }
   
    # Tags
    Write-Host "`nTags (comma-separated, optional):" -ForegroundColor Gray
    $tagsInput = Read-Host "Tags"
    $tags = if ($tagsInput) {
        $tagsInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    } else { @() }
   
    # Estimated time
    $estimatedTime = Read-Host "`nEstimated time in hours (optional)"
    $estimatedTime = if ($estimatedTime) { try { [double]$estimatedTime } catch { 0 } } else { 0 }
   
    # Add subtasks?
    $subtasks = @()
    Write-Host "`nAdd subtasks? (Y/N)" -ForegroundColor Gray
    $addSubtasks = Read-Host
    if ($addSubtasks -eq 'Y' -or $addSubtasks -eq 'y') {
        Write-Host "Enter subtasks (empty line to finish):" -ForegroundColor Gray
        while ($true) {
            $subtaskDesc = Read-Host "  Subtask"
            if ([string]::IsNullOrEmpty($subtaskDesc)) { break }
           
            $subtasks += @{
                Description = $subtaskDesc
                Completed = $false
                CompletedDate = $null
            }
        }
    }
   
    # Create task
    $newTask = @{
        Id = New-TodoId
        Description = $description
        Priority = $priority
        Category = $category
        ProjectKey = $projectKey
        StartDate = $startDate
        DueDate = $dueDate
        Tags = $tags
        Progress = 0
        Completed = $false
        CreatedDate = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
        CompletedDate = $null
        EstimatedTime = $estimatedTime
        TimeSpent = 0
        Subtasks = $subtasks
        Notes = ""
        LastModified = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
    }
   
    $script:Data.Tasks += $newTask
   
    if ($projectKey) {
        Update-ProjectStatistics -ProjectKey $projectKey
    }
   
    Save-UnifiedData
   
    Write-Success "Task added successfully!"
    Write-Host "ID: $($newTask.Id)" -ForegroundColor DarkGray
   
    # Ask if want to start timer
    if ($script:Data.Settings.EnableTimeTracking -and $projectKey) {
        Write-Host "`nStart timer for this task? (Y/N)" -ForegroundColor Gray
        $startTimer = Read-Host
        if ($startTimer -eq 'Y' -or $startTimer -eq 'y') {
            Start-Timer -ProjectKey $projectKey -TaskId $newTask.Id
        }
    }
}

function Quick-AddTask {
    param([string]$Input)
   
    if (-not $Input) {
        $Input = Read-Host "Quick add task"
    }
   
    # Parse syntax: "description #category @tag1,tag2 !priority due:date project:key est:hours"
    $description = $Input
    $category = $script:Data.Settings.DefaultCategory
    $tags = @()
    $priority = $script:Data.Settings.DefaultPriority
    $dueDate = $null
    $startDate = $null
    $projectKey = $null
    $estimatedTime = 0
   
    # Extract category
    if ($Input -match '#(\w+)') {
        $category = $Matches[1]
        $description = $description -replace '#\w+', ''
    }
   
    # Extract tags
    if ($Input -match '@([\w,]+)') {
        $tags = $Matches[1] -split ',' | ForEach-Object { $_.Trim() }
        $description = $description -replace '@[\w,]+', ''
    }
   
    # Extract priority
    if ($Input -match '!(critical|high|medium|low|c|h|m|l)') {
        $priority = switch ($Matches[1].ToLower()) {
            "c" { "Critical" }
            "critical" { "Critical" }
            "h" { "High" }
            "high" { "High" }
            "l" { "Low" }
            "low" { "Low" }
            default { "Medium" }
        }
        $description = $description -replace '!(critical|high|medium|low|c|h|m|l)', ''
    }
   
    # Extract project
    if ($Input -match 'project:(\S+)') {
        $projectKey = $Matches[1]
        $description = $description -replace 'project:\S+', ''
       
        if (-not (Get-ProjectOrTemplate $projectKey)) {
            Write-Warning "Unknown project: $projectKey - task created without project link"
            $projectKey = $null
        }
    }
   
    # Extract estimated time
    if ($Input -match 'est:(\d+\.?\d*)') {
        $estimatedTime = [double]$Matches[1]
        $description = $description -replace 'est:\d+\.?\d*', ''
    }
   
    # Extract due date
    if ($Input -match 'due:(\S+)') {
        $dueDateStr = $Matches[1]
        try {
            $dueDate = switch -Regex ($dueDateStr.ToLower()) {
                '^today$' { [datetime]::Today }
                '^tomorrow$' { [datetime]::Today.AddDays(1) }
                '^mon(day)?$' { Get-NextWeekday 1 }
                '^tue(sday)?$' { Get-NextWeekday 2 }
                '^wed(nesday)?$' { Get-NextWeekday 3 }
                '^thu(rsday)?$' { Get-NextWeekday 4 }
                '^fri(day)?$' { Get-NextWeekday 5 }
                '^sat(urday)?$' { Get-NextWeekday 6 }
                '^sun(day)?$' { Get-NextWeekday 0 }
                '^\+(\d+)$' { [datetime]::Today.AddDays([int]$Matches[1]) }
                default { [datetime]::Parse($dueDateStr) }
            }
            $dueDate = $dueDate.ToString("yyyy-MM-dd")
        }
        catch {
            Write-Warning "Invalid date format. Due date not set."
        }
        $description = $description -replace 'due:\S+', ''
    }
   
    # Clean description
    $description = $description.Trim() -replace '\s+', ' '
   
    if ([string]::IsNullOrEmpty($description)) {
        Write-Error "Task description cannot be empty!"
        return
    }
   
    # Create task
    $newTask = @{
        Id = New-TodoId
        Description = $description
        Priority = $priority
        Category = $category
        ProjectKey = $projectKey
        StartDate = $startDate
        DueDate = $dueDate
        Tags = $tags
        Progress = 0
        Completed = $false
        CreatedDate = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
        CompletedDate = $null
        EstimatedTime = $estimatedTime
        TimeSpent = 0
        Subtasks = @()
        Notes = ""
        LastModified = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
    }
   
    $script:Data.Tasks += $newTask
   
    if ($projectKey) {
        Update-ProjectStatistics -ProjectKey $projectKey
    }
   
    Save-UnifiedData
   
    Write-Success "Quick added: $description"
    if ($priority -ne $script:Data.Settings.DefaultPriority) {
        Write-Host "   Priority: $priority" -ForegroundColor Gray
    }
    if ($dueDate) {
        Write-Host "   Due: $(Format-TodoDate $dueDate)" -ForegroundColor Gray
    }
    if ($projectKey) {
        $project = Get-ProjectOrTemplate $projectKey
        Write-Host "   Project: $($project.Name)" -ForegroundColor Gray
    }
}

function Get-NextWeekday {
    param([int]$TargetDay)
   
    $today = [datetime]::Today
    $currentDay = [int]$today.DayOfWeek
    $daysToAdd = ($TargetDay - $currentDay + 7) % 7
    if ($daysToAdd -eq 0) { $daysToAdd = 7 }
   
    return $today.AddDays($daysToAdd)
}

function Complete-Task {
    param([string]$TaskId)
   
    if (-not $TaskId) {
        Show-TasksView
        $TaskId = Read-Host "`nEnter task ID to complete"
    }
   
    $task = $script:Data.Tasks | Where-Object { $_.Id -like "$TaskId*" } | Select-Object -First 1
   
    if (-not $task) {
        Write-Error "Task not found!"
        return
    }
   
    if ($task.Completed) {
        Write-Info "This task is already completed!"
        return
    }
   
    # Check for uncompleted subtasks
    if ($task.Subtasks -and ($task.Subtasks | Where-Object { -not $_.Completed }).Count -gt 0) {
        $uncompletedCount = ($task.Subtasks | Where-Object { -not $_.Completed }).Count
        Write-Warning "This task has $uncompletedCount uncompleted subtask(s)."
        Write-Host "Complete anyway? (Y/N)" -ForegroundColor Gray
        $confirm = Read-Host
        if ($confirm -ne 'Y' -and $confirm -ne 'y') {
            return
        }
    }
   
    $task.Completed = $true
    $task.Progress = 100
    $task.CompletedDate = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
    $task.LastModified = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
   
    if ($task.ProjectKey) {
        Update-ProjectStatistics -ProjectKey $task.ProjectKey
    }
   
    Save-UnifiedData
   
    Write-Success "Completed: $($task.Description)"
   
    # Show time tracking info
    if ($task.TimeSpent -gt 0) {
        Write-Host "   Time spent: $($task.TimeSpent) hours" -ForegroundColor Gray
        if ($task.EstimatedTime -gt 0) {
            $efficiency = [Math]::Round(($task.EstimatedTime / $task.TimeSpent) * 100, 0)
            Write-Host "   Efficiency: $efficiency% of estimate" -ForegroundColor Gray
        }
    }
}

function Update-TaskProgress {
    param([string]$TaskId)
   
    if (-not $TaskId) {
        Show-TasksView
        $TaskId = Read-Host "`nEnter task ID"
    }
   
    $task = $script:Data.Tasks | Where-Object { $_.Id -like "$TaskId*" } | Select-Object -First 1
   
    if (-not $task) {
        Write-Error "Task not found!"
        return
    }
   
    Write-Host "`nTask: $($task.Description)" -ForegroundColor Cyan
    Write-Host "Current progress: $($task.Progress)%"
   
    # Show progress bar
    $progressBar = "[" + ("█" * [math]::Floor($task.Progress / 10)) + ("░" * (10 - [math]::Floor($task.Progress / 10))) + "]"
    Write-Host $progressBar -ForegroundColor Green
   
    # If has subtasks, calculate progress based on them
    if ($task.Subtasks -and $task.Subtasks.Count -gt 0) {
        $completedSubtasks = ($task.Subtasks | Where-Object { $_.Completed }).Count
        $calculatedProgress = [Math]::Round(($completedSubtasks / $task.Subtasks.Count) * 100, 0)
        Write-Host "Based on subtasks: $calculatedProgress% ($completedSubtasks/$($task.Subtasks.Count) completed)" -ForegroundColor Gray
       
        Write-Host "`nUpdate based on: [S]ubtasks, [M]anual entry"
        $choice = Read-Host "Choice"
       
        if ($choice -eq 'S' -or $choice -eq 's') {
            $task.Progress = $calculatedProgress
            if ($task.Progress -eq 100) {
                $task.Completed = $true
                $task.CompletedDate = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
            }
            $task.LastModified = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
            Save-UnifiedData
            Write-Success "Progress updated to $calculatedProgress%!"
            return
        }
    }
   
    $newProgress = Read-Host "New progress (0-100)"
   
    try {
        $progress = [int]$newProgress
        if ($progress -lt 0 -or $progress -gt 100) { throw }
       
        $task.Progress = $progress
        if ($progress -eq 100) {
            $task.Completed = $true
            $task.CompletedDate = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
        }
        $task.LastModified = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
       
        Save-UnifiedData
        Write-Success "Progress updated!"
    }
    catch {
        Write-Error "Invalid progress value!"
    }
}

function Edit-Task {
    param([string]$TaskId)
   
    if (-not $TaskId) {
        Show-TasksView
        $TaskId = Read-Host "`nEnter task ID to edit"
    }
   
    $task = $script:Data.Tasks | Where-Object { $_.Id -like "$TaskId*" } | Select-Object -First 1
   
    if (-not $task) {
        Write-Error "Task not found!"
        return
    }
   
    Write-Header "Edit Task"
    Write-Host "Leave field empty to keep current value" -ForegroundColor Gray
   
    # Show current values and get updates
    Write-Host "`nCurrent: $($task.Description)"
    $newDesc = Read-Host "New description"
    if ($newDesc) { $task.Description = $newDesc }
   
    Write-Host "`nCurrent priority: $($task.Priority)"
    Write-Host "[C]ritical, [H]igh, [M]edium, [L]ow"
    $newPriority = Read-Host "New priority"
    if ($newPriority) {
        $task.Priority = switch ($newPriority.ToUpper()) {
            "C" { "Critical" }
            "H" { "High" }
            "M" { "Medium" }
            "L" { "Low" }
            default { $task.Priority }
        }
    }
   
    Write-Host "`nCurrent category: $($task.Category)"
    $newCategory = Read-Host "New category"
    if ($newCategory) { $task.Category = $newCategory }
   
    Write-Host "`nCurrent project: $(if ($task.ProjectKey) { Get-ProjectOrTemplate $task.ProjectKey | Select-Object -ExpandProperty Name } else { 'None' })"
    Write-Host "Enter new project key or 'none' to remove"
    $newProject = Read-Host "New project"
    if ($newProject -eq 'none') {
        $task.ProjectKey = $null
    } elseif ($newProject -and (Get-ProjectOrTemplate $newProject)) {
        $task.ProjectKey = $newProject
    }
   
    Write-Host "`nCurrent due date: $(if ($task.DueDate) { Format-TodoDate $task.DueDate } else { 'None' })"
    $newDueDate = Read-Host "New due date (or 'clear' to remove)"
    if ($newDueDate) {
        if ($newDueDate -eq 'clear') {
            $task.DueDate = $null
        } else {
            try {
                $dueDate = switch -Regex ($newDueDate.ToLower()) {
                    '^today$' { [datetime]::Today }
                    '^tomorrow$' { [datetime]::Today.AddDays(1) }
                    '^\+(\d+)$' { [datetime]::Today.AddDays([int]$Matches[1]) }
                    default { [datetime]::Parse($newDueDate) }
                }
                $task.DueDate = $dueDate.ToString("yyyy-MM-dd")
            }
            catch {
                Write-Warning "Invalid date format. Due date not changed."
            }
        }
    }
   
    Write-Host "`nCurrent estimated time: $($task.EstimatedTime) hours"
    $newEstimate = Read-Host "New estimate"
    if ($newEstimate) {
        try {
            $task.EstimatedTime = [double]$newEstimate
        } catch {
            Write-Warning "Invalid number format. Estimate not changed."
        }
    }
   
    Write-Host "`nCurrent tags: $($task.Tags -join ', ')"
    $newTags = Read-Host "New tags (comma-separated)"
    if ($newTags -ne $null) {
        $task.Tags = if ($newTags) {
            $newTags -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        } else { @() }
    }
   
    Write-Host "`nEdit notes? Current: $(if ($task.Notes) { 'Has notes' } else { 'No notes' }) (Y/N)"
    $editNotes = Read-Host
    if ($editNotes -eq 'Y' -or $editNotes -eq 'y') {
        Write-Host "Current notes:"
        Write-Host $task.Notes -ForegroundColor Gray
        Write-Host "New notes (empty to clear):"
        $task.Notes = Read-Host
    }
   
    $task.LastModified = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
   
    if ($task.ProjectKey) {
        Update-ProjectStatistics -ProjectKey $task.ProjectKey
    }
   
    Save-UnifiedData
   
    Write-Success "Task updated!"
}

function Manage-Subtasks {
    param([string]$TaskId)
   
    if (-not $TaskId) {
        Show-TasksView
        $TaskId = Read-Host "`nEnter task ID"
    }
   
    $task = $script:Data.Tasks | Where-Object { $_.Id -like "$TaskId*" } | Select-Object -First 1
   
    if (-not $task) {
        Write-Error "Task not found!"
        return
    }
   
    while ($true) {
        Clear-Host
        Write-Header "Manage Subtasks"
        Write-Host "Task: $($task.Description)" -ForegroundColor Yellow
        Write-Host ("=" * 50) -ForegroundColor DarkGray
       
        if ($task.Subtasks.Count -eq 0) {
            Write-Host "`nNo subtasks yet." -ForegroundColor Gray
        } else {
            Write-Host "`nSubtasks:"
            for ($i = 0; $i -lt $task.Subtasks.Count; $i++) {
                $subtask = $task.Subtasks[$i]
                $icon = if ($subtask.Completed) { "✓" } else { "○" }
                $color = if ($subtask.Completed) { "DarkGray" } else { "White" }
                Write-Host "  [$i] $icon $($subtask.Description)" -ForegroundColor $color
            }
           
            $completedCount = ($task.Subtasks | Where-Object { $_.Completed }).Count
            Write-Host "`nProgress: $completedCount/$($task.Subtasks.Count) completed" -ForegroundColor Green
        }
       
        Write-Host "`n[A]dd subtask, [C]omplete subtask, [D]elete subtask, [B]ack"
        $choice = Read-Host "Choice"
       
        switch ($choice.ToLower()) {
            "a" {
                $desc = Read-Host "Subtask description"
                if ($desc) {
                    $task.Subtasks += @{
                        Description = $desc
                        Completed = $false
                        CompletedDate = $null
                    }
                    $task.LastModified = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
                    Save-UnifiedData
                    Write-Success "Subtask added!"
                    Start-Sleep -Seconds 1
                }
            }
            "c" {
                if ($task.Subtasks.Count -eq 0) {
                    Write-Error "No subtasks to complete!"
                    Start-Sleep -Seconds 1
                    continue
                }
               
                $index = Read-Host "Subtask number"
                try {
                    $idx = [int]$index
                    if ($idx -ge 0 -and $idx -lt $task.Subtasks.Count) {
                        $task.Subtasks[$idx].Completed = -not $task.Subtasks[$idx].Completed
                        if ($task.Subtasks[$idx].Completed) {
                            $task.Subtasks[$idx].CompletedDate = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
                        } else {
                            $task.Subtasks[$idx].CompletedDate = $null
                        }
                       
                        # Update main task progress
                        $completedSubtasks = ($task.Subtasks | Where-Object { $_.Completed }).Count
                        $task.Progress = [Math]::Round(($completedSubtasks / $task.Subtasks.Count) * 100, 0)
                        $task.LastModified = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
                       
                        Save-UnifiedData
                        Write-Success "Subtask updated!"
                        Start-Sleep -Seconds 1
                    }
                } catch {
                    Write-Error "Invalid index!"
                    Start-Sleep -Seconds 1
                }
            }
            "d" {
                if ($task.Subtasks.Count -eq 0) {
                    Write-Error "No subtasks to delete!"
                    Start-Sleep -Seconds 1
                    continue
                }
               
                $index = Read-Host "Subtask number to delete"
                try {
                    $idx = [int]$index
                    if ($idx -ge 0 -and $idx -lt $task.Subtasks.Count) {
                        $task.Subtasks = @($task.Subtasks | Select-Object -Index (0..($task.Subtasks.Count-1) | Where-Object { $_ -ne $idx }))
                       
                        # Update progress
                        if ($task.Subtasks.Count -gt 0) {
                            $completedSubtasks = ($task.Subtasks | Where-Object { $_.Completed }).Count
                            $task.Progress = [Math]::Round(($completedSubtasks / $task.Subtasks.Count) * 100, 0)
                        } else {
                            # Reset progress if no subtasks
                            if ($task.Progress -eq 100 -and -not $task.Completed) {
                                $task.Progress = 0
                            }
                        }
                       
                        $task.LastModified = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
                        Save-UnifiedData
                        Write-Success "Subtask deleted!"
                        Start-Sleep -Seconds 1
                    }
                } catch {
                    Write-Error "Invalid index!"
                    Start-Sleep -Seconds 1
                }
            }
            "b" { return }
        }
    }
}

function Remove-Task {
    param([string]$TaskId)
   
    if (-not $TaskId) {
        Show-TasksView
        $TaskId = Read-Host "`nEnter task ID to delete"
    }
   
    $task = $script:Data.Tasks | Where-Object { $_.Id -like "$TaskId*" } | Select-Object -First 1
   
    if (-not $task) {
        Write-Error "Task not found!"
        return
    }
   
    Write-Warning "Delete task: '$($task.Description)'?"
    if ($task.TimeSpent -gt 0) {
        Write-Warning "This task has $($task.TimeSpent) hours logged!"
    }
    $confirm = Read-Host "Type 'yes' to confirm"
   
    if ($confirm -eq 'yes') {
        $script:Data.Tasks = $script:Data.Tasks | Where-Object { $_.Id -ne $task.Id }
       
        if ($task.ProjectKey) {
            Update-ProjectStatistics -ProjectKey $task.ProjectKey
        }
       
        Save-UnifiedData
        Write-Success "Task deleted!"
    } else {
        Write-Host "Deletion cancelled." -ForegroundColor Gray
    }
}

function Archive-CompletedTasks {
    $completed = $script:Data.Tasks | Where-Object { $_.Completed }
   
    # Auto-archive old completed tasks
    $cutoffDate = [datetime]::Today.AddDays(-$script:Data.Settings.AutoArchiveDays)
    $toArchive = $completed | Where-Object {
        [datetime]::Parse($_.CompletedDate) -lt $cutoffDate
    }
   
    if ($toArchive.Count -eq 0) {
        Write-Info "No completed tasks ready for archiving."
        Write-Host "   (Completed tasks older than $($script:Data.Settings.AutoArchiveDays) days)" -ForegroundColor Gray
       
        if ($completed.Count -gt 0) {
            Write-Host "`nArchive all $($completed.Count) completed task(s) anyway? (Y/N)" -ForegroundColor Yellow
            $confirm = Read-Host
           
            if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                $toArchive = $completed
            } else {
                return
            }
        } else {
            return
        }
    } else {
        Write-Host "Archive $($toArchive.Count) completed task(s) older than $($script:Data.Settings.AutoArchiveDays) days?" -ForegroundColor Yellow
        $confirm = Read-Host "Type 'yes' to confirm"
       
        if ($confirm -ne 'yes') {
            return
        }
    }
   
    $script:Data.ArchivedTasks += $toArchive
    $script:Data.Tasks = $script:Data.Tasks | Where-Object { $_ -notin $toArchive }
   
    # Update project statistics for affected projects
    $affectedProjects = $toArchive | Where-Object { $_.ProjectKey } | Select-Object -ExpandProperty ProjectKey -Unique
    foreach ($projectKey in $affectedProjects) {
        Update-ProjectStatistics -ProjectKey $projectKey
    }
   
    Save-UnifiedData
    Write-Success "Archived $($toArchive.Count) task(s)!"
}

function View-TaskArchive {
    Clear-Host
    Write-Header "Archived Tasks"
   
    if ($script:Data.ArchivedTasks.Count -eq 0) {
        Write-Host "`n  No archived tasks." -ForegroundColor Gray
        return
    }
   
    $grouped = $script:Data.ArchivedTasks | Group-Object {
        [datetime]::Parse($_.CompletedDate).ToString("yyyy-MM")
    } | Sort-Object Name -Descending
   
    foreach ($group in $grouped) {
        $monthYear = [datetime]::ParseExact($group.Name, "yyyy-MM", $null).ToString("MMMM yyyy")
        Write-Host "`n  📅 $monthYear ($($group.Count) items)" -ForegroundColor Yellow
       
        foreach ($task in $group.Group | Sort-Object CompletedDate -Descending) {
            Write-Host "     ✓ $($task.Description)" -ForegroundColor DarkGray
            Write-Host "       Completed: $([datetime]::Parse($task.CompletedDate).ToString('MMM dd, yyyy'))" -ForegroundColor DarkGray
           
            if ($task.TimeSpent -gt 0) {
                Write-Host "       Time: $($task.TimeSpent)h" -ForegroundColor DarkGray -NoNewline
                if ($task.EstimatedTime -gt 0) {
                    Write-Host " (Est: $($task.EstimatedTime)h)" -ForegroundColor DarkGray
                } else {
                    Write-Host
                }
            }
           
            if ($task.ProjectKey) {
                $project = Get-ProjectOrTemplate $task.ProjectKey
                if ($project) {
                    Write-Host "       Project: $($project.Name)" -ForegroundColor DarkGray
                }
            }
        }
    }
   
    # Statistics
    $totalArchived = $script:Data.ArchivedTasks.Count
    $totalTime = ($script:Data.ArchivedTasks | Measure-Object -Property TimeSpent -Sum).Sum
    $totalEstimated = ($script:Data.ArchivedTasks | Measure-Object -Property EstimatedTime -Sum).Sum
   
    Write-Host "`n" ("=" * 60) -ForegroundColor DarkGray
    Write-Host "  Total archived: $totalArchived tasks" -ForegroundColor Green
    if ($totalTime -gt 0) {
        Write-Host "  Total time spent: $totalTime hours" -ForegroundColor Green
        if ($totalEstimated -gt 0) {
            $efficiency = [Math]::Round(($totalEstimated / $totalTime) * 100, 0)
            Write-Host "  Average efficiency: $efficiency% of estimates" -ForegroundColor Green
        }
    }
   
    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Get-TaskStatus {
    param($Task)
   
    if ($Task.Completed) { return "Completed" }
    if ($Task.Progress -ge 100) { return "Done" }
    if ($Task.Progress -gt 0) { return "In Progress" }
    if ($Task.DueDate) {
        $daysUntil = ([datetime]::Parse($Task.DueDate) - [datetime]::Today).Days
        if ($daysUntil -lt 0) { return "Overdue" }
        if ($daysUntil -eq 0) { return "Due Today" }
        if ($daysUntil -le 3) { return "Due Soon" }
    }
    if ($Task.StartDate) {
        $daysUntil = ([datetime]::Parse($Task.StartDate) - [datetime]::Today).Days
        if ($daysUntil -gt 0) { return "Scheduled" }
    }
    return "Pending"
}

#endregion

#region Project Management Functions - Full Project functionality

function Add-Project {
    Write-Header "Add New Project"
   
    $key = Read-Host "Project Key (short identifier)"
    if ($script:Data.Projects.ContainsKey($key) -or $script:Data.Settings.TimeTrackerTemplates.ContainsKey($key)) {
        Write-Error "Project key already exists"
        return
    }
   
    Write-Host "`nBasic Information:" -ForegroundColor Yellow
    $name = Read-Host "Project Name"
    $id1 = Read-Host "ID1 (custom identifier)"
    $id2 = Read-Host "ID2 (max 9 chars)"
   
    Write-Host "`nClient & Department:" -ForegroundColor Yellow
    $client = Read-Host "Client Name"
    $department = Read-Host "Department"
   
    Write-Host "`nBilling Information:" -ForegroundColor Yellow
    Write-Host "Billing Type: [B]illable, [N]on-Billable, [F]ixed Price"
    $billingChoice = Read-Host "Choice (B/N/F)"
    $billingType = switch ($billingChoice.ToUpper()) {
        "B" { "Billable" }
        "F" { "Fixed Price" }
        default { "Non-Billable" }
    }
   
    $rate = 0
    $budget = 0
   
    if ($billingType -ne "Non-Billable") {
        $rateInput = Read-Host "Hourly Rate (default: $($script:Data.Settings.DefaultRate))"
        if ($rateInput) {
            $rate = [double]$rateInput
        } else {
            $rate = $script:Data.Settings.DefaultRate
        }
       
        $budgetInput = Read-Host "Budget Hours (0 for unlimited)"
        if ($budgetInput) {
            $budget = [double]$budgetInput
        }
    }
   
    Write-Host "`nProject Status:" -ForegroundColor Yellow
    Write-Host "[A]ctive, [O]n Hold, [C]ompleted"
    $statusChoice = Read-Host "Status (default: Active)"
    $status = switch ($statusChoice.ToUpper()) {
        "O" { "On Hold" }
        "C" { "Completed" }
        default { "Active" }
    }
   
    $notes = Read-Host "`nProject Notes (optional)"
   
    $startDate = (Get-Date).ToString("yyyy-MM-dd")
   
    $script:Data.Projects[$key] = @{
        Name = $name
        Id1 = $id1
        Id2 = $id2
        Client = $client
        Department = $department
        BillingType = $billingType
        Rate = $rate
        Budget = $budget
        Status = $status
        Notes = $notes
        StartDate = $startDate
        TotalHours = 0
        TotalBilled = 0
        CompletedTasks = 0
        ActiveTasks = 0
        Manager = ""
        Priority = "Medium"
        DueDate = $null
        CreatedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }
   
    Save-UnifiedData
    Write-Success "Project added: $key"
}

function Import-ProjectFromExcel {
    Write-Header "Import Project from Excel"
   
    $filePath = Read-Host "Enter Excel file path"
   
    if (-not (Test-Path $filePath)) {
        Write-Error "File not found!"
        return
    }
   
    try {
        Write-Info "Reading Excel form..."
       
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
       
        $workbook = $excel.Workbooks.Open($filePath)
       
        # Get worksheet
        $worksheetName = $script:Data.Settings.ExcelFormConfig.WorksheetName
        $worksheet = $null
       
        try {
            $worksheet = $workbook.Worksheets.Item($worksheetName)
        } catch {
            Write-Warning "Worksheet '$worksheetName' not found, using first sheet"
            $worksheet = $workbook.Worksheets.Item(1)
        }
       
        # Read project data
        $projectData = @{
            SourceFile = $filePath
            ImportedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            CreatedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            TotalHours = 0
            TotalBilled = 0
            CompletedTasks = 0
            ActiveTasks = 0
        }
       
        # Read standard fields from defined cells
        foreach ($field in $script:Data.Settings.ExcelFormConfig.StandardFields.Keys) {
            $fieldConfig = $script:Data.Settings.ExcelFormConfig.StandardFields[$field]
           
            try {
                # Read the value cell
                $value = $worksheet.Range($fieldConfig.ValueCell).Text
               
                # Optionally verify the label matches
                if ($fieldConfig.LabelCell) {
                    $label = $worksheet.Range($fieldConfig.LabelCell).Text
                    if ($label -and $label -notlike "*$($fieldConfig.Label)*") {
                        Write-Warning "Label mismatch for $field : Expected '$($fieldConfig.Label)', found '$label'"
                    }
                }
               
                # Store the value if not empty
                if ($value -and $value.Trim() -ne "") {
                    $projectData[$field] = $value.Trim()
                }
               
            } catch {
                Write-Warning "Could not read field $field : $_"
            }
        }
       
        # Close Excel
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
       
        # Show what was read
        Write-Host "`nData read from form:" -ForegroundColor $script:Data.Settings.Theme.Info
        foreach ($key in $projectData.Keys | Sort-Object) {
            if ($key -notin @('SourceFile', 'ImportedDate', 'CreatedDate', 'TotalHours', 'TotalBilled', 'CompletedTasks', 'ActiveTasks')) {
                Write-Host "  $key : $($projectData[$key])" -ForegroundColor $script:Data.Settings.Theme.Subtle
            }
        }
       
        # Determine project key
        $projectKey = if ($projectData.Id2) { $projectData.Id2 } elseif ($projectData.Id1) { $projectData.Id1 } else { Read-Host "`nProject Key (short identifier)" }
       
        # Check if exists
        if ($script:Data.Projects.ContainsKey($projectKey)) {
            Write-Warning "Project $projectKey already exists!"
            $overwrite = Read-Host "Overwrite? (Y/N)"
            if ($overwrite -ne 'Y' -and $overwrite -ne 'y') {
                return
            }
        }
       
        # Add any missing fields with defaults
        if (-not $projectData.Status) { $projectData.Status = "Active" }
        if (-not $projectData.Priority) { $projectData.Priority = "Medium" }
        if (-not $projectData.BillingType) { $projectData.BillingType = "Non-Billable" }
       
        # Save project
        $script:Data.Projects[$projectKey] = $projectData
        Save-UnifiedData
       
        Write-Success "Project imported successfully!"
        Write-Host "Project Key: $projectKey" -ForegroundColor $script:Data.Settings.Theme.Accent
        if ($projectData.Name) {
            Write-Host "Name: $($projectData.Name)" -ForegroundColor $script:Data.Settings.Theme.Accent
        }
       
    } catch {
        Write-Error "Failed to read Excel form: $_"
       
        # Cleanup on error
        if ($excel) {
            try {
                $excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            } catch {}
        }
    }
}

function Show-ProjectDetail {
    param([string]$ProjectKey)
   
    if (-not $ProjectKey) {
        Show-ProjectsAndTemplates
        $ProjectKey = Read-Host "`nEnter Project Key for details"
    }
   
    $project = Get-ProjectOrTemplate $ProjectKey
   
    if (-not $project) {
        Write-Error "Project not found!"
        return
    }
   
    # Update statistics if it's a real project
    if ($script:Data.Projects.ContainsKey($ProjectKey)) {
        Update-ProjectStatistics -ProjectKey $ProjectKey
    }
   
    Write-Header "Project Details: $($project.Name)"
   
    # Display all Excel-mapped fields with their labels
    Write-Host "`nImported Fields:" -ForegroundColor $script:Data.Settings.Theme.Accent
    foreach ($field in $script:Data.Settings.ExcelFormConfig.StandardFields.Keys) {
        $fieldConfig = $script:Data.Settings.ExcelFormConfig.StandardFields[$field]
        $fieldName = $fieldConfig.Field
       
        if ($project.$fieldName) {
            $label = $fieldConfig.Label
            $padding = " " * (20 - $label.Length)
            Write-Host "  $label :$padding" -NoNewline
           
            # Special formatting
            switch ($fieldName) {
                "Status" {
                    $color = switch ($project.Status) {
                        "Active" { "Green" }
                        "On Hold" { "Yellow" }
                        "Completed" { "DarkGray" }
                        default { "White" }
                    }
                    Write-Host $project.Status -ForegroundColor $color
                }
                "Priority" {
                    $priorityInfo = Get-PriorityInfo $project.Priority
                    Write-Host "$($priorityInfo.Icon) $($project.Priority)" -ForegroundColor $priorityInfo.Color
                }
                "BillingType" {
                    $color = if ($project.BillingType -eq "Billable") { "Green" } else { "Gray" }
                    Write-Host $project.BillingType -ForegroundColor $color
                }
                default {
                    Write-Host $project.$fieldName
                }
            }
        }
    }
   
    # Time and Budget Info
    if ($project.TotalHours -or $project.Budget) {
        Write-Host "`nTime & Budget:" -ForegroundColor $script:Data.Settings.Theme.Accent
        Write-Host "  Total Hours:         $($project.TotalHours)"
       
        if ($project.Budget -and $project.Budget -gt 0) {
            $percentUsed = if ($project.TotalHours -gt 0) { [Math]::Round(($project.TotalHours / $project.Budget) * 100, 1) } else { 0 }
            $remaining = [Math]::Round($project.Budget - $project.TotalHours, 2)
           
            Write-Host "  Budget:              $($project.Budget) hours"
            Write-Host "  Budget Used:         $percentUsed%"
            Write-Host "  Remaining:           $remaining hours"
           
            if ($percentUsed -ge 90) {
                Write-Warning "  Budget nearly exhausted!"
            } elseif ($percentUsed -ge 75) {
                Write-Warning "  75% of budget used"
            }
        }
       
        if ($project.BillingType -eq "Billable" -and $project.Rate -gt 0) {
            $totalValue = $project.TotalHours * $project.Rate
            Write-Host "  Total Value:         `$$([Math]::Round($totalValue, 2))" -ForegroundColor Green
        }
    }
   
    # Task Summary
    if ($project.ActiveTasks -gt 0 -or $project.CompletedTasks -gt 0) {
        Write-Host "`nTask Summary:" -ForegroundColor $script:Data.Settings.Theme.Accent
        Write-Host "  Active Tasks:        $($project.ActiveTasks)"
        Write-Host "  Completed Tasks:     $($project.CompletedTasks)"
       
        $totalTasks = $project.ActiveTasks + $project.CompletedTasks
        if ($totalTasks -gt 0) {
            $completionPercent = [Math]::Round(($project.CompletedTasks / $totalTasks) * 100, 1)
            Write-Host "  Completion:          $completionPercent%"
        }
    }
   
    # Recent Time Entries
    $recentEntries = $script:Data.TimeEntries |
        Where-Object { $_.ProjectKey -eq $ProjectKey } |
        Sort-Object Date -Descending |
        Select-Object -First 5
   
    if ($recentEntries.Count -gt 0) {
        Write-Host "`nRecent Time Entries:" -ForegroundColor $script:Data.Settings.Theme.Accent
        foreach ($entry in $recentEntries) {
            $date = [DateTime]::Parse($entry.Date).ToString("MMM dd")
            Write-Host "  $date : " -NoNewline
            Write-Host "$($entry.Hours)h" -ForegroundColor Cyan -NoNewline
            if ($entry.TaskId) {
                $task = $script:Data.Tasks | Where-Object { $_.Id -eq $entry.TaskId }
                if ($task) {
                    Write-Host " - $($task.Description)" -ForegroundColor Gray
                } else {
                    Write-Host ""
                }
            } elseif ($entry.Description) {
                Write-Host " - $($entry.Description)" -ForegroundColor Gray
            } else {
                Write-Host ""
            }
        }
    }
   
    # Active Tasks
    $activeTasks = $script:Data.Tasks | Where-Object {
        $_.ProjectKey -eq $ProjectKey -and -not $_.Completed
    } | Select-Object -First 5
   
    if ($activeTasks.Count -gt 0) {
        Write-Host "`nActive Tasks:" -ForegroundColor $script:Data.Settings.Theme.Accent
        foreach ($task in $activeTasks) {
            $priorityInfo = Get-PriorityInfo $task.Priority
            Write-Host "  $($priorityInfo.Icon) [$($task.Id.Substring(0,6))] $($task.Description)"
        }
    }
   
    # Metadata
    if ($project.SourceFile -or $project.ImportedDate -or $project.CreatedDate) {
        Write-Host "`nMetadata:" -ForegroundColor $script:Data.Settings.Theme.Accent
        if ($project.SourceFile) {
            Write-Host "  Source File:   $($project.SourceFile)" -ForegroundColor $script:Data.Settings.Theme.Subtle
        }
        if ($project.ImportedDate) {
            Write-Host "  Imported:      $($project.ImportedDate)" -ForegroundColor $script:Data.Settings.Theme.Subtle
        }
        if ($project.CreatedDate) {
            Write-Host "  Created:       $($project.CreatedDate)" -ForegroundColor $script:Data.Settings.Theme.Subtle
        }
    }
}

function Edit-Project {
    Show-ProjectsAndTemplates -Simple
    Write-Host ""
    $projectKey = Read-Host "Enter project key to edit"
   
    if (-not $script:Data.Projects.ContainsKey($projectKey)) {
        Write-Error "Project not found or cannot edit templates"
        return
    }
   
    $project = $script:Data.Projects[$projectKey]
   
    Write-Header "Edit Project: $projectKey"
    Write-Host "Leave field empty to keep current value" -ForegroundColor Gray
   
    # Show current values and get updates
    Write-Host "`nCurrent Name: $($project.Name)"
    $newName = Read-Host "New Name"
    if ($newName) { $project.Name = $newName }
   
    Write-Host "`nCurrent Client: $($project.Client)"
    $newClient = Read-Host "New Client"
    if ($newClient) { $project.Client = $newClient }
   
    Write-Host "`nCurrent Department: $($project.Department)"
    $newDept = Read-Host "New Department"
    if ($newDept) { $project.Department = $newDept }
   
    Write-Host "`nCurrent Status: $($project.Status)"
    Write-Host "[A]ctive, [O]n Hold, [C]ompleted"
    $newStatus = Read-Host "New Status"
    if ($newStatus) {
        $project.Status = switch ($newStatus.ToUpper()) {
            "A" { "Active" }
            "O" { "On Hold" }
            "C" { "Completed" }
            default { $project.Status }
        }
    }
   
    if ($project.BillingType -ne "Non-Billable") {
        Write-Host "`nCurrent Rate: `$$($project.Rate)"
        $newRate = Read-Host "New Rate"
        if ($newRate) { $project.Rate = [double]$newRate }
       
        Write-Host "`nCurrent Budget: $($project.Budget) hours"
        $newBudget = Read-Host "New Budget"
        if ($newBudget) { $project.Budget = [double]$newBudget }
    }
   
    Write-Host "`nCurrent Notes: $($project.Notes)"
    $newNotes = Read-Host "New Notes"
    if ($newNotes) { $project.Notes = $newNotes }
   
    Save-UnifiedData
    Write-Success "Project updated!"
}

#endregion

#region Excel Copy Functions - Full ExcelCopy.ps1 functionality

function New-ExcelCopyJob {
    Write-Header "Create Excel Copy Configuration"
   
    $jobName = Read-Host "Configuration name"
   
    if ($script:Data.ExcelCopyJobs.ContainsKey($jobName)) {
        Write-Warning "Configuration already exists!"
        $overwrite = Read-Host "Overwrite? (Y/N)"
        if ($overwrite -ne 'Y' -and $overwrite -ne 'y') { return }
    }
   
    $config = @{
        SourceWorkbook = @{
            FilePath = Read-Host "Source workbook path"
            SheetName = Read-Host "Source sheet name"
        }
        DestinationWorkbook = @{
            FilePath = Read-Host "Destination workbook path"
            SheetName = Read-Host "Destination sheet name"
            CreateIfNotExists = $true
        }
        CopyJobs = @()
        Options = @{
            MakeExcelVisible = $false
            OpenDestinationAfterCompletion = $false
        }
    }
   
    Write-Host "`nDefine copy jobs (empty source range to finish):" -ForegroundColor Gray
   
    while ($true) {
        Write-Host "`nCopy Job #$($config.CopyJobs.Count + 1)"
        $sourceRange = Read-Host "Source range (e.g., A1:A5)"
        if (-not $sourceRange) { break }
       
        $destCell = Read-Host "Destination start cell (e.g., B1)"
        $description = Read-Host "Description (optional)"
       
        $config.CopyJobs += @{
            SourceRange = $sourceRange
            DestinationStartCell = $destCell
            Description = $description
        }
    }
   
    if ($config.CopyJobs.Count -eq 0) {
        Write-Warning "No copy jobs defined!"
        return
    }
   
    Write-Host "`nMake Excel visible during operation? (Y/N)"
    $visible = Read-Host
    $config.Options.MakeExcelVisible = ($visible -eq 'Y' -or $visible -eq 'y')
   
    Write-Host "Open destination file after completion? (Y/N)"
    $open = Read-Host
    $config.Options.OpenDestinationAfterCompletion = ($open -eq 'Y' -or $open -eq 'y')
   
    $script:Data.ExcelCopyJobs[$jobName] = $config
    Save-UnifiedData
   
    Write-Success "Excel copy configuration '$jobName' created!"
   
    Write-Host "`nRun this configuration now? (Y/N)"
    $runNow = Read-Host
    if ($runNow -eq 'Y' -or $runNow -eq 'y') {
        Execute-ExcelCopyJob -JobName $jobName
    }
}

function Execute-ExcelCopyJob {
    param([string]$JobName)
   
    if (-not $JobName) {
        if ($script:Data.ExcelCopyJobs.Count -eq 0) {
            Write-Warning "No Excel copy configurations available"
            return
        }
       
        Write-Host "`nAvailable configurations:"
        foreach ($job in $script:Data.ExcelCopyJobs.Keys) {
            Write-Host "  - $job"
        }
       
        $JobName = Read-Host "`nConfiguration name"
    }
   
    $config = $script:Data.ExcelCopyJobs[$JobName]
    if (-not $config) {
        Write-Error "Configuration not found!"
        return
    }
   
    if (-not (Test-Path $config.SourceWorkbook.FilePath)) {
        Write-Error "Source file not found: $($config.SourceWorkbook.FilePath)"
        return
    }
   
    $excel = $null
    $sourceWorkbook = $null
    $sourceSheet = $null
    $destinationWorkbook = $null
    $destinationSheet = $null
   
    try {
        Write-Info "Starting Excel copy operation..."
       
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $config.Options.MakeExcelVisible
        $excel.DisplayAlerts = $false
       
        # Open source
        Write-Info "Opening source workbook..."
        $sourceWorkbook = $excel.Workbooks.Open($config.SourceWorkbook.FilePath)
       
        try {
            $sourceSheet = $sourceWorkbook.Worksheets.Item($config.SourceWorkbook.SheetName)
            Write-Success "Source sheet '$($sourceSheet.Name)' accessed."
        } catch {
            throw "Source sheet '$($config.SourceWorkbook.SheetName)' not found!"
        }
       
        # Open/create destination
        $newDestinationFile = $false
       
        if (Test-Path $config.DestinationWorkbook.FilePath) {
            Write-Info "Opening existing destination workbook..."
            $destinationWorkbook = $excel.Workbooks.Open($config.DestinationWorkbook.FilePath)
        } else {
            if ($config.DestinationWorkbook.CreateIfNotExists) {
                Write-Info "Creating new destination workbook..."
                $destinationWorkbook = $excel.Workbooks.Add()
                $newDestinationFile = $true
            } else {
                throw "Destination file does not exist!"
            }
        }
       
        # Get/create destination sheet
        try {
            $destinationSheet = $destinationWorkbook.Worksheets.Item($config.DestinationWorkbook.SheetName)
            Write-Success "Destination sheet '$($destinationSheet.Name)' accessed."
        } catch {
            Write-Info "Creating new destination sheet..."
           
            if ($destinationWorkbook.Worksheets.Count -eq 1 -and
                $destinationWorkbook.Worksheets.Item(1).Name -eq "Sheet1") {
                $destinationSheet = $destinationWorkbook.Worksheets.Item(1)
                $destinationSheet.Name = $config.DestinationWorkbook.SheetName
            } else {
                $lastSheet = $destinationWorkbook.Worksheets.Item($destinationWorkbook.Worksheets.Count)
                $destinationSheet = $destinationWorkbook.Worksheets.Add($null, $lastSheet)
                $destinationSheet.Name = $config.DestinationWorkbook.SheetName
            }
           
            Write-Success "Destination sheet created."
        }
       
        # Process copy jobs
        Write-Info "Processing $($config.CopyJobs.Count) copy job(s)..."
        $jobNumber = 0
       
        foreach ($job in $config.CopyJobs) {
            $jobNumber++
            $jobDescription = if ($job.Description) { " - $($job.Description)" } else { "" }
           
            Write-Info "Job $jobNumber/$($config.CopyJobs.Count): '$($job.SourceRange)' to '$($job.DestinationStartCell)'$jobDescription"
           
            try {
                $sourceRange = $sourceSheet.Range($job.SourceRange)
                $rowCount = $sourceRange.Rows.Count
                $colCount = $sourceRange.Columns.Count
               
                $destTopLeftCell = $destinationSheet.Range($job.DestinationStartCell)
                $destinationRange = $destTopLeftCell.Resize($rowCount, $colCount)
               
                # Copy values only
                $destinationRange.Value2 = $sourceRange.Value2
               
                Write-Success "  Copied $rowCount rows x $colCount columns"
            } catch {
                Write-Warning "  Failed to copy range: $_"
            }
        }
       
        Write-Success "Copy process completed."
       
        # Save destination
        if ($newDestinationFile) {
            $destinationFolder = Split-Path -Parent $config.DestinationWorkbook.FilePath
           
            if (-not (Test-Path $destinationFolder)) {
                New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
            }
           
            $destinationWorkbook.SaveAs($config.DestinationWorkbook.FilePath)
            Write-Success "New workbook saved."
        } else {
            $destinationWorkbook.Save()
            Write-Success "Destination workbook updated."
        }
       
        # Close workbooks
        $destinationWorkbook.Close($false)
        $sourceWorkbook.Close($false)
       
        # Open destination if requested
        if ($config.Options.OpenDestinationAfterCompletion) {
            Write-Info "Opening destination workbook..."
            Start-Process $config.DestinationWorkbook.FilePath
        }
       
    } catch {
        Write-Error "Excel copy failed: $_"
       
        # Cleanup on error
        try { if ($null -ne $sourceWorkbook) { $sourceWorkbook.Close($false) } } catch {}
        try { if ($null -ne $destinationWorkbook) { $destinationWorkbook.Close($false) } } catch {}
       
    } finally {
        # Cleanup
        if ($null -ne $excel) {
            $excel.Quit()
        }
       
        # Release COM objects
        $comObjectsToRelease = @($sourceSheet, $sourceWorkbook, $destinationSheet, $destinationWorkbook, $excel)
        foreach ($obj in $comObjectsToRelease) {
            if ($null -ne $obj) {
                try {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null
                } catch {}
            }
        }
       
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

#endregion

#region Reporting Functions

function Show-WeekReport {
    param([DateTime]$WeekStart = $script:Data.CurrentWeek)
   
    Write-Header "Week Report: $($WeekStart.ToString('yyyy-MM-dd')) to $($WeekStart.AddDays(4).ToString('yyyy-MM-dd'))"
   
    $weekDates = Get-WeekDates $WeekStart
    $weekEntries = $script:Data.TimeEntries | Where-Object {
        $entryDate = [DateTime]::Parse($_.Date)
        $entryDate -ge $weekDates[0] -and $entryDate -le $weekDates[4]
    }
   
    if ($weekEntries.Count -eq 0) {
        Write-Host "No time entries for this week" -ForegroundColor Gray
        return
    }
# Build report by project
    $projectHours = @{}
    $dayColumns = @{
        Monday = $weekDates[0]
        Tuesday = $weekDates[1]
        Wednesday = $weekDates[2]
        Thursday = $weekDates[3]
        Friday = $weekDates[4]
    }
   
    foreach ($entry in $weekEntries) {
        $entryDate = [DateTime]::Parse($entry.Date)
        $dayName = $entryDate.DayOfWeek.ToString()
       
        if (-not $projectHours.ContainsKey($entry.ProjectKey)) {
            $projectHours[$entry.ProjectKey] = @{
                Monday = 0
                Tuesday = 0
                Wednesday = 0
                Thursday = 0
                Friday = 0
            }
        }
       
        if ($dayColumns.ContainsKey($dayName)) {
            $projectHours[$entry.ProjectKey][$dayName] += $entry.Hours
        }
    }
   
    # Build output
    $output = @()
    $output += "Name`tID1`tID2`t`t`t`tMon`tTue`tWed`tThu`tFri`tTotal`tClient`tDept"
   
    $weekTotal = 0
    $billableTotal = 0
   
    foreach ($proj in $projectHours.GetEnumerator()) {
        $project = Get-ProjectOrTemplate $proj.Key
        if (-not $project) { continue }
       
        $formattedId2 = Format-Id2 $project.Id2
        $line = "$($project.Name)`t$($project.Id1)`t$formattedId2`t`t`t`t"
       
        $projectTotal = 0
        foreach ($day in @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")) {
            $hours = $proj.Value[$day]
            $line += "$hours`t"
            $projectTotal += $hours
        }
       
        $weekTotal += $projectTotal
        if ($project.BillingType -eq "Billable") {
            $billableTotal += $projectTotal
        }
       
        $line += "$projectTotal`t$($project.Client)`t$($project.Department)"
        $output += $line
    }
   
    # Display report
    Write-Host "`nTab-Delimited Output:" -ForegroundColor Yellow
    $output | ForEach-Object { Write-Host $_ }
   
    # Summary statistics
    Write-Host "`nWeek Summary:" -ForegroundColor Yellow
    Write-Host "  Total Hours:     $([Math]::Round($weekTotal, 2))"
    Write-Host "  Billable Hours:  $([Math]::Round($billableTotal, 2))"
    Write-Host "  Non-Billable:    $([Math]::Round($weekTotal - $billableTotal, 2))"
   
    if ($billableTotal -gt 0) {
        # Calculate billable value
        $billableValue = 0
        foreach ($proj in $projectHours.GetEnumerator()) {
            $project = Get-ProjectOrTemplate $proj.Key
            if ($project.BillingType -eq "Billable") {
                $projectHours = 0
                foreach ($day in @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")) {
                    $projectHours += $proj.Value[$day]
                }
                $billableValue += $projectHours * $project.Rate
            }
        }
        Write-Host "  Billable Value:  `$$([Math]::Round($billableValue, 2))" -ForegroundColor Green
    }
   
    # Tasks completed this week
    $weekCompletedTasks = $script:Data.Tasks | Where-Object {
        $_.Completed -and $_.CompletedDate -and
        [DateTime]::Parse($_.CompletedDate) -ge $weekDates[0] -and
        [DateTime]::Parse($_.CompletedDate) -le $weekDates[4]
    }
    if ($weekCompletedTasks.Count -gt 0) {
        Write-Host "  Tasks Completed: $($weekCompletedTasks.Count)" -ForegroundColor Green
    }
   
    # Copy to clipboard option
    Write-Host "`nCopy to clipboard? (Y/N): " -NoNewline -ForegroundColor Cyan
    $copy = Read-Host
    if ($copy -eq 'Y' -or $copy -eq 'y') {
        $output -join "`n" | Set-Clipboard
        Write-Success "Report copied to clipboard!"
    }
}

function Show-ExtendedReport {
    param([DateTime]$WeekStart = $script:Data.CurrentWeek)
   
    Write-Header "Extended Report: $($WeekStart.ToString('MMMM dd, yyyy'))"
   
    $weekDates = Get-WeekDates $WeekStart
    $allEntries = $script:Data.TimeEntries | Where-Object {
        $entryDate = [DateTime]::Parse($_.Date)
        $entryDate -ge $weekDates[0] -and $entryDate -le $weekDates[4]
    } | Sort-Object Date, StartTime
   
    if ($allEntries.Count -eq 0) {
        Write-Host "No entries for this week" -ForegroundColor Gray
        return
    }
   
    # Group by date
    $byDate = $allEntries | Group-Object Date
   
    foreach ($dateGroup in $byDate) {
        $date = [DateTime]::Parse($dateGroup.Name)
        Write-Host "`n$($date.ToString('dddd, MMMM dd'))" -ForegroundColor Yellow
        Write-Host ("-" * 50) -ForegroundColor DarkGray
       
        $dayTotal = 0
        foreach ($entry in $dateGroup.Group) {
            $project = Get-ProjectOrTemplate $entry.ProjectKey
           
            Write-Host "  " -NoNewline
            if ($entry.StartTime -and $entry.EndTime) {
                Write-Host "$($entry.StartTime)-$($entry.EndTime)" -ForegroundColor Gray -NoNewline
            } else {
                Write-Host "[Manual]   " -ForegroundColor DarkGray -NoNewline
            }
           
            Write-Host " $($entry.Hours)h" -ForegroundColor Cyan -NoNewline
            Write-Host " - " -NoNewline
            Write-Host "$($project.Name)" -ForegroundColor White -NoNewline
            Write-Host " ($($project.Client))" -ForegroundColor Gray -NoNewline
           
            if ($entry.TaskId) {
                $task = $script:Data.Tasks | Where-Object { $_.Id -eq $entry.TaskId }
                if ($task) {
                    Write-Host " - Task: $($task.Description)" -ForegroundColor DarkCyan
                } else {
                    Write-Host ""
                }
            } elseif ($entry.Description) {
                Write-Host " - $($entry.Description)" -ForegroundColor DarkCyan
            } else {
                Write-Host ""
            }
           
            $dayTotal += $entry.Hours
        }
       
        Write-Host "  " ("-" * 48) -ForegroundColor DarkGray
        Write-Host "  Day Total: $([Math]::Round($dayTotal, 2)) hours" -ForegroundColor Green
    }
   
    # Week summary by project
    Write-Host "`n`nWeek Summary by Project:" -ForegroundColor Yellow
    Write-Host ("-" * 50) -ForegroundColor DarkGray
   
    $byProject = $allEntries | Group-Object ProjectKey
    $grandTotal = 0
    $billableTotal = 0
    $billableValue = 0
   
    foreach ($projGroup in $byProject | Sort-Object { $_.Group[0].ProjectKey }) {
        $project = Get-ProjectOrTemplate $projGroup.Name
        $projTotal = ($projGroup.Group | Measure-Object -Property Hours -Sum).Sum
        $projTotal = [Math]::Round($projTotal, 2)
        $grandTotal += $projTotal
       
        Write-Host "  $($project.Name):" -NoNewline
        Write-Host (" " * (30 - $project.Name.Length)) -NoNewline
        Write-Host "$projTotal hours" -ForegroundColor Cyan -NoNewline
       
        if ($project.BillingType -eq "Billable") {
            $value = $projTotal * $project.Rate
            $billableTotal += $projTotal
            $billableValue += $value
            Write-Host " (`$$([Math]::Round($value, 2)))" -ForegroundColor Green
        } else {
            Write-Host " (Non-billable)" -ForegroundColor Gray
        }
       
        # Show tasks if any
        $projectTasks = $projGroup.Group | Where-Object { $_.TaskId } |
            Group-Object TaskId | Sort-Object { $_.Group[0].Date }
       
        if ($projectTasks.Count -gt 0) {
            foreach ($taskGroup in $projectTasks) {
                $task = $script:Data.Tasks | Where-Object { $_.Id -eq $taskGroup.Name }
                if ($task) {
                    $taskHours = ($taskGroup.Group | Measure-Object -Property Hours -Sum).Sum
                    Write-Host "    → $($task.Description): $([Math]::Round($taskHours, 2))h" -ForegroundColor DarkCyan
                }
            }
        }
    }
   
    Write-Host "`n" ("-" * 50) -ForegroundColor DarkGray
    Write-Host "  Total Hours:       $([Math]::Round($grandTotal, 2))" -ForegroundColor White
    Write-Host "  Billable Hours:    $([Math]::Round($billableTotal, 2))" -ForegroundColor Cyan
    Write-Host "  Non-Billable:      $([Math]::Round($grandTotal - $billableTotal, 2))" -ForegroundColor Gray
    Write-Host "  Total Value:       `$$([Math]::Round($billableValue, 2))" -ForegroundColor Green
   
    # Utilization
    $targetHours = $script:Data.Settings.HoursPerDay * $script:Data.Settings.DaysPerWeek
    $utilization = ($grandTotal / $targetHours) * 100
    Write-Host "  Utilization:       $([Math]::Round($utilization, 1))% of $targetHours target hours" -ForegroundColor Magenta
   
    # Task metrics
    $weekTasks = $script:Data.Tasks | Where-Object {
        ($_.CreatedDate -and [DateTime]::Parse($_.CreatedDate) -ge $weekDates[0] -and [DateTime]::Parse($_.CreatedDate) -le $weekDates[4]) -or
        ($_.CompletedDate -and [DateTime]::Parse($_.CompletedDate) -ge $weekDates[0] -and [DateTime]::Parse($_.CompletedDate) -le $weekDates[4])
    }
   
    if ($weekTasks.Count -gt 0) {
        Write-Host "`n  Task Activity:" -ForegroundColor Yellow
        $created = ($weekTasks | Where-Object { $_.CreatedDate -and [DateTime]::Parse($_.CreatedDate) -ge $weekDates[0] }).Count
        $completed = ($weekTasks | Where-Object { $_.CompletedDate -and [DateTime]::Parse($_.CompletedDate) -ge $weekDates[0] }).Count
        Write-Host "    Created:  $created"
        Write-Host "    Completed: $completed"
    }
}

#endregion

#region UI View Functions

function Show-ProjectsAndTemplates {
    param([switch]$Simple)
   
    if (-not $Simple) {
        Write-Header "Projects & Templates"
    }
   
    Write-Host "`nActive Projects:" -ForegroundColor Yellow
    $activeProjects = $script:Data.Projects.GetEnumerator() |
        Where-Object { $_.Value.Status -eq "Active" } |
        Sort-Object { $_.Value.Name }
   
    if ($activeProjects.Count -eq 0) {
        Write-Host "  None" -ForegroundColor Gray
    } else {
        foreach ($proj in $activeProjects) {
            Write-Host "  ● [$($proj.Key)]" -NoNewline -ForegroundColor Green
            Write-Host " $($proj.Value.Name)" -NoNewline
            Write-Host " - $($proj.Value.Client)" -ForegroundColor Gray -NoNewline
           
            if ($proj.Value.BillingType -eq "Billable") {
                Write-Host " ($" -NoNewline -ForegroundColor DarkGreen
                Write-Host "$($proj.Value.Rate)/hr" -NoNewline -ForegroundColor DarkGreen
                Write-Host ")" -ForegroundColor DarkGreen -NoNewline
            }
           
            # Show task count
            $taskCount = $proj.Value.ActiveTasks
            if ($taskCount -gt 0) {
                Write-Host " [$taskCount task$(if ($taskCount -ne 1) {'s'})]" -ForegroundColor Cyan -NoNewline
            }
            Write-Host ""
        }
    }
   
    # Other status projects
    $otherProjects = $script:Data.Projects.GetEnumerator() |
        Where-Object { $_.Value.Status -ne "Active" } |
        Sort-Object { $_.Value.Status }, { $_.Value.Name }
   
    if ($otherProjects.Count -gt 0) {
        Write-Host "`nOther Projects:" -ForegroundColor Yellow
        foreach ($proj in $otherProjects) {
            $statusIcon = switch ($proj.Value.Status) {
                "On Hold" { "◐" }
                "Completed" { "○" }
                default { "?" }
            }
            $statusColor = switch ($proj.Value.Status) {
                "On Hold" { "Yellow" }
                "Completed" { "DarkGray" }
                default { "White" }
            }
           
            Write-Host "  " -NoNewline
            Write-Host $statusIcon -ForegroundColor $statusColor -NoNewline
            Write-Host " [$($proj.Key)]" -NoNewline -ForegroundColor $statusColor
            Write-Host " $($proj.Value.Name)" -NoNewline
            Write-Host " - $($proj.Value.Client)" -ForegroundColor Gray -NoNewline
            Write-Host " ($($proj.Value.Status))" -ForegroundColor $statusColor
        }
    }
   
    Write-Host "`nTemplates:" -ForegroundColor Yellow
    foreach ($tmpl in $script:Data.Settings.TimeTrackerTemplates.GetEnumerator()) {
        Write-Host "  ● [$($tmpl.Key)]" -NoNewline -ForegroundColor Blue
        Write-Host " $($tmpl.Value.Name)" -NoNewline
        Write-Host " - Internal" -ForegroundColor Gray
    }
}

function Show-TasksView {
    param(
        [string]$Filter = "",
        [string]$SortBy = "Smart",
        [switch]$ShowCompleted,
        [string]$View = "Default"
    )
   
    # Apply filter
    $filtered = $script:Data.Tasks
    if ($Filter) {
        $filtered = $filtered | Where-Object {
            $_.Description -like "*$Filter*" -or
            $_.Category -like "*$Filter*" -or
            $_.Tags -like "*$Filter*" -or
            $_.ProjectKey -like "*$Filter*"
        }
    }
   
    # Filter completed based on settings
    if (-not $ShowCompleted) {
        $cutoffDate = [datetime]::Today.AddDays(-$script:Data.Settings.ShowCompletedDays)
        $filtered = $filtered | Where-Object {
            -not $_.Completed -or
            ([datetime]::Parse($_.CompletedDate) -ge $cutoffDate)
        }
    }
   
    # Smart sort
    $sorted = switch ($SortBy) {
        "Smart" {
            $filtered | Sort-Object @{e={
                switch(Get-TaskStatus $_) {
                    "Overdue" { 1 }
                    "Due Today" { 2 }
                    "Due Soon" { 3 }
                    "In Progress" { 4 }
                    "Pending" { 5 }
                    "Scheduled" { 6 }
                    "Completed" { 7 }
                    default { 8 }
                }
            }}, @{e={
                switch($_.Priority) {
                    "Critical" { 1 }
                    "High" { 2 }
                    "Medium" { 3 }
                    "Low" { 4 }
                    default { 5 }
                }
            }}, DueDate, CreatedDate
        }
        "Priority" {
            $filtered | Sort-Object @{e={
                switch($_.Priority) {
                    "Critical" { 1 }
                    "High" { 2 }
                    "Medium" { 3 }
                    "Low" { 4 }
                    default { 5 }
                }
            }}, DueDate, CreatedDate
        }
        "DueDate" { $filtered | Sort-Object DueDate, Priority }
        "Created" { $filtered | Sort-Object CreatedDate }
        "Category" { $filtered | Sort-Object Category, Priority }
        "Project" { $filtered | Sort-Object ProjectKey, Priority }
        default { $filtered }
    }
   
    if ($sorted.Count -eq 0) {
        Write-Host "`n  📭 No tasks found!" -ForegroundColor Yellow
        return
    }
   
    # Different view modes
    switch ($View) {
        "Kanban" { Show-KanbanView $sorted }
        "Timeline" { Show-TimelineView $sorted }
        "Project" { Show-ProjectTaskView $sorted }
        default { Show-TaskListView $sorted }
    }
   
    # Statistics
    Show-TaskStatistics $sorted
}

function Show-TaskListView {
    param($Tasks)
   
    # Group by category
    $groups = $Tasks | Group-Object Category
   
    foreach ($group in $groups) {
        $categoryName = if ([string]::IsNullOrEmpty($group.Name)) { "Uncategorized" } else { $group.Name }
        Write-Host "`n  📁 $categoryName" -ForegroundColor Magenta
        Write-Host "  " ("-" * 56) -ForegroundColor DarkGray
       
        foreach ($task in $group.Group) {
            Show-TaskItem $task
        }
    }
}

function Show-TaskItem {
    param($Task)
   
    $icon = if ($Task.Completed) { "✅" } else { "⬜" }
    $priorityInfo = Get-PriorityInfo $Task.Priority
    $id = $Task.Id.Substring(0, 6)
    $status = Get-TaskStatus $Task
   
    # Main line
    Write-Host "  $icon [$id] " -NoNewline
   
    # Priority icon
    Write-Host $priorityInfo.Icon -NoNewline
    Write-Host " " -NoNewline
   
    # Description (strikethrough if completed)
    if ($Task.Completed) {
        Write-Host $Task.Description -ForegroundColor DarkGray
    } else {
        $color = switch ($status) {
            "Overdue" { "Red" }
            "Due Today" { "Yellow" }
            "Due Soon" { "Cyan" }
            "In Progress" { "Blue" }
            default { "White" }
        }
        Write-Host $Task.Description -ForegroundColor $color
    }
   
    # Details line
    Write-Host "      " -NoNewline
   
    # Status badge
    if ($status -ne "Pending" -and $status -ne "Completed") {
        $statusColor = switch ($status) {
            "Overdue" { "Red" }
            "Due Today" { "Yellow" }
            "Due Soon" { "Cyan" }
            "In Progress" { "Blue" }
            "Scheduled" { "DarkCyan" }
            default { "Gray" }
        }
        Write-Host "[$status]" -ForegroundColor $statusColor -NoNewline
        Write-Host " " -NoNewline
    }
   
    # Due date
    if ($Task.DueDate) {
        $dueDateStr = Format-TodoDate $Task.DueDate
        $daysUntil = ([datetime]::Parse($Task.DueDate) - [datetime]::Today).Days
        $dateColor = if ($daysUntil -lt 0) { "Red" }
                    elseif ($daysUntil -eq 0) { "Yellow" }
                    elseif ($daysUntil -le 3) { "Cyan" }
                    else { "Gray" }
       
        Write-Host "📅 $dueDateStr" -ForegroundColor $dateColor -NoNewline
    }
   
    # Project
    if ($Task.ProjectKey) {
        if ($Task.DueDate -or ($status -ne "Pending" -and $status -ne "Completed")) {
            Write-Host " | " -NoNewline -ForegroundColor DarkGray
        }
        $project = Get-ProjectOrTemplate $Task.ProjectKey
        if ($project) {
            Write-Host "🏗️  $($project.Name)" -ForegroundColor Magenta -NoNewline
        }
    }
   
    # Tags
    if ($Task.Tags -and $Task.Tags.Count -gt 0) {
        if ($Task.DueDate -or $Task.ProjectKey -or ($status -ne "Pending" -and $status -ne "Completed")) {
            Write-Host " | " -NoNewline -ForegroundColor DarkGray
        }
        Write-Host "🏷️  $($Task.Tags -join ', ')" -ForegroundColor DarkCyan -NoNewline
    }
   
    # Progress
    if ($Task.Progress -gt 0 -and -not $Task.Completed) {
        if ($Task.DueDate -or $Task.Tags -or $Task.ProjectKey -or ($status -ne "Pending" -and $status -ne "Completed")) {
            Write-Host " | " -NoNewline -ForegroundColor DarkGray
        }
        $progressBar = "[" + ("█" * [math]::Floor($Task.Progress / 10)) + ("░" * (10 - [math]::Floor($Task.Progress / 10))) + "]"
        Write-Host "$progressBar $($Task.Progress)%" -ForegroundColor Green -NoNewline
    }
   
    # Time spent
    if ($Task.TimeSpent -gt 0) {
        if ($Task.DueDate -or $Task.Tags -or $Task.ProjectKey -or $Task.Progress -gt 0 -or ($status -ne "Pending" -and $status -ne "Completed")) {
            Write-Host " | " -NoNewline -ForegroundColor DarkGray
        }
        Write-Host "⏱️  $($Task.TimeSpent)h" -ForegroundColor Blue -NoNewline
    }
   
    Write-Host # New line
   
    # Subtasks if any
    if ($Task.Subtasks -and $Task.Subtasks.Count -gt 0) {
        $completedSubtasks = ($Task.Subtasks | Where-Object { $_.Completed }).Count
        Write-Host "      📌 Subtasks: $completedSubtasks/$($Task.Subtasks.Count) completed" -ForegroundColor DarkCyan
       
        foreach ($subtask in $Task.Subtasks | Select-Object -First 3) {
            $subIcon = if ($subtask.Completed) { "✓" } else { "○" }
            Write-Host "         $subIcon $($subtask.Description)" -ForegroundColor Gray
        }
       
        if ($Task.Subtasks.Count -gt 3) {
            Write-Host "         ... and $($Task.Subtasks.Count - 3) more" -ForegroundColor DarkGray
        }
    }
}

function Show-KanbanView {
    param($Tasks)
   
    $columns = @{
        "To Do" = $Tasks | Where-Object { -not $_.Completed -and $_.Progress -eq 0 }
        "In Progress" = $Tasks | Where-Object { -not $_.Completed -and $_.Progress -gt 0 -and $_.Progress -lt 100 }
        "Done" = $Tasks | Where-Object { $_.Completed -or $_.Progress -eq 100 }
    }
   
    Write-Host "`n  KANBAN BOARD" -ForegroundColor Cyan
    Write-Host "  " ("=" * 70) -ForegroundColor DarkGray
   
    # Find max items in any column
    $maxItems = ($columns.Values | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
   
    # Headers
    Write-Host "  ┌────────────────────┬────────────────────┬────────────────────┐"
    Write-Host "  │ " -NoNewline
    Write-Host "TO DO" -ForegroundColor Red -NoNewline
    Write-Host (" " * 14) -NoNewline
    Write-Host "│ " -NoNewline
    Write-Host "IN PROGRESS" -ForegroundColor Yellow -NoNewline
    Write-Host (" " * 8) -NoNewline
    Write-Host "│ " -NoNewline
    Write-Host "DONE" -ForegroundColor Green -NoNewline
    Write-Host (" " * 15) -NoNewline
    Write-Host "│"
    Write-Host "  ├────────────────────┼────────────────────┼────────────────────┤"
   
    # Items
    for ($i = 0; $i -lt $maxItems; $i++) {
        Write-Host "  │ " -NoNewline
       
        foreach ($columnName in @("To Do", "In Progress", "Done")) {
            $items = $columns[$columnName]
            if ($i -lt $items.Count) {
                $item = $items[$i]
                $text = $item.Description
                if ($text.Length -gt 16) {
                    $text = $text.Substring(0, 15) + "…"
                }
               
                $priorityInfo = Get-PriorityInfo $item.Priority
                Write-Host $priorityInfo.Icon -NoNewline
                Write-Host " $text" -NoNewline
                Write-Host (" " * (17 - $text.Length)) -NoNewline
            } else {
                Write-Host (" " * 19) -NoNewline
            }
            Write-Host "│" -NoNewline
            if ($columnName -ne "Done") { Write-Host " " -NoNewline }
        }
        Write-Host
    }
   
    Write-Host "  └────────────────────┴────────────────────┴────────────────────┘"
}

function Show-TimelineView {
    param($Tasks)
   
    Write-Host "`n  📅 TIMELINE VIEW" -ForegroundColor Cyan
    Write-Host "  " ("=" * 60) -ForegroundColor DarkGray
   
    # Group by date
    $today = [datetime]::Today
    $groups = @{
        "Overdue" = $Tasks | Where-Object { $_.DueDate -and [datetime]::Parse($_.DueDate) -lt $today -and -not $_.Completed }
        "Today" = $Tasks | Where-Object { $_.DueDate -and [datetime]::Parse($_.DueDate).Date -eq $today -and -not $_.Completed }
        "This Week" = $Tasks | Where-Object {
            $_.DueDate -and
            [datetime]::Parse($_.DueDate) -gt $today -and
            [datetime]::Parse($_.DueDate) -le $today.AddDays(7) -and
            -not $_.Completed
        }
        "Next Week" = $Tasks | Where-Object {
            $_.DueDate -and
            [datetime]::Parse($_.DueDate) -gt $today.AddDays(7) -and
            [datetime]::Parse($_.DueDate) -le $today.AddDays(14) -and
            -not $_.Completed
        }
        "Later" = $Tasks | Where-Object {
            $_.DueDate -and
            [datetime]::Parse($_.DueDate) -gt $today.AddDays(14) -and
            -not $_.Completed
        }
        "No Date" = $Tasks | Where-Object { -not $_.DueDate -and -not $_.Completed }
    }
   
    foreach ($period in @("Overdue", "Today", "This Week", "Next Week", "Later", "No Date")) {
        $items = $groups[$period]
        if ($items.Count -eq 0) { continue }
       
        $color = switch ($period) {
            "Overdue" { "Red" }
            "Today" { "Yellow" }
            "This Week" { "Cyan" }
            "Next Week" { "Blue" }
            "Later" { "DarkCyan" }
            "No Date" { "Gray" }
        }
       
        Write-Host "`n  ⏰ $period ($($items.Count))" -ForegroundColor $color
       
        foreach ($task in $items | Sort-Object DueDate, Priority) {
            $priorityInfo = Get-PriorityInfo $task.Priority
            Write-Host "     $($priorityInfo.Icon) " -NoNewline
           
            if ($task.DueDate -and $period -ne "No Date") {
                $date = [datetime]::Parse($task.DueDate)
                Write-Host "$($date.ToString('MMM dd')) - " -NoNewline -ForegroundColor Gray
            }
           
            Write-Host "$($task.Description)" -NoNewline
           
            if ($task.ProjectKey) {
                $project = Get-ProjectOrTemplate $task.ProjectKey
                if ($project) {
                    Write-Host " [$($project.Name)]" -NoNewline -ForegroundColor Magenta
                }
            }
           
            Write-Host
        }
    }
}

function Show-ProjectTaskView {
    param($Tasks)
   
    Write-Host "`n  🏗️  PROJECT VIEW" -ForegroundColor Cyan
    Write-Host "  " ("=" * 60) -ForegroundColor DarkGray
   
    # Group by project
    $groups = $Tasks | Group-Object ProjectKey | Sort-Object Name
   
    foreach ($group in $groups) {
        $projectKey = $group.Name
        $project = if ($projectKey) { Get-ProjectOrTemplate $projectKey } else { $null }
        $projectName = if ($project) { $project.Name } else { "No Project" }
       
        $active = ($group.Group | Where-Object { -not $_.Completed }).Count
        $completed = ($group.Group | Where-Object { $_.Completed }).Count
       
        Write-Host "`n  📂 $projectName " -NoNewline -ForegroundColor Magenta
        Write-Host "($active active, $completed completed)" -ForegroundColor Gray
       
        # Calculate project progress
        $totalTasks = $group.Group.Count
        $progress = if ($totalTasks -gt 0) {
            [Math]::Round(($completed / $totalTasks) * 100, 0)
        } else { 0 }
       
        $progressBar = "[" + ("█" * [math]::Floor($progress / 10)) + ("░" * (10 - [math]::Floor($progress / 10))) + "]"
        Write-Host "  $progressBar $progress%" -ForegroundColor Green
       
        # Show tasks
        foreach ($task in $group.Group | Sort-Object Completed, Priority, DueDate) {
            Show-TaskItem $task
        }
       
        # Project statistics
        $totalEstimated = ($group.Group | Measure-Object -Property EstimatedTime -Sum).Sum
        $totalSpent = ($group.Group | Measure-Object -Property TimeSpent -Sum).Sum
       
        if ($totalEstimated -gt 0 -or $totalSpent -gt 0) {
            Write-Host "  " ("-" * 40) -ForegroundColor DarkGray
            if ($totalEstimated -gt 0) {
                Write-Host "  Estimated: $totalEstimated hours" -ForegroundColor Gray
            }
            if ($totalSpent -gt 0) {
                Write-Host "  Spent: $totalSpent hours" -ForegroundColor Gray
                if ($totalEstimated -gt 0) {
                    $efficiency = [Math]::Round(($totalEstimated / $totalSpent) * 100, 0)
                    Write-Host "  Efficiency: $efficiency%" -ForegroundColor Gray
                }
            }
        }
    }
}

function Show-TaskStatistics {
    param($Tasks)
   
    $stats = @{
        Total = $Tasks.Count
        Completed = ($Tasks | Where-Object { $_.Completed }).Count
        Critical = ($Tasks | Where-Object { $_.Priority -eq "Critical" -and -not $_.Completed }).Count
        High = ($Tasks | Where-Object { $_.Priority -eq "High" -and -not $_.Completed }).Count
        Medium = ($Tasks | Where-Object { $_.Priority -eq "Medium" -and -not $_.Completed }).Count
        Low = ($Tasks | Where-Object { $_.Priority -eq "Low" -and -not $_.Completed }).Count
        Overdue = ($Tasks | Where-Object {
            $_.DueDate -and ([datetime]::Parse($_.DueDate) -lt [datetime]::Today) -and -not $_.Completed
        }).Count
        DueToday = ($Tasks | Where-Object {
            $_.DueDate -and ([datetime]::Parse($_.DueDate).Date -eq [datetime]::Today) -and -not $_.Completed
        }).Count
        InProgress = ($Tasks | Where-Object { $_.Progress -gt 0 -and $_.Progress -lt 100 -and -not $_.Completed }).Count
    }
   
    Write-Host "`n" ("=" * 70) -ForegroundColor DarkGray
    Write-Host "  📊 " -NoNewline
    Write-Host "Total: $($stats.Total)" -NoNewline
    Write-Host " | " -NoNewline -ForegroundColor DarkGray
    Write-Host "✅ Done: $($stats.Completed)" -ForegroundColor Green -NoNewline
    Write-Host " | " -NoNewline -ForegroundColor DarkGray
   
    if ($stats.Critical -gt 0) {
        Write-Host "🔥 Critical: $($stats.Critical)" -ForegroundColor Magenta -NoNewline
        Write-Host " | " -NoNewline -ForegroundColor DarkGray
    }
   
    Write-Host "🔴 High: $($stats.High)" -ForegroundColor Red -NoNewline
    Write-Host " | " -NoNewline -ForegroundColor DarkGray
    Write-Host "🟡 Med: $($stats.Medium)" -ForegroundColor Yellow -NoNewline
    Write-Host " | " -NoNewline -ForegroundColor DarkGray
    Write-Host "🟢 Low: $($stats.Low)" -ForegroundColor Green -NoNewline
   
    if ($stats.InProgress -gt 0) {
        Write-Host " | " -NoNewline -ForegroundColor DarkGray
        Write-Host "🔄 In Progress: $($stats.InProgress)" -ForegroundColor Blue -NoNewline
    }
   
    if ($stats.DueToday -gt 0) {
        Write-Host " | " -NoNewline -ForegroundColor DarkGray
        Write-Host "📅 Due Today: $($stats.DueToday)" -ForegroundColor Yellow -NoNewline
    }
   
    if ($stats.Overdue -gt 0) {
        Write-Host " | " -NoNewline -ForegroundColor DarkGray
        Write-Host "⚠️  Overdue: $($stats.Overdue)" -ForegroundColor Red -NoNewline
    }
   
    Write-Host
}

function Show-Dashboard {
    Clear-Host
    Write-Host @"
╔═══════════════════════════════════════════════════════════╗
║          UNIFIED PRODUCTIVITY SUITE v4.0                  ║
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
    Write-Host "  [M] Manual Time Entry    [S] Start Timer"
    Write-Host "  [A] Add Task            [V] View Active Timers"
    Write-Host "  [T] Today's Tasks       [W] Week Report"
    Write-Host "  [P] Projects            [H] Help"
   
    Write-Host "`n🔧 FULL MENU OPTIONS" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════════════════" -ForegroundColor DarkGray
    Write-Host "  [1] Time Management     [4] Projects & Clients"
    Write-Host "  [2] Task Management     [5] Excel Integration"
    Write-Host "  [3] Reports & Analytics [6] Settings & Config"
    Write-Host "`n  [Q] Quit"
}

#endregion

#region Main Menu Functions

function Show-TimeManagementMenu {
    while ($true) {
        Write-Header "Time Management"
       
        Write-Host "[1] Manual Time Entry (Preferred)"
        Write-Host "[2] Start Timer"
        Write-Host "[3] Stop Timer"
        Write-Host "[4] View Active Timers"
        Write-Host "[5] Quick Time Entry"
        Write-Host "[6] Edit Time Entry"
        Write-Host "[7] Delete Time Entry"
        Write-Host "[8] Today's Time Log"
        Write-Host "`n[B] Back to Dashboard"
       
        $choice = Read-Host "`nChoice"
       
        switch ($choice) {
            "1" { Add-ManualTimeEntry }
            "2" { Start-Timer }
            "3" { Stop-Timer }
            "4" {
                Show-ActiveTimers
                Write-Host "`nPress any key to continue..."
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "5" {
                Write-Host "Enter: PROJECT HOURS [DESCRIPTION]" -ForegroundColor Gray
                $input = Read-Host "Quick entry"
                Quick-TimeEntry $input
            }
            "6" { Edit-TimeEntry }
            "7" { Delete-TimeEntry }
            "8" { Show-TodayTimeLog }
            "B" { return }
            "b" { return }
        }
       
        if ($choice -ne "B" -and $choice -ne "b") {
            Write-Host "`nPress Enter to continue..."
            Read-Host
        }
    }
}

function Show-TaskManagementMenu {
    $filter = ""
    $sortBy = "Smart"
    $showCompleted = $false
    $viewMode = "Default"
   
    while ($true) {
        Clear-Host
        Write-Header "Task Management"
       
        Show-TasksView -Filter $filter -SortBy $sortBy -ShowCompleted:$showCompleted -View $viewMode
       
        Write-Host "`n" ("=" * 70) -ForegroundColor DarkGray
        Write-Host "📝 " -NoNewline
        if ($filter) { Write-Host "Filter: '$filter' | " -NoNewline -ForegroundColor Cyan }
        Write-Host "Sort: $sortBy | View: $viewMode | " -NoNewline
        if ($showCompleted) { Write-Host "Showing: All" -NoNewline } else { Write-Host "Showing: Active" -NoNewline }
        Write-Host
       
        Write-Host "`n[A]dd [C]omplete [E]dit [D]elete [P]rogress [ST]ubtasks"
        Write-Host "[F]ilter [S]ort [V]iew [T]oggle completed"
        Write-Host "a[R]chive view[AR]chive"
        Write-Host "Quick: 'qa <task>' | 'c <id>' | 'e <id>' | 'd <id>'"
        Write-Host "`n[B] Back to Dashboard"
       
        $choice = Read-Host "`nCommand"
       
        # Handle quick add
        if ($choice -like "qa *") {
            Quick-AddTask -Input $choice.Substring(3)
            continue
        }
       
        # Handle commands with ID
        if ($choice -match '^([cdep])\s+(.+)$') {
            $cmd = $Matches[1]
            $id = $Matches[2]
            switch ($cmd.ToLower()) {
                "c" { Complete-Task -TaskId $id }
                "d" { Remove-Task -TaskId $id }
                "e" { Edit-Task -TaskId $id }
                "p" { Update-TaskProgress -TaskId $id }
            }
            continue
        }
       
        switch ($choice.ToLower()) {
            "a" { Add-TodoTask }
            "c" { Complete-Task }
            "e" { Edit-Task }
            "d" { Remove-Task }
            "p" { Update-TaskProgress }
            "st" { Manage-Subtasks }
            "f" {
                $filter = Read-Host "Filter (empty to clear)"
                if ([string]::IsNullOrEmpty($filter)) { $filter = "" }
            }
            "s" {
                Write-Host "Sort by: [S]mart, [P]riority, [D]ue date, [C]reated, c[A]tegory, p[R]oject"
                $sortChoice = Read-Host "Choice"
                $sortBy = switch ($sortChoice.ToLower()) {
                    "p" { "Priority" }
                    "d" { "DueDate" }
                    "c" { "Created" }
                    "a" { "Category" }
                    "r" { "Project" }
                    default { "Smart" }
                }
            }
            "v" {
                Write-Host "View mode: [D]efault, [K]anban, [T]imeline, [P]roject"
                $viewChoice = Read-Host "Choice"
                $viewMode = switch ($viewChoice.ToLower()) {
                    "k" { "Kanban" }
                    "t" { "Timeline" }
                    "p" { "Project" }
                    default { "Default" }
                }
            }
            "t" { $showCompleted = -not $showCompleted }
            "r" { Archive-CompletedTasks }
            "ar" { View-TaskArchive }
            "b" { return }
            default {
                if (-not [string]::IsNullOrEmpty($choice)) {
                    Write-Warning "Unknown command"
                    Start-Sleep -Seconds 1
                }
            }
        }
    }
}

function Show-ReportsMenu {
    while ($true) {
        Write-Header "Reports & Analytics"
       
        Write-Host "[1] Week Report (Tab-delimited)"
        Write-Host "[2] Extended Week Report"
        Write-Host "[3] Month Summary"
        Write-Host "[4] Project Summary"
        Write-Host "[5] Task Analytics"
        Write-Host "[6] Time Analytics"
        Write-Host "[7] Export Data"
        Write-Host "[8] Change Report Week"
        Write-Host "`n[B] Back to Dashboard"
       
        Write-Host "`nCurrent week: $($script:Data.CurrentWeek.ToString('yyyy-MM-dd'))" -ForegroundColor Gray
       
        $choice = Read-Host "`nChoice"
       
        switch ($choice) {
            "1" { Show-WeekReport }
            "2" { Show-ExtendedReport }
            "3" { Show-MonthSummary }
            "4" { Show-ProjectSummary }
            "5" { Show-TaskAnalytics }
            "6" { Show-TimeAnalytics }
            "7" { Export-Data }
            "8" {
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
                        }
                    }
                }
                Save-UnifiedData
            }
            "B" { return }
            "b" { return }
        }
       
        if ($choice -ne "B" -and $choice -ne "b") {
            Write-Host "`nPress Enter to continue..."
            Read-Host
        }
    }
}

function Show-ProjectsMenu {
    while ($true) {
        Write-Header "Projects & Clients"
       
        Show-ProjectsAndTemplates
       
        Write-Host "`n[1] Add Project"
        Write-Host "[2] Import from Excel"
        Write-Host "[3] View Project Details"
        Write-Host "[4] Edit Project"
        Write-Host "[5] Configure Excel Form"
        Write-Host "[6] Batch Import Projects"
        Write-Host "[7] Export Projects"
        Write-Host "`n[B] Back to Dashboard"
       
        $choice = Read-Host "`nChoice"
       
        switch ($choice) {
            "1" { Add-Project }
            "2" { Import-ProjectFromExcel }
            "3" { Show-ProjectDetail }
            "4" { Edit-Project }
            "5" { Configure-ExcelForm }
            "6" { Batch-ImportProjects }
            "7" { Export-Projects }
            "B" { return }
            "b" { return }
        }
       
        if ($choice -ne "B" -and $choice -ne "b") {
            Write-Host "`nPress Enter to continue..."
            Read-Host
        }
    }
}

function Show-ExcelIntegrationMenu {
    while ($true) {
        Write-Header "Excel Integration"
       
        Write-Host "📊 Excel Copy Configurations:" -ForegroundColor Yellow
        if ($script:Data.ExcelCopyJobs.Count -eq 0) {
            Write-Host "  None configured" -ForegroundColor Gray
        } else {
            foreach ($job in $script:Data.ExcelCopyJobs.Keys) {
                Write-Host "  - $job"
            }
        }
       
        Write-Host "`n[1] Create Excel Copy Configuration"
        Write-Host "[2] Run Excel Copy Job"
        Write-Host "[3] Edit Copy Configuration"
        Write-Host "[4] Delete Copy Configuration"
        Write-Host "[5] Import Project from Excel"
        Write-Host "[6] Configure Excel Form Mapping"
        Write-Host "[7] Test Excel Connection"
        Write-Host "`n[B] Back to Dashboard"
       
        $choice = Read-Host "`nChoice"
       
        switch ($choice) {
            "1" { New-ExcelCopyJob }
            "2" { Execute-ExcelCopyJob }
            "3" { Edit-ExcelCopyJob }
            "4" { Remove-ExcelCopyJob }
            "5" { Import-ProjectFromExcel }
            "6" { Configure-ExcelForm }
            "7" { Test-ExcelConnection }
            "B" { return }
            "b" { return }
        }
       
        if ($choice -ne "B" -and $choice -ne "b") {
            Write-Host "`nPress Enter to continue..."
            Read-Host
        }
    }
}

function Show-SettingsMenu {
    while ($true) {
        Write-Header "Settings & Configuration"
       
        Write-Host "Current Settings:" -ForegroundColor Yellow
        Write-Host "  Default Rate:        `$$($script:Data.Settings.DefaultRate)/hour"
        Write-Host "  Hours per Day:       $($script:Data.Settings.HoursPerDay)"
        Write-Host "  Days per Week:       $($script:Data.Settings.DaysPerWeek)"
        Write-Host "  Default Priority:    $($script:Data.Settings.DefaultPriority)"
        Write-Host "  Default Category:    $($script:Data.Settings.DefaultCategory)"
        Write-Host "  Show Completed Days: $($script:Data.Settings.ShowCompletedDays)"
        Write-Host "  Auto-Archive Days:   $($script:Data.Settings.AutoArchiveDays)"
       
        Write-Host "`n[1] Time Tracking Settings"
        Write-Host "[2] Task Settings"
        Write-Host "[3] Excel Form Configuration"
        Write-Host "[4] UI Theme Settings"
        Write-Host "[5] Backup Data"
        Write-Host "[6] Restore from Backup"
        Write-Host "[7] Export All Data"
        Write-Host "[8] Import Data"
        Write-Host "[9] Reset to Defaults"
        Write-Host "`n[B] Back to Dashboard"
       
        $choice = Read-Host "`nChoice"
       
        switch ($choice) {
            "1" { Edit-TimeTrackingSettings }
            "2" { Edit-TaskSettings }
            "3" { Configure-ExcelForm }
            "4" { Edit-ThemeSettings }
            "5" {
                Backup-Data
                Write-Host "`nPress Enter to continue..."
                Read-Host
            }
            "6" { Restore-FromBackup }
            "7" { Export-AllData }
            "8" { Import-Data }
            "9" { Reset-ToDefaults }
            "B" { return }
            "b" { return }
        }
    }
}

#endregion

#region Helper Functions

function Edit-TimeEntry {
    Write-Header "Edit Time Entry"
   
    # Show recent entries
    $recentEntries = $script:Data.TimeEntries |
        Sort-Object Date -Descending |
        Select-Object -First 20
   
    Write-Host "Recent entries:"
    for ($i = 0; $i -lt $recentEntries.Count; $i++) {
        $entry = $recentEntries[$i]
        $project = Get-ProjectOrTemplate $entry.ProjectKey
        Write-Host "  [$i] $($entry.Date): $($entry.Hours)h - $($project.Name)"
        if ($entry.Description) {
            Write-Host "      $($entry.Description)" -ForegroundColor Gray
        }
    }
   
    $index = Read-Host "`nSelect entry number"
    try {
        $idx = [int]$index
        if ($idx -ge 0 -and $idx -lt $recentEntries.Count) {
            $entry = $recentEntries[$idx]
           
            Write-Host "`nLeave empty to keep current value" -ForegroundColor Gray
           
            Write-Host "Current hours: $($entry.Hours)"
            $newHours = Read-Host "New hours"
            if ($newHours) {
                $entry.Hours = [double]$newHours
            }
           
            Write-Host "Current description: $($entry.Description)"
            $newDesc = Read-Host "New description"
            if ($newDesc -ne $null) {
                $entry.Description = $newDesc
            }
           
            $entry.LastModified = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
           
            # Update project statistics
            Update-ProjectStatistics -ProjectKey $entry.ProjectKey
           
            Save-UnifiedData
            Write-Success "Entry updated!"
        }
    } catch {
        Write-Error "Invalid selection"
    }
}

function Delete-TimeEntry {
    Write-Header "Delete Time Entry"
   
    # Show recent entries
    $recentEntries = $script:Data.TimeEntries |
        Sort-Object Date -Descending |
        Select-Object -First 20
   
    Write-Host "Recent entries:"
    for ($i = 0; $i -lt $recentEntries.Count; $i++) {
        $entry = $recentEntries[$i]
        $project = Get-ProjectOrTemplate $entry.ProjectKey
        Write-Host "  [$i] $($entry.Date): $($entry.Hours)h - $($project.Name)"
        if ($entry.Description) {
            Write-Host "      $($entry.Description)" -ForegroundColor Gray
        }
    }
   
    $index = Read-Host "`nSelect entry number to delete"
    try {
        $idx = [int]$index
        if ($idx -ge 0 -and $idx -lt $recentEntries.Count) {
            $entry = $recentEntries[$idx]
           
            Write-Warning "Delete this entry? Type 'yes' to confirm"
            $confirm = Read-Host
           
            if ($confirm -eq 'yes') {
                $script:Data.TimeEntries = $script:Data.TimeEntries |
                    Where-Object { $_.Id -ne $entry.Id }
               
                # Update task time if applicable
                if ($entry.TaskId) {
                    $task = $script:Data.Tasks | Where-Object { $_.Id -eq $entry.TaskId }
                    if ($task) {
                        $task.TimeSpent = [Math]::Max(0, $task.TimeSpent - $entry.Hours)
                    }
                }
               
                # Update project statistics
                Update-ProjectStatistics -ProjectKey $entry.ProjectKey
               
                Save-UnifiedData
                Write-Success "Entry deleted!"
            }
        }
    } catch {
        Write-Error "Invalid selection"
    }
}

function Show-TodayTimeLog {
    Write-Header "Today's Time Log"
   
    $today = (Get-Date).ToString("yyyy-MM-dd")
    $todayEntries = $script:Data.TimeEntries |
        Where-Object { $_.Date -eq $today } |
        Sort-Object StartTime
   
    if ($todayEntries.Count -eq 0) {
        Write-Host "No time entries for today" -ForegroundColor Gray
        return
    }
   
    $totalHours = 0
    foreach ($entry in $todayEntries) {
        $project = Get-ProjectOrTemplate $entry.ProjectKey
       
        Write-Host "`n" -NoNewline
        if ($entry.StartTime -and $entry.EndTime) {
            Write-Host "$($entry.StartTime)-$($entry.EndTime)" -ForegroundColor Gray -NoNewline
        } else {
            Write-Host "[Manual Entry]" -ForegroundColor DarkGray -NoNewline
        }
       
        Write-Host " $($entry.Hours)h" -ForegroundColor Cyan -NoNewline
        Write-Host " - $($project.Name)" -NoNewline
       
        if ($entry.TaskId) {
            $task = $script:Data.Tasks | Where-Object { $_.Id -eq $entry.TaskId }
            if ($task) {
                Write-Host " - Task: $($task.Description)" -ForegroundColor DarkCyan
            } else {
                Write-Host ""
            }
        } elseif ($entry.Description) {
            Write-Host " - $($entry.Description)" -ForegroundColor DarkCyan
        } else {
            Write-Host ""
        }
       
        $totalHours += $entry.Hours
    }
   
    Write-Host "`n" ("-" * 50) -ForegroundColor DarkGray
    Write-Host "Total: $([Math]::Round($totalHours, 2)) hours" -ForegroundColor Green
   
    # Show vs target
    $targetToday = $script:Data.Settings.HoursPerDay
    $percent = [Math]::Round(($totalHours / $targetToday) * 100, 0)
    Write-Host "Target: $targetToday hours ($percent%)" -ForegroundColor Gray
}

function Show-MonthSummary {
    Write-Header "Month Summary"
   
    $monthStart = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0
    $monthEnd = $monthStart.AddMonths(1).AddDays(-1)
   
    Write-Host "Month: $($monthStart.ToString('MMMM yyyy'))" -ForegroundColor Yellow
   
    $monthEntries = $script:Data.TimeEntries | Where-Object {
        $date = [DateTime]::Parse($_.Date)
        $date -ge $monthStart -and $date -le $monthEnd
    }
   
    if ($monthEntries.Count -eq 0) {
        Write-Host "No entries for this month" -ForegroundColor Gray
        return
    }
   
    # Group by project
    $byProject = $monthEntries | Group-Object ProjectKey
   
    Write-Host "`nBy Project:" -ForegroundColor Yellow
    $monthTotal = 0
   
    foreach ($group in $byProject | Sort-Object { $_.Group[0].ProjectKey }) {
        $project = Get-ProjectOrTemplate $group.Name
        $hours = ($group.Group | Measure-Object -Property Hours -Sum).Sum
        $hours = [Math]::Round($hours, 2)
        $monthTotal += $hours
       
        Write-Host "  $($project.Name): $hours hours"
       
        if ($project.BillingType -eq "Billable" -and $project.Rate -gt 0) {
            $value = $hours * $project.Rate
            Write-Host "    Value: `$$([Math]::Round($value, 2))" -ForegroundColor Green
        }
    }
   
    Write-Host "`nTotal Hours: $([Math]::Round($monthTotal, 2))" -ForegroundColor Green
   
    # Task summary
    $monthTasks = $script:Data.Tasks | Where-Object {
        ($_.CreatedDate -and [DateTime]::Parse($_.CreatedDate) -ge $monthStart -and [DateTime]::Parse($_.CreatedDate) -le $monthEnd) -or
        ($_.CompletedDate -and [DateTime]::Parse($_.CompletedDate) -ge $monthStart -and [DateTime]::Parse($_.CompletedDate) -le $monthEnd)
    }
   
    if ($monthTasks.Count -gt 0) {
        Write-Host "`nTask Summary:" -ForegroundColor Yellow
        $created = ($monthTasks | Where-Object { $_.CreatedDate -and [DateTime]::Parse($_.CreatedDate) -ge $monthStart }).Count
        $completed = ($monthTasks | Where-Object { $_.CompletedDate -and [DateTime]::Parse($_.CompletedDate) -ge $monthStart }).Count
        Write-Host "  Created: $created"
        Write-Host "  Completed: $completed"
    }
}

function Show-ProjectSummary {
    Write-Header "Project Summary"
   
    foreach ($proj in $script:Data.Projects.GetEnumerator() | Sort-Object { $_.Value.Status }, { $_.Value.Name }) {
        $project = $proj.Value
        $key = $proj.Key
       
        Update-ProjectStatistics -ProjectKey $key
       
        Write-Host "`n$($project.Name) [$key]" -ForegroundColor Cyan
        Write-Host "Status: $($project.Status) | Client: $($project.Client)" -ForegroundColor Gray
       
        if ($project.TotalHours -gt 0) {
            Write-Host "Hours: $($project.TotalHours)"
           
            if ($project.Budget -gt 0) {
                $percent = [Math]::Round(($project.TotalHours / $project.Budget) * 100, 1)
                Write-Host "Budget: $percent% of $($project.Budget) hours"
            }
           
            if ($project.BillingType -eq "Billable" -and $project.Rate -gt 0) {
                $value = $project.TotalHours * $project.Rate
                Write-Host "Value: `$$([Math]::Round($value, 2))" -ForegroundColor Green
            }
        }
       
        if ($project.ActiveTasks -gt 0 -or $project.CompletedTasks -gt 0) {
            Write-Host "Tasks: $($project.ActiveTasks) active, $($project.CompletedTasks) completed"
        }
    }
}

function Show-TaskAnalytics {
    Write-Header "Task Analytics"
   
    $allTasks = $script:Data.Tasks
    $activeTasks = $allTasks | Where-Object { -not $_.Completed }
    $completedTasks = $allTasks | Where-Object { $_.Completed }
   
    Write-Host "Total Tasks: $($allTasks.Count)" -ForegroundColor Yellow
    Write-Host "  Active: $($activeTasks.Count)"
    Write-Host "  Completed: $($completedTasks.Count)"
   
    if ($completedTasks.Count -gt 0) {
        # Completion time analysis
        $completionTimes = @()
        foreach ($task in $completedTasks | Where-Object { $_.CreatedDate }) {
            $created = [DateTime]::Parse($task.CreatedDate)
            $completed = [DateTime]::Parse($task.CompletedDate)
            $days = ($completed - $created).Days
            $completionTimes += $days
        }
       
        if ($completionTimes.Count -gt 0) {
            $avgCompletion = [Math]::Round(($completionTimes | Measure-Object -Average).Average, 1)
            Write-Host "`nAverage completion time: $avgCompletion days" -ForegroundColor Green
        }
       
        # Efficiency analysis
        $withEstimates = $completedTasks | Where-Object { $_.EstimatedTime -gt 0 -and $_.TimeSpent -gt 0 }
        if ($withEstimates.Count -gt 0) {
            $totalEstimated = ($withEstimates | Measure-Object -Property EstimatedTime -Sum).Sum
            $totalSpent = ($withEstimates | Measure-Object -Property TimeSpent -Sum).Sum
            $efficiency = [Math]::Round(($totalEstimated / $totalSpent) * 100, 0)
            Write-Host "Overall efficiency: $efficiency% of estimates" -ForegroundColor Cyan
        }
    }
   
    # Priority breakdown
    Write-Host "`nBy Priority:" -ForegroundColor Yellow
    $byPriority = $allTasks | Group-Object Priority
    foreach ($group in $byPriority) {
        $active = ($group.Group | Where-Object { -not $_.Completed }).Count
        $total = $group.Count
        Write-Host "  $($group.Name): $active active / $total total"
    }
   
    # Category breakdown
    Write-Host "`nBy Category:" -ForegroundColor Yellow
    $byCategory = $allTasks | Group-Object Category | Sort-Object Count -Descending | Select-Object -First 10
    foreach ($group in $byCategory) {
        $categoryName = if ($group.Name) { $group.Name } else { "Uncategorized" }
        $active = ($group.Group | Where-Object { -not $_.Completed }).Count
        Write-Host "  $categoryName : $active active / $($group.Count) total"
    }
}

function Show-TimeAnalytics {
    Write-Header "Time Analytics"
   
    # This month
    $monthStart = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0
    $monthEntries = $script:Data.TimeEntries | Where-Object {
        [DateTime]::Parse($_.Date) -ge $monthStart
    }
   
    $monthHours = ($monthEntries | Measure-Object -Property Hours -Sum).Sum
    $monthHours = if ($monthHours) { [Math]::Round($monthHours, 2) } else { 0 }
   
    Write-Host "This Month: $monthHours hours" -ForegroundColor Yellow
   
    # Last 30 days
    $thirtyDaysAgo = (Get-Date).AddDays(-30).Date
    $last30Entries = $script:Data.TimeEntries | Where-Object {
        [DateTime]::Parse($_.Date) -ge $thirtyDaysAgo
    }
   
    $last30Hours = ($last30Entries | Measure-Object -Property Hours -Sum).Sum
    $last30Hours = if ($last30Hours) { [Math]::Round($last30Hours, 2) } else { 0 }
   
    Write-Host "Last 30 Days: $last30Hours hours" -ForegroundColor Yellow
   
    # Daily average
    $daysWithEntries = ($last30Entries | Group-Object Date).Count
    if ($daysWithEntries -gt 0) {
        $dailyAvg = [Math]::Round($last30Hours / $daysWithEntries, 2)
        Write-Host "Daily Average: $dailyAvg hours (over $daysWithEntries working days)" -ForegroundColor Green
    }
   
    # By day of week
    Write-Host "`nBy Day of Week:" -ForegroundColor Yellow
    $byDayOfWeek = $last30Entries | Group-Object { [DateTime]::Parse($_.Date).DayOfWeek } | Sort-Object Name
   
    foreach ($group in $byDayOfWeek) {
        $hours = ($group.Group | Measure-Object -Property Hours -Sum).Sum
        $hours = [Math]::Round($hours, 2)
        Write-Host "  $($group.Name): $hours hours"
    }
   
    # Top projects
    Write-Host "`nTop Projects (Last 30 Days):" -ForegroundColor Yellow
    $byProject = $last30Entries | Group-Object ProjectKey | Sort-Object { ($_.Group | Measure-Object -Property Hours -Sum).Sum } -Descending | Select-Object -First 5
   
    foreach ($group in $byProject) {
        $project = Get-ProjectOrTemplate $group.Name
        $hours = ($group.Group | Measure-Object -Property Hours -Sum).Sum
        $hours = [Math]::Round($hours, 2)
        Write-Host "  $($project.Name): $hours hours"
    }
}

function Configure-ExcelForm {
    Write-Header "Configure Excel Form Mapping"
   
    Write-Host "Current worksheet name: $($script:Data.Settings.ExcelFormConfig.WorksheetName)"
    $newName = Read-Host "New worksheet name (Enter to keep)"
    if ($newName) {
        $script:Data.Settings.ExcelFormConfig.WorksheetName = $newName
    }
   
    Write-Host "`nConfigure field mappings:" -ForegroundColor Yellow
    Write-Host "Format: LabelCell,ValueCell (e.g., A5,B5)" -ForegroundColor Gray
    Write-Host "Enter blank to skip field" -ForegroundColor Gray
   
    foreach ($field in $script:Data.Settings.ExcelFormConfig.StandardFields.Keys) {
        $current = $script:Data.Settings.ExcelFormConfig.StandardFields[$field]
        Write-Host "`n$field ($($current.Label))" -ForegroundColor Cyan
        Write-Host "Current: $($current.LabelCell),$($current.ValueCell)" -ForegroundColor Gray
       
        $input = Read-Host "New cells"
        if ($input -match '^([A-Z]+\d+),([A-Z]+\d+)$') {
            $script:Data.Settings.ExcelFormConfig.StandardFields[$field].LabelCell = $Matches[1]
            $script:Data.Settings.ExcelFormConfig.StandardFields[$field].ValueCell = $Matches[2]
        }
    }
   
    Save-UnifiedData
    Write-Success "Configuration saved!"
}

function Edit-TimeTrackingSettings {
    Write-Header "Time Tracking Settings"
   
    Write-Host "Leave empty to keep current value" -ForegroundColor Gray
   
    Write-Host "`nDefault hourly rate: $($script:Data.Settings.DefaultRate)"
    $newRate = Read-Host "New rate"
    if ($newRate) { $script:Data.Settings.DefaultRate = [double]$newRate }
   
    Write-Host "`nHours per day: $($script:Data.Settings.HoursPerDay)"
    $newHours = Read-Host "New hours"
    if ($newHours) { $script:Data.Settings.HoursPerDay = [int]$newHours }
   
    Write-Host "`nDays per week: $($script:Data.Settings.DaysPerWeek)"
    $newDays = Read-Host "New days"
    if ($newDays) { $script:Data.Settings.DaysPerWeek = [int]$newDays }
   
    Save-UnifiedData
    Write-Success "Settings updated!"
}

function Edit-TaskSettings {
    Write-Header "Task Settings"
   
    Write-Host "Leave empty to keep current value" -ForegroundColor Gray
   
    Write-Host "`nDefault priority: $($script:Data.Settings.DefaultPriority)"
    Write-Host "[C]ritical, [H]igh, [M]edium, [L]ow"
    $newPriority = Read-Host "New priority"
    if ($newPriority) {
        $script:Data.Settings.DefaultPriority = switch ($newPriority.ToUpper()) {
            "C" { "Critical" }
            "H" { "High" }
            "L" { "Low" }
            default { "Medium" }
        }
    }
   
    Write-Host "`nDefault category: $($script:Data.Settings.DefaultCategory)"
    $newCategory = Read-Host "New category"
    if ($newCategory) { $script:Data.Settings.DefaultCategory = $newCategory }
   
    Write-Host "`nShow completed tasks for days: $($script:Data.Settings.ShowCompletedDays)"
    $newDays = Read-Host "New days"
    if ($newDays) { $script:Data.Settings.ShowCompletedDays = [int]$newDays }
   
    Write-Host "`nAuto-archive after days: $($script:Data.Settings.AutoArchiveDays)"
    $newArchive = Read-Host "New days"
    if ($newArchive) { $script:Data.Settings.AutoArchiveDays = [int]$newArchive }
   
    Save-UnifiedData
    Write-Success "Settings updated!"
}

function Export-AllData {
    Write-Header "Export All Data"
   
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $exportPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "ProductivitySuite_Export_$timestamp"
    New-Item -ItemType Directory -Path $exportPath -Force | Out-Null
   
    # Export main data
    $script:Data | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $exportPath "unified_data.json")
   
    # Export time entries as CSV
    if ($script:Data.TimeEntries.Count -gt 0) {
        $timeExport = $script:Data.TimeEntries | ForEach-Object {
            $project = Get-ProjectOrTemplate $_.ProjectKey
            [PSCustomObject]@{
                Date = $_.Date
                ProjectKey = $_.ProjectKey
                ProjectName = $project.Name
                Hours = $_.Hours
                Description = $_.Description
                TaskId = $_.TaskId
                StartTime = $_.StartTime
                EndTime = $_.EndTime
            }
        }
        $timeExport | Export-Csv (Join-Path $exportPath "time_entries.csv") -NoTypeInformation
    }
   
    # Export tasks as CSV
    if ($script:Data.Tasks.Count -gt 0) {
        $taskExport = $script:Data.Tasks | ForEach-Object {
            $project = if ($_.ProjectKey) { Get-ProjectOrTemplate $_.ProjectKey } else { $null }
            [PSCustomObject]@{
                Id = $_.Id
                Description = $_.Description
                Priority = $_.Priority
                Category = $_.Category
                ProjectName = if ($project) { $project.Name } else { "" }
                Status = Get-TaskStatus $_
                DueDate = $_.DueDate
                Progress = $_.Progress
                TimeSpent = $_.TimeSpent
                EstimatedTime = $_.EstimatedTime
                Tags = $_.Tags -join ","
            }
        }
        $taskExport | Export-Csv (Join-Path $exportPath "tasks.csv") -NoTypeInformation
    }
   
    # Export projects as CSV
    if ($script:Data.Projects.Count -gt 0) {
        $projectExport = $script:Data.Projects.GetEnumerator() | ForEach-Object {
            $proj = $_.Value
            [PSCustomObject]@{
                Key = $_.Key
                Name = $proj.Name
                Id1 = $proj.Id1
                Id2 = $proj.Id2
                Client = $proj.Client
                Department = $proj.Department
                Status = $proj.Status
                BillingType = $proj.BillingType
                Rate = $proj.Rate
                Budget = $proj.Budget
                TotalHours = $proj.TotalHours
                ActiveTasks = $proj.ActiveTasks
                CompletedTasks = $proj.CompletedTasks
            }
        }
        $projectExport | Export-Csv (Join-Path $exportPath "projects.csv") -NoTypeInformation
    }
   
    Write-Success "Data exported to: $exportPath"
   
    # Open folder
    Start-Process $exportPath
}

function Show-Help {
    Clear-Host
    Write-Header "Help & Documentation"
   
    Write-Host @"
UNIFIED PRODUCTIVITY SUITE v4.0
===============================

This integrated suite combines time tracking, task management, project
management, and Excel integration into a seamless productivity system.

QUICK KEYS FROM DASHBOARD:
-------------------------
[M] Manual Time Entry - Log time with project/task selection
[S] Start Timer      - Begin tracking time in real-time
[A] Add Task         - Create new task with full options
[V] View Timers      - See all running timers
[T] Today's Tasks    - View tasks for today
[W] Week Report      - Generate weekly time report
[P] Projects         - View all projects
[H] Help            - This help screen

TIME TRACKING:
-------------
- Manual entry is the preferred method for accurate time logging
- Timers are available for real-time tracking
- Link time entries to specific tasks for detailed tracking
- Budget warnings alert you when projects approach limits
- Quick entry format: PROJECT HOURS [DESCRIPTION]

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

EXCEL INTEGRATION:
-----------------
- Configure form mappings for consistent data import
- Create reusable copy configurations
- Batch import multiple projects
- Export data in multiple formats

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

"@ -ForegroundColor White
   
    Write-Host "`nPress any key to return..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

#endregion

#region Main Program

function Start-UnifiedProductivitySuite {
    Load-UnifiedData
   
    # Quick command processing from dashboard
    while ($true) {
        Show-Dashboard
       
        Write-Host "`nCommand: " -NoNewline -ForegroundColor Yellow
        $choice = Read-Host
       
        # Process quick commands
        switch ($choice.ToUpper()) {
            "M" { Add-ManualTimeEntry }
            "S" { Start-Timer }
            "A" { Add-TodoTask }
            "V" {
                Show-ActiveTimers
                Write-Host "`nPress any key to continue..."
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "T" {
                Clear-Host
                Write-Header "Today's Tasks"
                $todayTasks = $script:Data.Tasks | Where-Object {
                    (-not $_.Completed) -and
                    ((-not $_.StartDate) -or ([DateTime]::Parse($_.StartDate) -le [DateTime]::Today)) -and
                    ((-not $_.DueDate) -or ([DateTime]::Parse($_.DueDate) -eq [DateTime]::Today))
                }
                Show-TaskListView $todayTasks
                Write-Host "`nPress any key to continue..."
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
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
            "1" { Show-TimeManagementMenu }
            "2" { Show-TaskManagementMenu }
            "3" { Show-ReportsMenu }
            "4" { Show-ProjectsMenu }
            "5" { Show-ExcelIntegrationMenu }
            "6" { Show-SettingsMenu }
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
                # Check for quick time entry
                if ($choice -match '^q\s+(.+)') {
                    Quick-TimeEntry $choice.Substring(2)
                }
                # Check for quick task add
                elseif ($choice -match '^qa\s+(.+)') {
                    Quick-AddTask -Input $choice.Substring(3)
                }
                elseif (-not [string]::IsNullOrEmpty($choice)) {
                    Write-Warning "Unknown command. Press [H] for help."
                    Start-Sleep -Seconds 1
                }
            }
        }
    }
}

# Entry point
Start-UnifiedProductivitySuite

#endregion