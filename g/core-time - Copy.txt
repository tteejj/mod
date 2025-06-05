# Core Time Management Module
# Handles time entries, timers, time-based reporting, and time-related settings.

#region Time Entry and Timers

function Add-ManualTimeEntry {
    Write-Header "Manual Time Entry"
   
    # Show projects and templates
    Show-ProjectsAndTemplates -Simple # Assuming -Simple is preferred here for brevity
   
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
            $start = [DateTime]::Parse("$($date.ToString("yyyy-MM-dd")) $($Matches[1])")
            $end = [DateTime]::Parse("$($date.ToString("yyyy-MM-dd")) $($Matches[2])")
           
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
        $projectTasks = $script:Data.Tasks | Where-Object { $_.ProjectKey -eq $projectKey -and (-not $_.Completed) -and ($_.IsCommand -ne $true) }
        if ($projectTasks.Count -gt 0) {
            Write-Host "`nActive tasks for this project:"
            foreach ($taskItem in $projectTasks) { # Renamed $task to $taskItem to avoid conflict
                Write-Host "  [$($taskItem.Id.Substring(0,6))] $($taskItem.Description)"
            }
            $taskIdInput = Read-Host "`nTask ID (partial ok)" # Renamed $taskId to $taskIdInput
           
            $matchedTask = $script:Data.Tasks | Where-Object { $_.Id -like "$taskIdInput*" -and ($_.IsCommand -ne $true) } | Select-Object -First 1
            if ($matchedTask) {
                $taskId = $matchedTask.Id
                # Update task time spent
                $matchedTask.TimeSpent = [Math]::Round($matchedTask.TimeSpent + $hours, 2)
            } else {
                Write-Warning "Task not found, proceeding without task link"
                $taskId = $null
            }
        } else {
            Write-Info "No active tasks for this project to link."
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
    if ($script:Data.Projects.ContainsKey($projectKey)) {
        Update-ProjectStatistics -ProjectKey $projectKey
    }
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
            $activeTasks = $script:Data.Tasks | Where-Object { (-not $_.Completed) -and ($_.IsCommand -ne $true) }
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
                foreach ($taskItem in $group.Group | Sort-Object Priority) { # Renamed $task to $taskItem
                    $priorityInfo = Get-PriorityInfo $taskItem.Priority
                    Write-Host "  $($priorityInfo.Icon) [$($taskItem.Id.Substring(0,6))] $($taskItem.Description)"
                }
            }
           
            $taskIdInput = Read-Host "`nTask ID (partial ok)" # Renamed $TaskId to $taskIdInput
            $task = $script:Data.Tasks | Where-Object { $_.Id -like "$taskIdInput*" -and ($_.IsCommand -ne $true) } | Select-Object -First 1
           
            if (-not $task) {
                Write-Error "Task not found"
                return
            }
           
            $TaskId = $task.Id
            $ProjectKey = $task.ProjectKey
            if (-not $ProjectKey) { # If task has no project, timer cannot be for project
                 Write-Error "Task is not linked to a project. Cannot start project-based timer for it."
                 return
            }

        } else {
            Show-ProjectsAndTemplates -Simple
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
   
    # Update project stats if it's a real project
    if ($script:Data.Projects.ContainsKey($timer.ProjectKey)) {
        Update-ProjectStatistics -ProjectKey $timer.ProjectKey
    }
   
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
    if ($script:Data.Projects.ContainsKey($projectKey)) {
        Update-ProjectStatistics -ProjectKey $projectKey
    }
    Save-UnifiedData
   
    Write-Success "Quick entry: $hours hours for $($project.Name)"
}

function Edit-TimeEntry {
    Write-Header "Edit Time Entry"
   
    # Show recent entries
    $recentEntries = $script:Data.TimeEntries |
        Sort-Object @{Expression = {[datetime]$_.Date}; Descending = $true}, @{Expression = {[datetime]$_.EnteredAt}; Descending = $true} |
        Select-Object -First 20
   
    if ($recentEntries.Count -eq 0) {
        Write-Warning "No time entries to edit."
        return
    }

    Write-Host "Recent entries:"
    for ($i = 0; $i -lt $recentEntries.Count; $i++) {
        $entry = $recentEntries[$i]
        $project = Get-ProjectOrTemplate $entry.ProjectKey
        Write-Host "  [$i] $($entry.Date): $($entry.Hours)h - $($project.Name)"
        if ($entry.Description) {
            Write-Host "      $($entry.Description)" -ForegroundColor Gray
        }
    }
   
    $indexInput = Read-Host "`nSelect entry number"
    try {
        $idx = [int]$indexInput
        if ($idx -ge 0 -and $idx -lt $recentEntries.Count) {
            $entryToEdit = $recentEntries[$idx]
            # Find the actual entry in $script:Data.TimeEntries by Id to modify it
            $originalEntry = $script:Data.TimeEntries | Where-Object { $_.Id -eq $entryToEdit.Id } | Select-Object -First 1
            
            if (-not $originalEntry) {
                Write-Error "Original entry not found in data store. This should not happen."
                return
            }

            Write-Host "`nEditing Entry ID: $($originalEntry.Id)" -ForegroundColor DarkGray
            Write-Host "Leave empty to keep current value" -ForegroundColor Gray
           
            Write-Host "Current hours: $($originalEntry.Hours)"
            $newHours = Read-Host "New hours"
            if ($newHours) {
                $originalEntry.Hours = [double]$newHours
            }
           
            Write-Host "Current description: $($originalEntry.Description)"
            $newDesc = Read-Host "New description"
            if ($newDesc -ne $null) { # Allow empty string to clear description
                $originalEntry.Description = $newDesc
            }
           
            $originalEntry.LastModified = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
           
            # Update project statistics if it's a real project
            if ($script:Data.Projects.ContainsKey($originalEntry.ProjectKey)) {
                Update-ProjectStatistics -ProjectKey $originalEntry.ProjectKey
            }
           
            Save-UnifiedData
            Write-Success "Entry updated!"
        } else {
            Write-Error "Invalid selection number."
        }
    } catch {
        Write-Error "Invalid selection input."
    }
}

function Delete-TimeEntry {
    Write-Header "Delete Time Entry"
   
    # Show recent entries
    $recentEntries = $script:Data.TimeEntries |
        Sort-Object @{Expression = {[datetime]$_.Date}; Descending = $true}, @{Expression = {[datetime]$_.EnteredAt}; Descending = $true} |
        Select-Object -First 20

    if ($recentEntries.Count -eq 0) {
        Write-Warning "No time entries to delete."
        return
    }
   
    Write-Host "Recent entries:"
    for ($i = 0; $i -lt $recentEntries.Count; $i++) {
        $entry = $recentEntries[$i]
        $project = Get-ProjectOrTemplate $entry.ProjectKey
        Write-Host "  [$i] $($entry.Date): $($entry.Hours)h - $($project.Name)"
        if ($entry.Description) {
            Write-Host "      $($entry.Description)" -ForegroundColor Gray
        }
    }
   
    $indexInput = Read-Host "`nSelect entry number to delete"
    try {
        $idx = [int]$indexInput
        if ($idx -ge 0 -and $idx -lt $recentEntries.Count) {
            $entryToDelete = $recentEntries[$idx]
           
            Write-Warning "Delete this entry? ($($entryToDelete.Date): $($entryToDelete.Hours)h for $( (Get-ProjectOrTemplate $entryToDelete.ProjectKey).Name )) Type 'yes' to confirm"
            $confirm = Read-Host
           
            if ($confirm -eq 'yes') {
                $originalHours = $entryToDelete.Hours # Store original hours before removing
                $originalTaskId = $entryToDelete.TaskId
                $originalProjectKey = $entryToDelete.ProjectKey

                $script:Data.TimeEntries = $script:Data.TimeEntries |
                    Where-Object { $_.Id -ne $entryToDelete.Id }
               
                # Update task time if applicable
                if ($originalTaskId) {
                    $task = $script:Data.Tasks | Where-Object { $_.Id -eq $originalTaskId }
                    if ($task) {
                        $task.TimeSpent = [Math]::Max(0, [Math]::Round($task.TimeSpent - $originalHours, 2))
                    }
                }
               
                # Update project statistics if it's a real project
                if ($script:Data.Projects.ContainsKey($originalProjectKey)) {
                    Update-ProjectStatistics -ProjectKey $originalProjectKey
                }
               
                Save-UnifiedData
                Write-Success "Entry deleted!"
            } else {
                Write-Info "Deletion cancelled."
            }
        } else {
             Write-Error "Invalid selection number."
        }
    } catch {
        Write-Error "Invalid selection input."
    }
}

#endregion

#region Time Reporting

function Show-TodayTimeLog {
    Write-Header "Today's Time Log"
   
    $today = (Get-Date).ToString("yyyy-MM-dd")
    $todayEntries = $script:Data.TimeEntries |
        Where-Object { $_.Date -eq $today } |
        Sort-Object StartTime # Assuming StartTime is populated for timed entries
   
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
    if ($targetToday -gt 0) {
        $percent = [Math]::Round(($totalHours / $targetToday) * 100, 0)
        Write-Host "Target: $targetToday hours ($percent%)" -ForegroundColor Gray
    } else {
        Write-Host "Target: $targetToday hours" -ForegroundColor Gray
    }
}

function Show-WeekReport {
    param([DateTime]$WeekStart = $script:Data.CurrentWeek)
   
    Write-Header "Week Report: $($WeekStart.ToString('yyyy-MM-dd')) to $($WeekStart.AddDays(4).ToString('yyyy-MM-dd'))"
   
    $weekDates = Get-WeekDates $WeekStart # Helper function expected from helper.ps1
    $weekEntries = $script:Data.TimeEntries | Where-Object {
        $entryDate = [DateTime]::Parse($_.Date)
        $entryDate -ge $weekDates[0] -and $entryDate -le $weekDates[4] # Mon-Fri
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
    $output += "Name`tID1`tID2`t`t`t`tMon`tTue`tWed`tThu`tFri`tTotal`tClient`tDept" # Original format
   
    $weekTotal = 0
    $billableTotal = 0
   
    foreach ($projEnum in $projectHours.GetEnumerator()) { # Renamed $proj to $projEnum
        $project = Get-ProjectOrTemplate $projEnum.Key
        if (-not $project) { continue }
       
        $formattedId2 = Format-Id2 $project.Id2 # Helper function expected from helper.ps1
        $line = "$($project.Name)`t$($project.Id1)`t$formattedId2`t`t`t`t" # Original format with empty tabs
       
        $projectTotalHours = 0 # Renamed $projectTotal to $projectTotalHours
        foreach ($day in @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")) {
            $hours = $projEnum.Value[$day]
            $line += "$([Math]::Round($hours,2))`t"
            $projectTotalHours += $hours
        }
       
        $weekTotal += $projectTotalHours
        if ($project.BillingType -eq "Billable") {
            $billableTotal += $projectTotalHours
        }
       
        $line += "$([Math]::Round($projectTotalHours,2))`t$($project.Client)`t$($project.Department)"
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
        foreach ($projEnum in $projectHours.GetEnumerator()) { # Renamed $proj to $projEnum
            $project = Get-ProjectOrTemplate $projEnum.Key
            if ($project.BillingType -eq "Billable" -and $project.Rate -gt 0) {
                $currentProjectHours = 0
                foreach ($day in @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")) {
                    $currentProjectHours += $projEnum.Value[$day]
                }
                $billableValue += $currentProjectHours * $project.Rate
            }
        }
        Write-Host "  Billable Value:  `$$([Math]::Round($billableValue, 2))" -ForegroundColor Green
    }
   
    # Tasks completed this week
    $weekCompletedTasks = $script:Data.Tasks | Where-Object {
        $_.Completed -and $_.CompletedDate -and
        [DateTime]::Parse($_.CompletedDate) -ge $weekDates[0] -and
        [DateTime]::Parse($_.CompletedDate) -le $weekDates[4].AddDays(1).AddSeconds(-1) # End of Friday
    }
    if ($weekCompletedTasks.Count -gt 0) {
        Write-Host "  Tasks Completed: $($weekCompletedTasks.Count)" -ForegroundColor Green
    }
   
    # Copy to clipboard option
    Write-Host "`nCopy to clipboard? (Y/N): " -NoNewline -ForegroundColor Cyan
    $copy = Read-Host
    if ($copy -eq 'Y' -or $copy -eq 'y') {
        $output -join "`r`n" | Copy-ToClipboard # Use `Copy-ToClipboard` helper
        Write-Success "Report copied to clipboard!"
    }
}

function Show-ExtendedReport {
    param([DateTime]$WeekStart = $script:Data.CurrentWeek)
   
    Write-Header "Extended Report: $($WeekStart.ToString('MMMM dd, yyyy'))"
   
    $weekDates = Get-WeekDates $WeekStart
    $allEntries = $script:Data.TimeEntries | Where-Object {
        $entryDate = [DateTime]::Parse($_.Date)
        $entryDate -ge $weekDates[0] -and $entryDate -le $weekDates[4] # Mon-Fri
    } | Sort-Object @{Expression = {[datetime]$_.Date}}, StartTime
   
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
   
    foreach ($projGroup in $byProject | Sort-Object { (Get-ProjectOrTemplate $_.Name).Name }) {
        $project = Get-ProjectOrTemplate $projGroup.Name
        $projTotalHours = ($projGroup.Group | Measure-Object -Property Hours -Sum).Sum
        $projTotalHours = [Math]::Round($projTotalHours, 2)
        $grandTotal += $projTotalHours
       
        Write-Host "  $($project.Name):" -NoNewline
        Write-Host (" " * (30 - $project.Name.Length)) -NoNewline # Adjust padding as needed
        Write-Host "$projTotalHours hours" -ForegroundColor Cyan -NoNewline
       
        if ($project.BillingType -eq "Billable" -and $project.Rate -gt 0) {
            $value = $projTotalHours * $project.Rate
            $billableTotal += $projTotalHours
            $billableValue += $value
            Write-Host " (`$$([Math]::Round($value, 2)))" -ForegroundColor Green
        } else {
            Write-Host " (Non-billable)" -ForegroundColor Gray
        }
       
        # Show tasks if any
        $projectTasks = $projGroup.Group | Where-Object { $_.TaskId } |
            Group-Object TaskId | Sort-Object { ($script:Data.Tasks | Where-Object {$_.Id -eq $_.Name} | Select -ExpandProperty Description) }
       
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
    if ($targetHours -gt 0) {
        $utilization = ($grandTotal / $targetHours) * 100
        Write-Host "  Utilization:       $([Math]::Round($utilization, 1))% of $targetHours target hours" -ForegroundColor Magenta
    }
   
    # Task metrics
    $weekTasks = $script:Data.Tasks | Where-Object {
        ($_.CreatedDate -and [DateTime]::Parse($_.CreatedDate) -ge $weekDates[0] -and [DateTime]::Parse($_.CreatedDate) -le $weekDates[4].AddDays(1).AddSeconds(-1)) -or
        ($_.CompletedDate -and [DateTime]::Parse($_.CompletedDate) -ge $weekDates[0] -and [DateTime]::Parse($_.CompletedDate) -le $weekDates[4].AddDays(1).AddSeconds(-1))
    }
   
    if ($weekTasks.Count -gt 0) {
        Write-Host "`n  Task Activity:" -ForegroundColor Yellow
        $created = ($weekTasks | Where-Object { $_.CreatedDate -and [DateTime]::Parse($_.CreatedDate) -ge $weekDates[0] -and [DateTime]::Parse($_.CreatedDate) -le $weekDates[4].AddDays(1).AddSeconds(-1) }).Count
        $completed = ($weekTasks | Where-Object { $_.CompletedDate -and [DateTime]::Parse($_.CompletedDate) -ge $weekDates[0] -and [DateTime]::Parse($_.CompletedDate) -le $weekDates[4].AddDays(1).AddSeconds(-1) }).Count
        Write-Host "    Created:  $created"
        Write-Host "    Completed: $completed"
    }
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
   
    foreach ($group in $byProject | Sort-Object { (Get-ProjectOrTemplate $_.Name).Name }) {
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
        $created = ($monthTasks | Where-Object { $_.CreatedDate -and [DateTime]::Parse($_.CreatedDate) -ge $monthStart -and [DateTime]::Parse($_.CreatedDate) -le $monthEnd }).Count
        $completed = ($monthTasks | Where-Object { $_.CompletedDate -and [DateTime]::Parse($_.CompletedDate) -ge $monthStart -and [DateTime]::Parse($_.CompletedDate) -le $monthEnd }).Count
        Write-Host "  Created: $created"
        Write-Host "  Completed: $completed"
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
   
    Write-Host "This Month ($($monthStart.ToString('MMMM yyyy'))): $monthHours hours" -ForegroundColor Yellow
   
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
        Write-Host "Daily Average: $dailyAvg hours (over $daysWithEntries working days in last 30)" -ForegroundColor Green
    }
   
    # By day of week (Last 30 days)
    Write-Host "`nBy Day of Week (Last 30 Days):" -ForegroundColor Yellow
    $byDayOfWeek = $last30Entries | Group-Object { [DateTime]::Parse($_.Date).DayOfWeek } | Sort-Object Name
   
    foreach ($group in $byDayOfWeek) {
        $hours = ($group.Group | Measure-Object -Property Hours -Sum).Sum
        $hours = [Math]::Round($hours, 2)
        Write-Host "  $([System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.GetDayName($group.Name)): $hours hours"
    }
   
    # Top projects (Last 30 Days)
    Write-Host "`nTop Projects (Last 30 Days):" -ForegroundColor Yellow
    $byProject = $last30Entries | Group-Object ProjectKey | 
        Sort-Object { ($_.Group | Measure-Object -Property Hours -Sum).Sum } -Descending | 
        Select-Object -First 5
   
    foreach ($group in $byProject) {
        $project = Get-ProjectOrTemplate $group.Name
        $hours = ($group.Group | Measure-Object -Property Hours -Sum).Sum
        $hours = [Math]::Round($hours, 2)
        Write-Host "  $($project.Name): $hours hours"
    }
}

function Export-FormattedTimesheet {
    param(
        [DateTime]$WeekStart = $script:Data.CurrentWeek,
        [string]$FilePath
    )

    Write-Header "Export Formatted Timesheet"

    if (-not $FilePath) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $FilePath = Join-Path ([Environment]::GetFolderPath("Desktop")) "FormattedTimesheet_Export_$timestamp.csv"
    }

    $weekDates = Get-WeekDates $WeekStart # Mon-Fri
    $weekEntries = $script:Data.TimeEntries | Where-Object {
        $entryDate = [DateTime]::Parse($_.Date)
        $entryDate -ge $weekDates[0] -and $entryDate -le $weekDates[4]
    }

    if ($weekEntries.Count -eq 0) {
        Write-Warning "No time entries for the week starting $($WeekStart.ToString('yyyy-MM-dd')). Export cancelled."
        return
    }

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
                Id1 = ""
                Id2 = "" # Formatted ID2
                Monday = 0
                Tuesday = 0
                Wednesday = 0
                Thursday = 0
                Friday = 0
            }
            # Populate Id1 and Id2 once per project
            $project = Get-ProjectOrTemplate $entry.ProjectKey
            if ($project) {
                $projectHours[$entry.ProjectKey].Id1 = $project.Id1
                $projectHours[$entry.ProjectKey].Id2 = Format-Id2 $project.Id2
            }
        }
        
        if ($dayColumns.ContainsKey($dayName)) {
            $projectHours[$entry.ProjectKey][$dayName] += $entry.Hours
        }
    }

    $exportData = @()
    foreach ($projKey in $projectHours.Keys) {
        $projData = $projectHours[$projKey]
        $exportData += [PSCustomObject]@{
            Id1 = $projData.Id1
            Id2 = $projData.Id2
            Monday = [Math]::Round($projData.Monday, 2)
            Tuesday = [Math]::Round($projData.Tuesday, 2)
            Wednesday = [Math]::Round($projData.Wednesday, 2)
            Thursday = [Math]::Round($projData.Thursday, 2)
            Friday = [Math]::Round($projData.Friday, 2)
        }
    }

    try {
        $exportData | Export-Csv -Path $FilePath -NoTypeInformation
        Write-Success "Formatted timesheet exported to: $FilePath"
        Write-Host "Open file? (Y/N)" -NoNewline
        $openChoice = Read-Host
        if ($openChoice -eq 'y' -or $openChoice -eq 'Y') {
            Start-Process $FilePath
        }
    } catch {
        Write-Error "Failed to export timesheet: $_"
    }
}

#endregion

#region Time Settings and Utilities

function Edit-TimeTrackingSettings {
    Write-Header "Time Tracking Settings"
   
    Write-Host "Leave empty to keep current value" -ForegroundColor Gray
   
    Write-Host "`nDefault hourly rate: $($script:Data.Settings.DefaultRate)"
    $newRate = Read-Host "New rate"
    if ($newRate) { 
        try { $script:Data.Settings.DefaultRate = [double]$newRate } 
        catch { Write-Warning "Invalid rate format. Not changed." }
    }
   
    Write-Host "`nHours per day: $($script:Data.Settings.HoursPerDay)"
    $newHours = Read-Host "New hours"
    if ($newHours) { 
        try { $script:Data.Settings.HoursPerDay = [int]$newHours }
        catch { Write-Warning "Invalid hours format. Not changed." }
    }
   
    Write-Host "`nDays per week: $($script:Data.Settings.DaysPerWeek)"
    $newDays = Read-Host "New days"
    if ($newDays) { 
        try { $script:Data.Settings.DaysPerWeek = [int]$newDays }
        catch { Write-Warning "Invalid days format. Not changed." }
    }
   
    Save-UnifiedData
    Write-Success "Settings updated!"
}

function Show-BudgetWarning {
    param([string]$ProjectKey)
   
    $project = Get-ProjectOrTemplate $ProjectKey
    # Only show for actual projects, not templates, and if budget is set
    if (-not $script:Data.Projects.ContainsKey($ProjectKey) -or -not $project -or $project.BillingType -eq "Non-Billable" -or -not $project.Budget -or $project.Budget -eq 0) {
        return
    }
   
    Update-ProjectStatistics -ProjectKey $ProjectKey # Ensure stats are fresh
    $percentUsed = ($project.TotalHours / $project.Budget) * 100
   
    if ($percentUsed -ge 100) {
        Write-Warning "BUDGET EXCEEDED for $($project.Name): $([Math]::Round($percentUsed, 1))% used. Budget: $($project.Budget)h, Used: $($project.TotalHours)h."
    } elseif ($percentUsed -ge 90) {
        Write-Warning "Budget alert for $($project.Name): $([Math]::Round($percentUsed, 1))% used ($([Math]::Round($project.Budget - $project.TotalHours, 2)) hours remaining)"
    } elseif ($percentUsed -ge 75) {
        Write-Info "Budget notice for $($project.Name): $([Math]::Round($percentUsed, 1))% used"
    }
}

#endregion