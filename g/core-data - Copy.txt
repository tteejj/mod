# Core Data Management Module
# Projects, tasks, todos, and command snippets

#region Data Model Initialization

# Initialize the unified data model
$script:Data = @{
    Projects = @{}      # Master project repository with full TimeTracker template support
    Tasks = @()         # Full TodoTracker task model with subtasks
    TimeEntries = @()   # All time entries with manual and timer support
    ActiveTimers = @{}  # Currently running timers
    ArchivedTasks = @() # TodoTracker archive
    ExcelCopyJobs = @{} # Saved Excel copy configurations
    CurrentWeek = Get-Date -Hour 0 -Minute 0 -Second 0
    Settings = Get-DefaultSettings
}

function Get-DefaultSettings {
    return @{
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
        # Command Snippets Settings
        CommandSnippets = @{
            EnableHotkeys = $true
            AutoCopyToClipboard = $true
            ShowInTaskList = $false
            DefaultCategory = "Commands"
            RecentLimit = 10
        }
        # Excel Integration Settings
        ExcelFormConfig = @{
            WorksheetName = "Project Info"
            StandardFields = @{
                "Id1" = @{ LabelCell = "A5"; ValueCell = "B5"; Label = "Project ID"; Field = "Id1" }
                "Id2" = @{ LabelCell = "A6"; ValueCell = "B6"; Label = "Task Code"; Field = "Id2" }
                "Name" = @{ LabelCell = "A7"; ValueCell = "B7"; Label = "Project Name"; Field = "Name" }
                "FullName" = @{ LabelCell = "A8"; ValueCell = "B8"; Label = "Full Description"; Field = "FullName" }
                "AssignedDate" = @{ LabelCell = "A9"; ValueCell = "B9"; Label = "Start Date"; Field = "AssignedDate" }
                "DueDate" = @{ LabelCell = "A10"; ValueCell = "B10"; Label = "End Date"; Field = "DueDate" }
                "Manager" = @{ LabelCell = "A11"; ValueCell = "B11"; Label = "Project Manager"; Field = "Manager" }
                "Budget" = @{ LabelCell = "A12"; ValueCell = "B12"; Label = "Budget"; Field = "Budget" }
                "Status" = @{ LabelCell = "A13"; ValueCell = "B13"; Label = "Status"; Field = "Status" }
                "Priority" = @{ LabelCell = "A14"; ValueCell = "B14"; Label = "Priority"; Field = "Priority" }
                "Department" = @{ LabelCell = "A15"; ValueCell = "B15"; Label = "Department"; Field = "Department" }
                "Client" = @{ LabelCell = "A16"; ValueCell = "B16"; Label = "Client"; Field = "Client" }
                "BillingType" = @{ LabelCell = "A17"; ValueCell = "B17"; Label = "Billing Type"; Field = "BillingType" }
                "Rate" = @{ LabelCell = "A18"; ValueCell = "B18"; Label = "Hourly Rate"; Field = "Rate" }
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

#region Project Management

function Get-ProjectOrTemplate {
    param([string]$Key)
    
    if ($script:Data.Projects.ContainsKey($Key)) {
        return $script:Data.Projects[$Key]
    } elseif ($script:Data.Settings.TimeTrackerTemplates.ContainsKey($Key)) {
        return $script:Data.Settings.TimeTrackerTemplates[$Key]
    }
    
    return $null
}

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

function Update-ProjectStatistics {
    param([string]$ProjectKey)
    
    $project = $script:Data.Projects[$ProjectKey]
    if (-not $project) { return }
    
    # Calculate total hours
    $projectEntries = $script:Data.TimeEntries | Where-Object { $_.ProjectKey -eq $ProjectKey }
    $project.TotalHours = [Math]::Round(($projectEntries | Measure-Object -Property Hours -Sum).Sum, 2)
    
    # Update task counts
    $projectTasks = $script:Data.Tasks | Where-Object { $_.ProjectKey -eq $ProjectKey -and $_.IsCommand -ne $true }
    $project.CompletedTasks = ($projectTasks | Where-Object { $_.Completed }).Count
    $project.ActiveTasks = ($projectTasks | Where-Object { -not $_.Completed }).Count
}

function Export-Projects {
    Write-Header "Export Projects"
    
    $exportData = @()
    foreach ($proj in $script:Data.Projects.GetEnumerator()) {
        $exportData += [PSCustomObject]@{
            Key = $proj.Key
            Name = $proj.Value.Name
            Id1 = $proj.Value.Id1
            Id2 = $proj.Value.Id2
            Client = $proj.Value.Client
            Department = $proj.Value.Department
            Status = $proj.Value.Status
            BillingType = $proj.Value.BillingType
            Rate = $proj.Value.Rate
            Budget = $proj.Value.Budget
            TotalHours = $proj.Value.TotalHours
            ActiveTasks = $proj.Value.ActiveTasks
            CompletedTasks = $proj.Value.CompletedTasks
        }
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $exportFile = Join-Path ([Environment]::GetFolderPath("Desktop")) "Projects_Export_$timestamp.csv"
    
    $exportData | Export-Csv $exportFile -NoTypeInformation
    Write-Success "Projects exported to: $exportFile"
    
    # Open file
    Start-Process $exportFile
}

function Batch-ImportProjects {
    Write-Warning "Feature not yet implemented"
}

#endregion

#region Command Snippets System

function Add-CommandSnippet {
    Write-Header "Add Command Snippet"
    
    $name = Read-Host "Command name/description"
    if ([string]::IsNullOrEmpty($name)) {
        Write-Error "Command name cannot be empty!"
        return
    }
    
    Write-Host "`nEnter command (press Enter twice to finish):" -ForegroundColor Gray
    $lines = @()
    while ($true) {
        $line = Read-Host
        if ([string]::IsNullOrEmpty($line) -and $lines.Count -gt 0) {
            break
        }
        $lines += $line
    }
    
    $command = $lines -join "`n"
    if ([string]::IsNullOrEmpty($command)) {
        Write-Error "Command cannot be empty!"
        return
    }
    
    # Category
    $existingCategories = $script:Data.Tasks | 
        Where-Object { $_.IsCommand -eq $true } | 
        Select-Object -ExpandProperty Category -Unique | 
        Where-Object { $_ }
    
    if ($existingCategories) {
        Write-Host "`nExisting categories: $($existingCategories -join ', ')" -ForegroundColor DarkCyan
    }
    $category = Read-Host "Category (default: $($script:Data.Settings.CommandSnippets.DefaultCategory))"
    if ([string]::IsNullOrEmpty($category)) {
        $category = $script:Data.Settings.CommandSnippets.DefaultCategory
    }
    
    # Tags
    Write-Host "`nTags (comma-separated, optional):" -ForegroundColor Gray
    $tagsInput = Read-Host "Tags"
    $tags = if ($tagsInput) {
        $tagsInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    } else { @() }
    
    # Hotkey
    $hotkey = ""
    if ($script:Data.Settings.CommandSnippets.EnableHotkeys) {
        Write-Host "`nAssign hotkey (optional, e.g., 'ctrl+1'):" -ForegroundColor Gray
        $hotkey = Read-Host "Hotkey"
    }
    
    # Create as special task
    $snippet = @{
        Id = New-TodoId
        Description = $name
        Priority = "Low"
        Category = $category
        ProjectKey = $null
        StartDate = $null
        DueDate = $null
        Tags = $tags
        Progress = 0
        Completed = $false
        CreatedDate = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
        CompletedDate = $null
        EstimatedTime = 0
        TimeSpent = 0
        Subtasks = @()
        Notes = $command
        LastModified = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
        IsCommand = $true
        Hotkey = $hotkey
        LastUsed = $null
        UseCount = 0
    }
    
    $script:Data.Tasks += $snippet
    Save-UnifiedData
    
    Write-Success "Command snippet added: $name"
    
    if ($script:Data.Settings.CommandSnippets.AutoCopyToClipboard) {
        if (Copy-ToClipboard $command) {
            Write-Info "Command copied to clipboard!"
        }
    }
}

function Get-CommandSnippet {
    param(
        [string]$Id,
        [string]$SearchTerm,
        [string]$Category,
        [string[]]$Tags
    )
    
    $snippets = $script:Data.Tasks | Where-Object { $_.IsCommand -eq $true }
    
    if ($Id) {
        return $snippets | Where-Object { $_.Id -like "$Id*" } | Select-Object -First 1
    }
    
    if ($SearchTerm) {
        $snippets = $snippets | Where-Object { 
            $_.Description -like "*$SearchTerm*" -or 
            $_.Notes -like "*$SearchTerm*" -or
            ($_.Tags -join " ") -like "*$SearchTerm*"
        }
    }
    
    if ($Category) {
        $snippets = $snippets | Where-Object { $_.Category -eq $Category }
    }
    
    if ($Tags) {
        $snippets = $snippets | Where-Object {
            $snippetTags = $_.Tags
            $found = $false
            foreach ($tag in $Tags) {
                if ($tag -in $snippetTags) {
                    $found = $true
                    break
                }
            }
            $found
        }
    }
    
    return $snippets
}

function Search-CommandSnippets {
    Write-Header "Search Command Snippets"
    
    $searchTerm = Read-Host "Search term (leave empty for all)"
    
    $snippets = Get-CommandSnippet -SearchTerm $searchTerm | Sort-Object UseCount -Descending
    
    if ($snippets.Count -eq 0) {
        Write-Host "No snippets found" -ForegroundColor Gray
        return
    }
    
    # Display results in table
    $tableData = $snippets | ForEach-Object {
        [PSCustomObject]@{
            ID = $_.Id.Substring(0, 6)
            Name = $_.Description
            Category = $_.Category
            Tags = ($_.Tags -join ", ")
            Used = $_.UseCount
            Hotkey = if ($_.Hotkey) { $_.Hotkey } else { "-" }
        }
    }
    
    $tableData | Format-TableUnicode -Columns @(
        @{Name="ID"; Title="ID"; Width=8}
        @{Name="Name"; Title="Name"; Width=30}
        @{Name="Category"; Title="Category"; Width=15}
        @{Name="Tags"; Title="Tags"; Width=20}
        @{Name="Used"; Title="Used"; Width=6; Align="Right"}
        @{Name="Hotkey"; Title="Hotkey"; Width=10}
    ) -Title "Command Snippets"
    
    Write-Host "`nEnter snippet ID to copy/execute, or press Enter to cancel"
    $selectedId = Read-Host
    
    if ($selectedId) {
        Execute-CommandSnippet -Id $selectedId
    }
}

function Execute-CommandSnippet {
    param([string]$Id)
    
    $snippet = Get-CommandSnippet -Id $Id
    if (-not $snippet) {
        Write-Error "Snippet not found!"
        return
    }
    
    Write-Host "`nCommand: $($snippet.Description)" -ForegroundColor Cyan
    Write-Host "Category: $($snippet.Category)" -ForegroundColor Gray
    if ($snippet.Tags.Count -gt 0) {
        Write-Host "Tags: $($snippet.Tags -join ', ')" -ForegroundColor Gray
    }
    
    Write-Host "`nCommand content:" -ForegroundColor Yellow
    Write-Host $snippet.Notes -ForegroundColor White
    
    # Update usage stats
    $snippet.LastUsed = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $snippet.UseCount++
    Save-UnifiedData
    
    Write-Host "`n[C]opy to clipboard, [E]xecute (PowerShell), [B]oth, or [Enter] to cancel"
    $action = Read-Host
    
    switch ($action.ToUpper()) {
        "C" {
            if (Copy-ToClipboard $snippet.Notes) {
                Write-Success "Command copied to clipboard!"
            }
        }
        "E" {
            Write-Warning "Execute this command?"
            $confirm = Read-Host "Type 'yes' to confirm"
            if ($confirm -eq 'yes') {
                try {
                    Invoke-Expression $snippet.Notes
                    Write-Success "Command executed!"
                } catch {
                    Write-Error "Execution failed: $_"
                }
            }
        }
        "B" {
            if (Copy-ToClipboard $snippet.Notes) {
                Write-Success "Command copied to clipboard!"
            }
            Write-Warning "Execute this command?"
            $confirm = Read-Host "Type 'yes' to confirm"
            if ($confirm -eq 'yes') {
                try {
                    Invoke-Expression $snippet.Notes
                    Write-Success "Command executed!"
                } catch {
                    Write-Error "Execution failed: $_"
                }
            }
        }
    }
}

function Remove-CommandSnippet {
    param([string]$Id)
    
    if (-not $Id) {
        Search-CommandSnippets
        $Id = Read-Host "`nEnter snippet ID to delete"
    }
    
    $snippet = Get-CommandSnippet -Id $Id
    if (-not $snippet) {
        Write-Error "Snippet not found!"
        return
    }
    
    Write-Warning "Delete snippet: '$($snippet.Description)'?"
    $confirm = Read-Host "Type 'yes' to confirm"
    
    if ($confirm -eq 'yes') {
        $script:Data.Tasks = $script:Data.Tasks | Where-Object { $_.Id -ne $snippet.Id }
        Save-UnifiedData
        Write-Success "Snippet deleted!"
    }
}

function Manage-CommandSnippets {
    while ($true) {
        Write-Header "Command Snippets"
        
        $snippetCount = ($script:Data.Tasks | Where-Object { $_.IsCommand -eq $true }).Count
        Write-Host "Total snippets: $snippetCount" -ForegroundColor Gray
        
        # Show recent snippets
        $recent = Get-RecentCommandSnippets -Count 5
        if ($recent.Count -gt 0) {
            Write-Host "`nRecent snippets:" -ForegroundColor Yellow
            foreach ($snippet in $recent) {
                Write-Host "  [$($snippet.Id.Substring(0,6))] $($snippet.Description)" -NoNewline
                if ($snippet.Hotkey) {
                    Write-Host " ($($snippet.Hotkey))" -NoNewline -ForegroundColor DarkCyan
                }
                Write-Host " - Used: $($snippet.UseCount)" -ForegroundColor Gray
            }
        }
        
        Write-Host "`n[A]dd snippet"
        Write-Host "[S]earch/Browse snippets"
        Write-Host "[E]xecute by ID"
        Write-Host "[D]elete snippet"
        Write-Host "[C]ategories"
        Write-Host "[H]otkeys"
        Write-Host "[B]ack"
        
        $choice = Read-Host "`nChoice"
        
        switch ($choice.ToUpper()) {
            "A" { Add-CommandSnippet }
            "S" { Search-CommandSnippets }
            "E" {
                $id = Read-Host "Snippet ID"
                Execute-CommandSnippet -Id $id
            }
            "D" { Remove-CommandSnippet }
            "C" { Show-SnippetCategories }
            "H" { Show-SnippetHotkeys }
            "B" { return }
        }
        
        if ($choice -ne "B" -and $choice -ne "b") {
            Write-Host "`nPress Enter to continue..."
            Read-Host
        }
    }
}

function Get-RecentCommandSnippets {
    param([int]$Count = 10)
    
    $snippets = $script:Data.Tasks | Where-Object { $_.IsCommand -eq $true }
    
    # Sort by last used, then by use count
    $sorted = $snippets | Sort-Object @{
        Expression = { if ($_.LastUsed) { [DateTime]::Parse($_.LastUsed) } else { [DateTime]::MinValue } }
        Descending = $true
    }, @{
        Expression = { $_.UseCount }
        Descending = $true
    }
    
    return $sorted | Select-Object -First $Count
}

function Show-SnippetCategories {
    Write-Header "Snippet Categories"
    
    $snippets = $script:Data.Tasks | Where-Object { $_.IsCommand -eq $true }
    $categories = $snippets | Group-Object Category | Sort-Object Count -Descending
    
    if ($categories.Count -eq 0) {
        Write-Host "No categories found" -ForegroundColor Gray
        return
    }
    
    Write-Host "Category usage:" -ForegroundColor Yellow
    foreach ($cat in $categories) {
        Write-Host "  $($cat.Name): $($cat.Count) snippet(s)"
        
        # Show top snippets in category
        $topInCategory = $cat.Group | Sort-Object UseCount -Descending | Select-Object -First 3
        foreach ($snippet in $topInCategory) {
            Write-Host "    - $($snippet.Description)" -ForegroundColor Gray
        }
    }
}

function Show-SnippetHotkeys {
    Write-Header "Snippet Hotkeys"
    
    $snippetsWithHotkeys = $script:Data.Tasks | Where-Object { $_.IsCommand -eq $true -and $_.Hotkey }
    
    if ($snippetsWithHotkeys.Count -eq 0) {
        Write-Host "No hotkeys assigned" -ForegroundColor Gray
        return
    }
    
    Write-Host "Assigned hotkeys:" -ForegroundColor Yellow
    foreach ($snippet in $snippetsWithHotkeys | Sort-Object Hotkey) {
        Write-Host "  $($snippet.Hotkey): $($snippet.Description)"
    }
    
    Write-Warning "`nNote: Hotkey functionality requires external keyboard hook implementation"
}

function Edit-CommandSnippetSettings {
    Write-Header "Command Snippet Settings"
    
    Write-Host "Current settings:" -ForegroundColor Yellow
    Write-Host "  Enable Hotkeys:       $(if ($script:Data.Settings.CommandSnippets.EnableHotkeys) { 'Yes' } else { 'No' })"
    Write-Host "  Auto-Copy:           $(if ($script:Data.Settings.CommandSnippets.AutoCopyToClipboard) { 'Yes' } else { 'No' })"
    Write-Host "  Show in Task List:   $(if ($script:Data.Settings.CommandSnippets.ShowInTaskList) { 'Yes' } else { 'No' })"
    Write-Host "  Default Category:    $($script:Data.Settings.CommandSnippets.DefaultCategory)"
    Write-Host "  Recent Limit:        $($script:Data.Settings.CommandSnippets.RecentLimit)"
    
    Write-Host "`nLeave empty to keep current value" -ForegroundColor Gray
    
    Write-Host "`nEnable hotkeys? (Y/N)"
    $hotkeys = Read-Host
    if ($hotkeys) {
        $script:Data.Settings.CommandSnippets.EnableHotkeys = ($hotkeys -eq 'Y' -or $hotkeys -eq 'y')
    }
    
    Write-Host "`nAuto-copy to clipboard? (Y/N)"
    $autoCopy = Read-Host
    if ($autoCopy) {
        $script:Data.Settings.CommandSnippets.AutoCopyToClipboard = ($autoCopy -eq 'Y' -or $autoCopy -eq 'y')
    }
    
    Write-Host "`nShow command snippets in task list? (Y/N)"
    $showInTasks = Read-Host
    if ($showInTasks) {
        $script:Data.Settings.CommandSnippets.ShowInTaskList = ($showInTasks -eq 'Y' -or $showInTasks -eq 'y')
    }
    
    Write-Host "`nDefault category: $($script:Data.Settings.CommandSnippets.DefaultCategory)"
    $newCategory = Read-Host "New default category"
    if ($newCategory) {
        $script:Data.Settings.CommandSnippets.DefaultCategory = $newCategory
    }
    
    Write-Host "`nRecent snippets limit: $($script:Data.Settings.CommandSnippets.RecentLimit)"
    $newLimit = Read-Host "New limit"
    if ($newLimit) {
        try {
            $script:Data.Settings.CommandSnippets.RecentLimit = [int]$newLimit
        } catch {
            Write-Warning "Invalid number format"
        }
    }
    
    Save-UnifiedData
    Write-Success "Settings updated!"
}

#endregion

#region Task Management Functions

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
    $existingCategories = $script:Data.Tasks | 
        Where-Object { $_.IsCommand -ne $true } |
        Select-Object -ExpandProperty Category -Unique | 
        Where-Object { $_ }
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

function Complete-Task {
    param([string]$TaskId)
    
    if (-not $TaskId) {
        Show-TasksView
        $TaskId = Read-Host "`nEnter task ID to complete"
    }
    
    $task = $script:Data.Tasks | Where-Object { $_.Id -like "$TaskId*" -and $_.IsCommand -ne $true } | Select-Object -First 1
    
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
    
    $task = $script:Data.Tasks | Where-Object { $_.Id -like "$TaskId*" -and $_.IsCommand -ne $true } | Select-Object -First 1
    
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
    
    $task = $script:Data.Tasks | Where-Object { $_.Id -like "$TaskId*" -and $_.IsCommand -ne $true } | Select-Object -First 1
    
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
    
    $task = $script:Data.Tasks | Where-Object { $_.Id -like "$TaskId*" -and $_.IsCommand -ne $true } | Select-Object -First 1
    
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
    
    $task = $script:Data.Tasks | Where-Object { $_.Id -like "$TaskId*" -and $_.IsCommand -ne $true } | Select-Object -First 1
    
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
    $completed = $script:Data.Tasks | Where-Object { $_.Completed -and $_.IsCommand -ne $true }
    
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
        Write-Host "`n  📭 No archived tasks." -ForegroundColor Gray
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

#endregion

#region Task Status Functions

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

#endregion

#region Display Functions

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
            
            # Show task count (excluding command snippets)
            $taskCount = ($script:Data.Tasks | Where-Object { $_.ProjectKey -eq $proj.Key -and $_.IsCommand -ne $true -and -not $_.Completed }).Count
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
    
    # Use Format-TableUnicode for details
    $details = @(
        [PSCustomObject]@{Property="Name"; Value=$project.Name}
        [PSCustomObject]@{Property="Key"; Value=$ProjectKey}
        [PSCustomObject]@{Property="ID1"; Value=$project.Id1}
        [PSCustomObject]@{Property="ID2"; Value=$project.Id2}
        [PSCustomObject]@{Property="Client"; Value=$project.Client}
        [PSCustomObject]@{Property="Department"; Value=$project.Department}
        [PSCustomObject]@{Property="Status"; Value=$project.Status}
        [PSCustomObject]@{Property="Billing Type"; Value=$project.BillingType}
    )
    
    if ($project.BillingType -eq "Billable") {
        $details += [PSCustomObject]@{Property="Rate"; Value="`$$($project.Rate)/hr"}
    }
    
    if ($project.Budget -gt 0) {
        $details += [PSCustomObject]@{Property="Budget"; Value="$($project.Budget) hours"}
        
        if ($project.TotalHours -gt 0) {
            $percentUsed = [Math]::Round(($project.TotalHours / $project.Budget) * 100, 1)
            $details += [PSCustomObject]@{Property="Budget Used"; Value="$percentUsed%"}
            $details += [PSCustomObject]@{Property="Remaining"; Value="$([Math]::Round($project.Budget - $project.TotalHours, 2)) hours"}
        }
    }
    
    if ($project.TotalHours -gt 0) {
        $details += [PSCustomObject]@{Property="Total Hours"; Value=$project.TotalHours}
        
        if ($project.BillingType -eq "Billable" -and $project.Rate -gt 0) {
            $totalValue = $project.TotalHours * $project.Rate
            $details += [PSCustomObject]@{Property="Total Value"; Value="`$$([Math]::Round($totalValue, 2))"}
        }
    }
    
    $details | Format-TableUnicode -Columns @(
        @{Name="Property"; Title="Property"; Width=20}
        @{Name="Value"; Title="Value"; Width=40}
    ) -BorderStyle "Rounded"
    
    # Task Summary
    if ($project.ActiveTasks -gt 0 -or $project.CompletedTasks -gt 0) {
        Write-Host "`nTask Summary:" -ForegroundColor Yellow
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
        Write-Host "`nRecent Time Entries:" -ForegroundColor Yellow
        $entryData = $recentEntries | ForEach-Object {
            $task = if ($_.TaskId) {
                $t = $script:Data.Tasks | Where-Object { $_.Id -eq $_.TaskId }
                if ($t) { $t.Description } else { "" }
            } else { "" }
            
            [PSCustomObject]@{
                Date = [DateTime]::Parse($_.Date).ToString("MMM dd")
                Hours = "$($_.Hours)h"
                Task = if ($task) { $task } else { $_.Description }
            }
        }
        
        $entryData | Format-TableUnicode -Columns @(
            @{Name="Date"; Title="Date"; Width=10}
            @{Name="Hours"; Title="Hours"; Width=8; Align="Right"}
            @{Name="Task"; Title="Description"; Width=40}
        ) -BorderStyle "None"
    }
    
    # Active Tasks
    $activeTasks = $script:Data.Tasks | Where-Object {
        $_.ProjectKey -eq $ProjectKey -and -not $_.Completed -and $_.IsCommand -ne $true
    } | Select-Object -First 5
    
    if ($activeTasks.Count -gt 0) {
        Write-Host "`nActive Tasks:" -ForegroundColor Yellow
        foreach ($task in $activeTasks) {
            $priorityInfo = Get-PriorityInfo $task.Priority
            Write-Host "  $($priorityInfo.Icon) [$($task.Id.Substring(0,6))] $($task.Description)"
        }
    }
}

#endregion

#region Excel Import Functions

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
                    $projectData[$fieldConfig.Field] = $value.Trim()
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

#endregion

#region Task Management Views

function Show-TasksView {
    param(
        [string]$Filter = "",
        [string]$SortBy = "Smart",
        [switch]$ShowCompleted,
        [string]$View = "Default"
    )
    
    # Apply filter
    $filtered = $script:Data.Tasks | Where-Object { $_.IsCommand -ne $true }
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
        if ($choice -match '^([cdep])\s+(.+)) {
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

#endregion