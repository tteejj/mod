# Helper Functions Module
# Utility functions for file I/O, date handling, validation, etc.

#region Configuration

$script:DataPath = Join-Path $env:USERPROFILE ".ProductivitySuite"
$script:UnifiedDataFile = Join-Path $script:DataPath "unified_data.json"
$script:ConfigFile = Join-Path $script:DataPath "config.json" # Not actively used in provided code
$script:BackupPath = Join-Path $script:DataPath "backups"
$script:ExcelCopyConfigFile = Join-Path $script:DataPath "excelcopy_configs.json" # Not actively used

# Ensure directories exist
@($script:DataPath, $script:BackupPath) | ForEach-Object {
    if (-not (Test-Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
    }
}

#endregion

#region PowerShell 5.1 Compatibility Functions

function global:ConvertFrom-JsonToHashtable {
    param([string]$JsonString)
    
    function Convert-PSObjectToHashtable {
        param($InputObject)
        
        if ($null -eq $InputObject) { return $null }
        
        if ($InputObject -is [PSCustomObject]) {
            $hashtable = @{}
            $InputObject.PSObject.Properties | ForEach-Object {
                $hashtable[$_.Name] = Convert-PSObjectToHashtable $_.Value
            }
            return $hashtable
        }
        elseif ($InputObject -is [array]) {
            return @($InputObject | ForEach-Object { Convert-PSObjectToHashtable $_ })
        }
        else {
            return $InputObject
        }
    }
    
    $psobject = $JsonString | ConvertFrom-Json
    return Convert-PSObjectToHashtable $psobject
}

#endregion

#region Data Persistence

function global:Load-UnifiedData {
    try {
        if (Test-Path $script:UnifiedDataFile) {
            $jsonContent = Get-Content $script:UnifiedDataFile -Raw
            $loadedData = ConvertFrom-JsonToHashtable $jsonContent
            
            # Ensure $script:Data and $script:Data.Settings are initialized with defaults first
            # This is guaranteed by core-data.ps1 loading before this function is called from main.ps1
            if (-not $script:Data) {
                Write-Error "CRITICAL: \$script:Data not initialized before Load-UnifiedData call. This indicates a script loading order problem."
                # Attempt to initialize to prevent further errors, though this indicates a load order problem
                # Get-DefaultSettings must be available (defined in core-data.ps1)
                $script:Data = @{ Settings = (Get-DefaultSettings); Projects = @{}; Tasks = @(); TimeEntries = @(); ActiveTimers = @{}; ArchivedTasks = @{}; ExcelCopyJobs = @{}; CurrentWeek = (Get-WeekStart (Get-Date)) } 
            } elseif (-not $script:Data.Settings) {
                 $script:Data.Settings = (Get-DefaultSettings)
            }

            # Deep merge to preserve structure and handle missing keys in saved data
            foreach ($topLevelKey in $loadedData.Keys) {
                if ($topLevelKey -eq "Settings" -and $script:Data.ContainsKey($topLevelKey) -and $script:Data.Settings -is [hashtable]) {
                    # Merge settings carefully, ensuring all default keys are present
                    $defaultSettings = Get-DefaultSettings
                    foreach ($settingKey in $defaultSettings.Keys) {
                        if ($loadedData.Settings.ContainsKey($settingKey)) {
                            if ($settingKey -eq "Theme" -and $loadedData.Settings.Theme -is [hashtable] -and $defaultSettings.Theme -is [hashtable]) {
                                # Deep merge for Theme specifically
                                foreach ($themeColorKey in $defaultSettings.Theme.Keys) {
                                    if ($loadedData.Settings.Theme.ContainsKey($themeColorKey)) {
                                        $script:Data.Settings.Theme[$themeColorKey] = $loadedData.Settings.Theme[$themeColorKey]
                                    } # Else, keep default theme color
                                }
                            } elseif ($settingKey -eq "TimeTrackerTemplates" -and $loadedData.Settings.TimeTrackerTemplates -is [hashtable]) {
                                # For TimeTrackerTemplates, replace entirely or merge carefully
                                $script:Data.Settings.TimeTrackerTemplates = $loadedData.Settings.TimeTrackerTemplates
                            } elseif ($settingKey -eq "CommandSnippets" -and $loadedData.Settings.CommandSnippets -is [hashtable]) {
                                foreach($csKey in $defaultSettings.CommandSnippets.Keys){
                                     if($loadedData.Settings.CommandSnippets.ContainsKey($csKey)){
                                         $script:Data.Settings.CommandSnippets[$csKey] = $loadedData.Settings.CommandSnippets[$csKey]
                                     } # Else, keep default
                                }
                            } elseif ($settingKey -eq "ExcelFormConfig" -and $loadedData.Settings.ExcelFormConfig -is [hashtable]) {
                                # Deep merge ExcelFormConfig - Replace ?: with if-else
                                if ($loadedData.Settings.ExcelFormConfig.WorksheetName) {
                                    $script:Data.Settings.ExcelFormConfig.WorksheetName = $loadedData.Settings.ExcelFormConfig.WorksheetName
                                } else {
                                    $script:Data.Settings.ExcelFormConfig.WorksheetName = $defaultSettings.ExcelFormConfig.WorksheetName
                                }
                                foreach($fieldKey in $defaultSettings.ExcelFormConfig.StandardFields.Keys){
                                    if($loadedData.Settings.ExcelFormConfig.StandardFields.ContainsKey($fieldKey)){
                                        $script:Data.Settings.ExcelFormConfig.StandardFields[$fieldKey] = $loadedData.Settings.ExcelFormConfig.StandardFields[$fieldKey]
                                    } # Else, keep default
                                }
                            }
                            else {
                                $script:Data.Settings[$settingKey] = $loadedData.Settings[$settingKey]
                            }
                        } # Else, the key is missing in loaded data, so the default from Get-DefaultSettings (already in $script:Data.Settings) remains.
                    }
                } elseif ($script:Data.ContainsKey($topLevelKey)) { # For other top-level keys like Projects, Tasks, etc.
                    $script:Data[$topLevelKey] = $loadedData[$topLevelKey]
                } # Else, loaded data has a top-level key not in default $script:Data structure; ignore it to maintain schema.
            }
            
            # Ensure CurrentWeek is a DateTime
            if ($script:Data.CurrentWeek -is [string]) {
                try {
                    $script:Data.CurrentWeek = [DateTime]::Parse($script:Data.CurrentWeek)
                } catch {
                    Write-Warning "Could not parse CurrentWeek '$($script:Data.CurrentWeek)'. Resetting to current week."
                    $script:Data.CurrentWeek = Get-WeekStart (Get-Date) # Get-WeekStart must be global
                }
            } elseif ($null -eq $script:Data.CurrentWeek) { # If CurrentWeek was missing entirely
                 $script:Data.CurrentWeek = Get-WeekStart (Get-Date)
            }
        } else {
            Write-Info "No existing data file found. Starting with default data."
            # $script:Data is already initialized with defaults by core-data.ps1
        }
    } catch {
        Write-Warning "Could not load data due to an error, starting fresh: $_"
        # Ensure $script:Data is reset to defaults if loading fails catastrophically
        # This should be handled by the initial setup in core-data.ps1, but as a safeguard:
        $script:Data = @{ Settings = (Get-DefaultSettings); Projects = @{}; Tasks = @(); TimeEntries = @(); ActiveTimers = @{}; ArchivedTasks = @{}; ExcelCopyJobs = @{}; CurrentWeek = (Get-WeekStart (Get-Date)) } 
    }
}

function global:Save-UnifiedData {
    try {
        # Auto-backup
        if ((Get-Random -Maximum 10) -eq 0 -or -not (Test-Path $script:UnifiedDataFile)) {
            Backup-Data -Silent
        }
        
        $script:Data | ConvertTo-Json -Depth 10 | Set-Content $script:UnifiedDataFile -Encoding UTF8
    } catch {
        Write-Error "Failed to save data: $_"
    }
}

function global:Backup-Data {
    param([switch]$Silent)
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupFile = Join-Path $script:BackupPath "backup_$timestamp.json"
    
    try {
        $script:Data | ConvertTo-Json -Depth 10 | Set-Content $backupFile -Encoding UTF8
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

function global:Restore-FromBackup {
    Write-Header "Restore from Backup"
    
    $backups = Get-ChildItem $script:BackupPath -Filter "backup_*.json" | Sort-Object CreationTime -Descending
    
    if ($backups.Count -eq 0) {
        Write-Warning "No backups found"
        return
    }
    
    Write-Host "Available backups:"
    for ($i = 0; $i -lt $backups.Count; $i++) {
        $backupItem = $backups[$i]
        $date = $backupItem.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
        $size = [Math]::Round($backupItem.Length / 1KB, 2)
        Write-Host "  [$i] $date ($size KB)"
    }
    
    $indexInput = Read-Host "`nSelect backup number"
    try {
        $idx = [int]$indexInput
        if ($idx -ge 0 -and $idx -lt $backups.Count) {
            $selectedBackup = $backups[$idx]
            
            Write-Warning "This will replace all current data! Type 'yes' to confirm"
            $confirm = Read-Host
            
            if ($confirm -eq 'yes') {
                # Backup current data first
                Backup-Data -Silent
                
                # Load backup
                $jsonContent = Get-Content $selectedBackup.FullName -Raw
                $backupData = ConvertFrom-JsonToHashtable $jsonContent
                $script:Data = $backupData # This replaces the entire $script:Data
                
                # Re-initialize CurrentWeek to DateTime if it was stored as string in backup
                if ($script:Data.CurrentWeek -is [string]) {
                    try { $script:Data.CurrentWeek = [DateTime]::Parse($script:Data.CurrentWeek) }
                    catch { $script:Data.CurrentWeek = Get-WeekStart (Get-Date) }
                } elseif ($null -eq $script:Data.CurrentWeek) {
                    $script:Data.CurrentWeek = Get-WeekStart (Get-Date)
                }
                # Ensure settings are present, merge with defaults if backup is old/missing some
                $defaultSettings = Get-DefaultSettings
                if (-not $script:Data.Settings) { $script:Data.Settings = $defaultSettings }
                else {
                    foreach($key in $defaultSettings.Keys){
                        if(-not $script:Data.Settings.ContainsKey($key)){
                            $script:Data.Settings[$key] = $defaultSettings[$key]
                        }
                    }
                }

                Save-UnifiedData
                Initialize-ThemeSystem # Re-initialize theme after restoring data
                
                Write-Success "Data restored from backup!"
                Write-Info "A backup of your previous data was created."
            } else { Write-Info "Restore cancelled." }
        } else { Write-Error "Invalid selection number." }
    } catch {
        Write-Error "Invalid selection input: $_"
    }
}

#endregion

#region ID Generation

function global:New-TodoId {
    return [System.Guid]::NewGuid().ToString().Substring(0, 8)
}

function global:Format-Id2 {
    param([string]$Id2Input)

    $id2ToFormat = if ([string]::IsNullOrEmpty($Id2Input)) { "" } else { $Id2Input }
    
    if ($id2ToFormat.Length -gt 9) {
        $id2ToFormat = $id2ToFormat.Substring(0, 9)
    }
    
    $paddingNeeded = 12 - 2 - $id2ToFormat.Length # 12 total, V and S are 2 chars
    $zeros = "0" * [Math]::Max(0, $paddingNeeded) # Ensure padding is not negative
    
    return "V${zeros}${id2ToFormat}S"
}

#endregion

#region Date Functions

function global:Get-WeekStart {
    param([DateTime]$DateInput = (Get-Date))
    
    $daysFromMonday = [int]$DateInput.DayOfWeek
    # DayOfWeek: Sunday = 0, Monday = 1, ..., Saturday = 6
    if ($daysFromMonday -eq 0) { $daysFromMonday = 7 } # Adjust Sunday to be 7 for calculation
    $monday = $DateInput.AddDays(1 - $daysFromMonday) # If Monday (1), 1-1=0. If Sunday (7), 1-7=-6.
    
    return Get-Date $monday -Hour 0 -Minute 0 -Second 0
}

function global:Get-WeekDates {
    param([DateTime]$WeekStartDate)
    
    return @(0..4 | ForEach-Object { $WeekStartDate.AddDays($_) }) # Monday to Friday
}

function global:Format-TodoDate {
    param($DateString)
    if ([string]::IsNullOrEmpty($DateString)) { return "" }
    try {
        $date = [datetime]::Parse($DateString)
        $today = [datetime]::Today # Just the date part
        $diffDays = ($date.Date - $today).Days # Compare Date parts only
        
        $dateStr = $date.ToString("MMM dd")
        if ($diffDays -eq 0) { return "Today" }
        elseif ($diffDays -eq 1) { return "Tomorrow" }
        elseif ($diffDays -eq -1) { return "Yesterday" }
        elseif ($diffDays -gt 1 -and $diffDays -le 7) { return "$dateStr (in $diffDays days)" } 
        elseif ($diffDays -lt -1) { 
            $absDiff = [Math]::Abs($diffDays)
            return "$dateStr ($absDiff days ago)"
        }
        else { return $dateStr } # Covers dates far in future or other unhandled cases
    }
    catch { return $DateString } # Return original string if parsing fails
}

function global:Get-NextWeekday {
    param([int]$TargetDayOfWeek) # 0 for Sunday, 1 for Monday, ..., 6 for Saturday
    
    $today = [datetime]::Today
    $currentDayOfWeek = [int]$today.DayOfWeek
    $daysToAdd = ($TargetDayOfWeek - $currentDayOfWeek + 7) % 7
    if ($daysToAdd -eq 0) { $daysToAdd = 7 } # If today is the target day, get next week's target day
    
    return $today.AddDays($daysToAdd)
}

#endregion

#region Validation Functions

function global:Test-ExcelConnection {
    Write-Header "Test Excel Connection"
    $excel = $null # Initialize for finally block
    try {
        Write-Info "Testing Excel COM object creation..."
        $excel = New-Object -ComObject Excel.Application
        Write-Success "Excel COM object created successfully!"
        
        Write-Info "Excel version: $($excel.Version)"
        
        $excel.Quit()
        # ReleaseComObject calls are important
    } catch {
        Write-Error "Excel connection test failed: $_"
        Write-Warning "Make sure Microsoft Excel is installed on this system."
    } finally {
        if ($excel) {
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
            Remove-Variable excel -ErrorAction SilentlyContinue
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

#endregion

#region Import/Export Functions

function global:Export-AllData {
    Write-Header "Export All Data"
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $exportPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "ProductivitySuite_Export_$timestamp"
    New-Item -ItemType Directory -Path $exportPath -Force | Out-Null
    
    # Export main data
    $script:Data | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $exportPath "unified_data.json") -Encoding UTF8
    
    # Export time entries as CSV
    if ($script:Data.TimeEntries -and $script:Data.TimeEntries.Count -gt 0) {
        $timeExport = $script:Data.TimeEntries | ForEach-Object {
            $project = Get-ProjectOrTemplate $_.ProjectKey
            [PSCustomObject]@{
                Date = $_.Date
                ProjectKey = $_.ProjectKey
                ProjectName = if ($project) { $project.Name } else { "N/A" }
                Hours = $_.Hours
                Description = $_.Description
                TaskId = $_.TaskId
                StartTime = $_.StartTime
                EndTime = $_.EndTime
            }
        }
        $timeExport | Export-Csv (Join-Path $exportPath "time_entries.csv") -NoTypeInformation -Encoding UTF8
    }
    
    # Export tasks as CSV
    if ($script:Data.Tasks -and $script:Data.Tasks.Count -gt 0) {
        $taskExport = $script:Data.Tasks | ForEach-Object {
            $project = if ($_.ProjectKey) { Get-ProjectOrTemplate $_.ProjectKey } else { $null }
            [PSCustomObject]@{
                Id = $_.Id
                Description = $_.Description
                Priority = $_.Priority
                Category = $_.Category
                ProjectName = if ($project) { $project.Name } else { "" }
                Status = Get-TaskStatus $_ # Get-TaskStatus must be global
                DueDate = $_.DueDate
                Progress = $_.Progress
                TimeSpent = $_.TimeSpent
                EstimatedTime = $_.EstimatedTime
                Tags = if ($_.Tags) { $_.Tags -join "," } else { "" }
            }
        }
        $taskExport | Export-Csv (Join-Path $exportPath "tasks.csv") -NoTypeInformation -Encoding UTF8
    }
    
    # Export projects as CSV
    if ($script:Data.Projects -and $script:Data.Projects.Count -gt 0) {
        $projectExport = $script:Data.Projects.GetEnumerator() | ForEach-Object {
            $projValue = $_.Value
            [PSCustomObject]@{
                Key = $_.Key
                Name = $projValue.Name
                Id1 = $projValue.Id1
                Id2 = $projValue.Id2
                Client = $projValue.Client
                Department = $projValue.Department
                Status = $projValue.Status
                BillingType = $projValue.BillingType
                Rate = $projValue.Rate
                Budget = $projValue.Budget
                TotalHours = $projValue.TotalHours
                ActiveTasks = $projValue.ActiveTasks
                CompletedTasks = $projValue.CompletedTasks
            }
        }
        $projectExport | Export-Csv (Join-Path $exportPath "projects.csv") -NoTypeInformation -Encoding UTF8
    }
    
    # Export command snippets
    $commands = $script:Data.Tasks | Where-Object { $_.IsCommand -eq $true }
    if ($commands.Count -gt 0) {
        $commandExport = $commands | ForEach-Object {
            [PSCustomObject]@{
                Id = $_.Id
                Name = $_.Description
                Command = $_.Notes # Command text is in Notes field for snippets
                Category = $_.Category
                Tags = if ($_.Tags) { $_.Tags -join "," } else { "" }
                Hotkey = if ($_.Hotkey) { $_.Hotkey } else { "" }
                CreatedDate = $_.CreatedDate
            }
        }
        $commandExport | Export-Csv (Join-Path $exportPath "command_snippets.csv") -NoTypeInformation -Encoding UTF8
    }
    
    Write-Success "Data exported to: $exportPath"
    
    # Open folder
    try {
        Start-Process $exportPath
    } catch {
        Write-Warning "Could not open export folder: $_"
    }
}

function global:Import-Data {
    Write-Header "Import Data"
    
    Write-Warning "This will merge imported data with existing data or allow full replacement."
    Write-Host "Enter path to unified_data.json file:"
    $importFilePath = Read-Host
    
    if (-not (Test-Path $importFilePath -PathType Leaf)) {
        Write-Error "File not found or is a directory!"
        return
    }
    
    try {
        # Backup current data first
        Backup-Data -Silent
        
        $jsonContent = Get-Content $importFilePath -Raw
        $importedData = ConvertFrom-JsonToHashtable $jsonContent
        
        Write-Host "`nImport options:"
        Write-Host "[1] Merge with existing data (adds new, skips existing by ID/Key)"
        Write-Host "[2] Replace all data (current data will be overwritten!)"
        Write-Host "[3] Cancel"
        
        $choice = Read-Host "Choice"
        
        switch ($choice) {
            "1" { # Merge
                if ($importedData.Projects) {
                    foreach ($key in $importedData.Projects.Keys) {
                        if (-not $script:Data.Projects.ContainsKey($key)) {
                            $script:Data.Projects[$key] = $importedData.Projects[$key]
                            Write-Success "Imported project: $key"
                        } else {
                            Write-Warning "Skipped existing project (key already exists): $key"
                        }
                    }
                }
                
                if ($importedData.Tasks) {
                    $existingTaskIds = $script:Data.Tasks | ForEach-Object { $_.Id }
                    $importedTaskCount = 0
                    foreach ($task in $importedData.Tasks) {
                        if ($task.Id -notin $existingTaskIds) {
                            $script:Data.Tasks += $task
                            $importedTaskCount++
                        }
                    }
                    Write-Success "Imported $importedTaskCount new tasks"
                }
                
                if ($importedData.TimeEntries) {
                    if ($null -eq $script:Data.TimeEntries) { $script:Data.TimeEntries = @() }
                    $existingTimeEntryIds = $script:Data.TimeEntries | ForEach-Object { $_.Id } 
                    $importedTimeEntryCount = 0
                    foreach ($entry in $importedData.TimeEntries) {
                         if ($null -eq $entry.Id -or $entry.Id -notin $existingTimeEntryIds) { 
                            if ($null -eq $entry.Id) { $entry.Id = New-TodoId } 
                            $script:Data.TimeEntries += $entry
                            $importedTimeEntryCount++
                        }
                    }
                    Write-Success "Imported $importedTimeEntryCount new time entries"
                }
                
                # Merge settings carefully if present in import
                if ($importedData.Settings) {
                    Write-Info "Merging settings..."
                    $defaultSettings = Get-DefaultSettings
                    foreach ($settingKey in $defaultSettings.Keys) {
                        if ($importedData.Settings.ContainsKey($settingKey)) {
                            # Handle nested hashtables like Theme, CommandSnippets, ExcelFormConfig carefully
                            if ($script:Data.Settings[$settingKey] -is [hashtable] -and $importedData.Settings[$settingKey] -is [hashtable]) {
                                foreach ($subKey in $importedData.Settings[$settingKey].Keys) {
                                    if ($script:Data.Settings[$settingKey].ContainsKey($subKey)) {
                                        $script:Data.Settings[$settingKey][$subKey] = $importedData.Settings[$settingKey][$subKey]
                                    }
                                }
                            } else {
                                $script:Data.Settings[$settingKey] = $importedData.Settings[$settingKey]
                            }
                        }
                    }
                }

                Save-UnifiedData
                Initialize-ThemeSystem # Re-initialize theme if settings were part of import
                Write-Success "Data merge complete!"
            }
            "2" { # Replace
                Write-Warning "This will REPLACE ALL current data. Are you absolutely sure? Type 'yes' to confirm"
                $confirmReplace = Read-Host
                if ($confirmReplace -eq 'yes') {
                    $script:Data = $importedData
                    # Re-initialize CurrentWeek to DateTime if it was stored as string in import
                    if ($script:Data.CurrentWeek -is [string]) {
                        try { $script:Data.CurrentWeek = [DateTime]::Parse($script:Data.CurrentWeek) }
                        catch { $script:Data.CurrentWeek = Get-WeekStart (Get-Date) }
                    } elseif ($null -eq $script:Data.CurrentWeek) {
                         $script:Data.CurrentWeek = Get-WeekStart (Get-Date)
                    }
                    # Ensure settings structure is complete by merging with defaults
                    $defaultSettings = Get-DefaultSettings
                    if (-not $script:Data.Settings) { $script:Data.Settings = $defaultSettings }
                    else {
                        foreach($key in $defaultSettings.Keys){
                            if(-not $script:Data.Settings.ContainsKey($key)){
                                $script:Data.Settings[$key] = $defaultSettings[$key]
                            }
                            # For nested hashtables, ensure their structure too
                            elseif ($defaultSettings[$key] -is [hashtable] -and $script:Data.Settings[$key] -is [hashtable]) {
                                foreach($subKey in $defaultSettings[$key].Keys){
                                    if(-not $script:Data.Settings[$key].ContainsKey($subKey)){
                                        $script:Data.Settings[$key][$subKey] = $defaultSettings[$key][$subKey]
                                    }
                                }
                            }
                        }
                    }

                    Save-UnifiedData
                    Initialize-ThemeSystem # Re-initialize theme after replacing data
                    Write-Success "Data replaced successfully!"
                } else { Write-Info "Replacement cancelled."}
            }
            "3" {
                Write-Info "Import cancelled"
            }
            default {
                Write-Warning "Invalid choice. Import cancelled."
            }
        }
    } catch {
        Write-Error "Import failed: $_"
    }
}

#endregion

#region Reset Functions

function global:Reset-ToDefaults {
    Write-Header "Reset to Defaults"
    
    Write-Warning "This will reset all settings to defaults. Your data (tasks, projects, time entries) will be preserved."
    Write-Host "Type 'yes' to confirm:"
    $confirm = Read-Host
    
    if ($confirm -eq 'yes') {
        # Backup first
        Backup-Data -Silent
        
        # Reset settings while preserving data
        # Get-DefaultSettings must be available (defined in core-data.ps1)
        $script:Data.Settings = Get-DefaultSettings 
        
        Save-UnifiedData
        Initialize-ThemeSystem # Re-initialize theme after resetting settings
        
        Write-Success "Settings reset to defaults!"
        Write-Info "Your projects, tasks, and time entries remain untouched."
    } else { Write-Info "Reset cancelled."}
}

#endregion

#region Clipboard Functions

function global:Copy-ToClipboard {
    param([string]$TextToCopy)
    
    try {
        $TextToCopy | Set-Clipboard
        return $true
    } catch {
        Write-Warning "Could not copy to clipboard: $_"
        return $false
    }
}

function global:Get-FromClipboard {
    try {
        return Get-Clipboard
    } catch {
        Write-Warning "Could not read from clipboard: $_"
        return $null
    }
}

#endregion