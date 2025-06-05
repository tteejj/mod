# Helper Functions Module
# Utility functions for file I/O, date handling, validation, etc.

#region Configuration

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

#endregion

#region Data Persistence

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

function Restore-FromBackup {
    Write-Header "Restore from Backup"
    
    $backups = Get-ChildItem $script:BackupPath -Filter "backup_*.json" | Sort-Object CreationTime -Descending
    
    if ($backups.Count -eq 0) {
        Write-Warning "No backups found"
        return
    }
    
    Write-Host "Available backups:"
    for ($i = 0; $i -lt $backups.Count; $i++) {
        $backup = $backups[$i]
        $date = $backup.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
        $size = [Math]::Round($backup.Length / 1KB, 2)
        Write-Host "  [$i] $date ($size KB)"
    }
    
    $index = Read-Host "`nSelect backup number"
    try {
        $idx = [int]$index
        if ($idx -ge 0 -and $idx -lt $backups.Count) {
            $selectedBackup = $backups[$idx]
            
            Write-Warning "This will replace all current data! Type 'yes' to confirm"
            $confirm = Read-Host
            
            if ($confirm -eq 'yes') {
                # Backup current data first
                Backup-Data -Silent
                
                # Load backup
                $backupData = Get-Content $selectedBackup.FullName | ConvertFrom-Json -AsHashtable
                $script:Data = $backupData
                Save-UnifiedData
                
                Write-Success "Data restored from backup!"
                Write-Info "A backup of your previous data was created."
            }
        }
    } catch {
        Write-Error "Invalid selection"
    }
}

#endregion

#region ID Generation

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

#endregion

#region Date Functions

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

function Get-NextWeekday {
    param([int]$TargetDay)
    
    $today = [datetime]::Today
    $currentDay = [int]$today.DayOfWeek
    $daysToAdd = ($TargetDay - $currentDay + 7) % 7
    if ($daysToAdd -eq 0) { $daysToAdd = 7 }
    
    return $today.AddDays($daysToAdd)
}

#endregion

#region Validation Functions

function Test-ExcelConnection {
    Write-Header "Test Excel Connection"
    
    try {
        Write-Info "Testing Excel COM object creation..."
        $excel = New-Object -ComObject Excel.Application
        Write-Success "Excel COM object created successfully!"
        
        Write-Info "Excel version: $($excel.Version)"
        
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
        Write-Success "Excel connection test passed!"
    } catch {
        Write-Error "Excel connection test failed: $_"
        Write-Warning "Make sure Microsoft Excel is installed on this system."
    }
}

#endregion

#region Import/Export Functions

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
    
    # Export command snippets
    $commands = $script:Data.Tasks | Where-Object { $_.IsCommand -eq $true }
    if ($commands.Count -gt 0) {
        $commandExport = $commands | ForEach-Object {
            [PSCustomObject]@{
                Id = $_.Id
                Name = $_.Description
                Command = $_.Notes
                Category = $_.Category
                Tags = $_.Tags -join ","
                Hotkey = if ($_.Hotkey) { $_.Hotkey } else { "" }
                CreatedDate = $_.CreatedDate
            }
        }
        $commandExport | Export-Csv (Join-Path $exportPath "command_snippets.csv") -NoTypeInformation
    }
    
    Write-Success "Data exported to: $exportPath"
    
    # Open folder
    Start-Process $exportPath
}

function Import-Data {
    Write-Header "Import Data"
    
    Write-Warning "This will merge imported data with existing data."
    Write-Host "Enter path to unified_data.json file:"
    $importPath = Read-Host
    
    if (-not (Test-Path $importPath)) {
        Write-Error "File not found!"
        return
    }
    
    try {
        # Backup current data first
        Backup-Data -Silent
        
        $importData = Get-Content $importPath | ConvertFrom-Json -AsHashtable
        
        Write-Host "`nImport options:"
        Write-Host "[1] Merge with existing data"
        Write-Host "[2] Replace all data"
        Write-Host "[3] Cancel"
        
        $choice = Read-Host "Choice"
        
        switch ($choice) {
            "1" {
                # Merge data
                if ($importData.Projects) {
                    foreach ($key in $importData.Projects.Keys) {
                        if (-not $script:Data.Projects.ContainsKey($key)) {
                            $script:Data.Projects[$key] = $importData.Projects[$key]
                            Write-Success "Imported project: $key"
                        } else {
                            Write-Warning "Skipped existing project: $key"
                        }
                    }
                }
                
                if ($importData.Tasks) {
                    $existingIds = $script:Data.Tasks | ForEach-Object { $_.Id }
                    $imported = 0
                    foreach ($task in $importData.Tasks) {
                        if ($task.Id -notin $existingIds) {
                            $script:Data.Tasks += $task
                            $imported++
                        }
                    }
                    Write-Success "Imported $imported new tasks"
                }
                
                if ($importData.TimeEntries) {
                    $existingIds = $script:Data.TimeEntries | ForEach-Object { $_.Id }
                    $imported = 0
                    foreach ($entry in $importData.TimeEntries) {
                        if ($entry.Id -notin $existingIds) {
                            $script:Data.TimeEntries += $entry
                            $imported++
                        }
                    }
                    Write-Success "Imported $imported new time entries"
                }
                
                Save-UnifiedData
                Write-Success "Data merge complete!"
            }
            "2" {
                Write-Warning "Replace all data? Type 'yes' to confirm"
                $confirm = Read-Host
                if ($confirm -eq 'yes') {
                    $script:Data = $importData
                    Save-UnifiedData
                    Write-Success "Data replaced successfully!"
                }
            }
            "3" {
                Write-Info "Import cancelled"
            }
        }
    } catch {
        Write-Error "Import failed: $_"
    }
}

#endregion

#region Reset Functions

function Reset-ToDefaults {
    Write-Header "Reset to Defaults"
    
    Write-Warning "This will reset all settings to defaults. Data will be preserved."
    Write-Host "Type 'yes' to confirm:"
    $confirm = Read-Host
    
    if ($confirm -eq 'yes') {
        # Backup first
        Backup-Data -Silent
        
        # Reset settings while preserving data
        $script:Data.Settings = Get-DefaultSettings
        
        Save-UnifiedData
        Initialize-ThemeSystem
        
        Write-Success "Settings reset to defaults!"
    }
}

#endregion

#region Clipboard Functions

function Copy-ToClipboard {
    param([string]$Text)
    
    try {
        $Text | Set-Clipboard
        return $true
    } catch {
        Write-Warning "Could not copy to clipboard: $_"
        return $false
    }
}

function Get-FromClipboard {
    try {
        return Get-Clipboard
    } catch {
        Write-Warning "Could not read from clipboard: $_"
        return $null
    }
}

#endregion