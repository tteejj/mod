# Theme System Module
# Handles colors, styles, and visual presentation

#region Border Styles

$Global:borderStyles = @{
    None = @{
        TopLeft     = " "
        TopRight    = " "
        BottomLeft  = " "
        BottomRight = " "
        Horizontal  = " "
        Vertical    = " "
        Cross       = " "
        TLeft       = " "
        TRight      = " "
        TTop        = " "
        TBottom     = " "
    }
    Single = @{
        TopLeft     = "┌"
        TopRight    = "┐"
        BottomLeft  = "└"
        BottomRight = "┘"
        Horizontal  = "─"
        Vertical    = "│"
        Cross       = "┼"
        TLeft       = "├"
        TRight      = "┤"
        TTop        = "┬"
        TBottom     = "┴"
    }
    Double = @{
        TopLeft     = "╔"
        TopRight    = "╗"
        BottomLeft  = "╚"
        BottomRight = "╝"
        Horizontal  = "═"
        Vertical    = "║"
        Cross       = "╬"
        TLeft       = "╠"
        TRight      = "╣"
        TTop        = "╦"
        TBottom     = "╩"
    }
    Rounded = @{
        TopLeft     = "╭"
        TopRight    = "╮"
        BottomLeft  = "╰"
        BottomRight = "╯"
        Horizontal  = "─"
        Vertical    = "│"
        Cross       = "┼"
        TLeft       = "├"
        TRight      = "┤"
        TTop        = "┬"
        TBottom     = "┴"
    }
    Heavy = @{
        TopLeft     = "┏"
        TopRight    = "┓"
        BottomLeft  = "┗"
        BottomRight = "┛"
        Horizontal  = "━"
        Vertical    = "┃"
        Cross       = "╋"
        TLeft       = "┣"
        TRight      = "┫"
        TTop        = "┳"
        TBottom     = "┻"
    }
}

#endregion

#region Theme Management

# Default theme structure
$script:DefaultTheme = @{
    Name = "Default"
    Description = "Clean default theme"
    Palette = @{
        PrimaryFG = "#FFFFFF"
        SecondaryFG = "#808080"
        AccentFG = "#00D7FF"
        SuccessFG = "#5FFF87"
        ErrorFG = "#FF005F"
        WarningFG = "#FFFF00"
        InfoFG = "#5FD7FF"
        HeaderFG = "#00D7FF"
        SubtleFG = "#808080"
    }
    DataTable = @{
        BorderStyle = "Single"
        BorderFG = "#FFFFFF"
        Pad = 1
        Header = @{
            FG = "#00D7FF"
            BG = $null
            Separator = $true
            Case = "Default"
        }
        DataRow = @{
            FG = "#FFFFFF"
            BG = $null
        }
        AltRow = @{
            FG = "#808080"
            BG = $null
        }
        Highlight = @{
            Overdue = @{FG = "#FF005F"}
            DueSoon = @{FG = "#FFFF00"}
            Completed = @{FG = "#808080"}
            Selected = @{FG = "#000000"; BG = "#00D7FF"}
        }
    }
}

# Current theme storage
$script:CurrentTheme = $null

function global:Initialize-ThemeSystem {
    # Load theme from settings or use default
    if ($script:Data -and $script:Data.Settings -and $script:Data.Settings.Theme) {
        # Use the legacy theme colors if $script:Data.Settings.Theme is fully populated
        $legacyTheme = $script:Data.Settings.Theme
        if ($legacyTheme.Accent -and $legacyTheme.Success -and $legacyTheme.Error -and $legacyTheme.Warning -and $legacyTheme.Info -and $legacyTheme.Header -and $legacyTheme.Subtle) {
            $script:CurrentTheme = @{
                Name = "Legacy"
                Description = "Legacy theme from settings"
                Palette = @{
                    PrimaryFG = "#FFFFFF"
                    SecondaryFG = "#808080"
                    AccentFG = $(Get-ConsoleColorHex $legacyTheme.Accent)
                    SuccessFG = $(Get-ConsoleColorHex $legacyTheme.Success)
                    ErrorFG = $(Get-ConsoleColorHex $legacyTheme.Error)
                    WarningFG = $(Get-ConsoleColorHex $legacyTheme.Warning)
                    InfoFG = $(Get-ConsoleColorHex $legacyTheme.Info)
                    HeaderFG = $(Get-ConsoleColorHex $legacyTheme.Header)
                    SubtleFG = $(Get-ConsoleColorHex $legacyTheme.Subtle)
                }
                DataTable = $script:DefaultTheme.DataTable
            }
        } else {
            # If legacy theme is incomplete, fall back to default
            $script:CurrentTheme = $script:DefaultTheme
            Write-Warning "Legacy theme settings incomplete, using default theme."
        }
    } else {
        $script:CurrentTheme = $script:DefaultTheme
    }
}

function global:Get-ConsoleColorHex {
    param($ColorName)
    
    $colorMap = @{
        "Black" = "#000000"
        "DarkBlue" = "#000080"
        "DarkGreen" = "#008000"
        "DarkCyan" = "#008080"
        "DarkRed" = "#800000"
        "DarkMagenta" = "#800080"
        "DarkYellow" = "#808000"
        "Gray" = "#C0C0C0"
        "DarkGray" = "#808080"
        "Blue" = "#0000FF"
        "Green" = "#00FF00"
        "Cyan" = "#00FFFF"
        "Red" = "#FF0000"
        "Magenta" = "#FF00FF"
        "Yellow" = "#FFFF00"
        "White" = "#FFFFFF"
    }
    
    if ($ColorName -and $colorMap.ContainsKey($ColorName)) {
        return $colorMap[$ColorName]
    }
    return "#FFFFFF" # Default to white if color name is null or not found
}

function global:Get-BorderStyleChars {
    param(
        [string]$Style = "Single"
    )
    
    if ($Global:borderStyles.ContainsKey($Style)) {
        return $Global:borderStyles[$Style]
    }
    return $Global:borderStyles.Single
}

function global:Get-ThemeProperty {
    param(
        [string]$Path
    )
    
    $parts = $Path -split '\.'
    $current = $script:CurrentTheme
    
    foreach ($part in $parts) {
        if ($current -is [hashtable] -and $current.ContainsKey($part)) {
            $current = $current[$part]
        } else {
            # Fallback for palette colors if not found, to prevent errors
            if ($Path -like "Palette.*") {
                # Write-Warning "Theme property '$Path' not found, using default." # Can be noisy
                switch ($part) {
                    "HeaderFG" { return $script:DefaultTheme.Palette.HeaderFG }
                    "SuccessFG" { return $script:DefaultTheme.Palette.SuccessFG }
                    "ErrorFG" { return $script:DefaultTheme.Palette.ErrorFG }
                    "WarningFG" { return $script:DefaultTheme.Palette.WarningFG }
                    "InfoFG" { return $script:DefaultTheme.Palette.InfoFG }
                    "AccentFG" { return $script:DefaultTheme.Palette.AccentFG }
                    "SubtleFG" { return $script:DefaultTheme.Palette.SubtleFG }
                    default { return $script:DefaultTheme.Palette.PrimaryFG }
                }
            }
            return $null
        }
    }
    
    return $current
}

#endregion

#region PSStyle Support

function global:Get-PSStyleValue {
    param(
        [string]$FG,
        [string]$BG,
        [switch]$Bold,
        [switch]$Italic,
        [switch]$Underline
    )
    
    # For PowerShell 7.2+ with PSStyle support
    if ($PSVersionTable.PSVersion.Major -ge 7 -and $PSVersionTable.PSVersion.Minor -ge 2) {
        $style = ""
        
        if ($FG) {
            if ($FG -match '^#[0-9A-Fa-f]{6}$') {
                $style += "`e[38;2;$([Convert]::ToInt32($FG.Substring(1,2), 16));$([Convert]::ToInt32($FG.Substring(3,2), 16));$([Convert]::ToInt32($FG.Substring(5,2), 16))m"
            }
        }
        
        if ($BG) {
            if ($BG -match '^#[0-9A-Fa-f]{6}$') {
                $style += "`e[48;2;$([Convert]::ToInt32($BG.Substring(1,2), 16));$([Convert]::ToInt32($BG.Substring(3,2), 16));$([Convert]::ToInt32($BG.Substring(5,2), 16))m"
            }
        }
        
        if ($Bold) { $style += "`e[1m" }
        if ($Italic) { $style += "`e[3m" }
        if ($Underline) { $style += "`e[4m" }
        
        return $style
    }
    
    return ""
}

function global:Apply-PSStyle {
    param(
        [string]$Text,
        [string]$FG,
        [string]$BG,
        [switch]$Bold,
        [switch]$Italic,
        [switch]$Underline
    )
    
    $style = Get-PSStyleValue -FG $FG -BG $BG -Bold:$Bold -Italic:$Italic -Underline:$Underline
    
    if ($style) {
        return "${style}${Text}`e[0m"
    }
    
    return $Text
}

#endregion

#region Theme-Based UI Functions

function global:Write-Header {
    param([string]$Text)
    
    # Use legacy color if available, otherwise use theme
    if ($script:Data -and $script:Data.Settings -and $script:Data.Settings.Theme -and $script:Data.Settings.Theme.Header) {
        Write-Host "`n$Text" -ForegroundColor $script:Data.Settings.Theme.Header
        Write-Host ("=" * $Text.Length) -ForegroundColor DarkCyan # Keep underline consistent or theme it too
    } else {
        $headerColor = Get-ThemeProperty "Palette.HeaderFG"
        if ($headerColor) {
            Write-Host "`n$(Apply-PSStyle -Text $Text -FG $headerColor -Bold)"
            Write-Host ("=" * $Text.Length) -ForegroundColor DarkCyan # Consider theming this underline too
        } else { # Fallback if Get-ThemeProperty returns null
            Write-Host "`n$Text" -ForegroundColor Cyan
            Write-Host ("=" * $Text.Length) -ForegroundColor DarkCyan
        }
    }
}

function global:Write-Success {
    param([string]$Text)
    
    if ($script:Data -and $script:Data.Settings -and $script:Data.Settings.Theme -and $script:Data.Settings.Theme.Success) {
        Write-Host "✓ $Text" -ForegroundColor $script:Data.Settings.Theme.Success
    } else {
        $color = Get-ThemeProperty "Palette.SuccessFG"
        if ($color) {
            Write-Host "$(Apply-PSStyle -Text "✓ $Text" -FG $color)"
        } else {
            Write-Host "✓ $Text" -ForegroundColor Green
        }
    }
}

function global:Write-Warning {
    param([string]$Text)
    
    if ($script:Data -and $script:Data.Settings -and $script:Data.Settings.Theme -and $script:Data.Settings.Theme.Warning) {
        Write-Host "⚠ $Text" -ForegroundColor $script:Data.Settings.Theme.Warning
    } else {
        $color = Get-ThemeProperty "Palette.WarningFG"
        if ($color) {
            Write-Host "$(Apply-PSStyle -Text "⚠ $Text" -FG $color)"
        } else {
            Write-Host "⚠ $Text" -ForegroundColor Yellow
        }
    }
}

function global:Write-Error {
    param([string]$Text)
    
    if ($script:Data -and $script:Data.Settings -and $script:Data.Settings.Theme -and $script:Data.Settings.Theme.Error) {
        Write-Host "✗ $Text" -ForegroundColor $script:Data.Settings.Theme.Error
    } else {
        $color = Get-ThemeProperty "Palette.ErrorFG"
        if ($color) {
            Write-Host "$(Apply-PSStyle -Text "✗ $Text" -FG $color)"
        } else {
            Write-Host "✗ $Text" -ForegroundColor Red
        }
    }
}

function global:Write-Info {
    param([string]$Text)
    
    if ($script:Data -and $script:Data.Settings -and $script:Data.Settings.Theme -and $script:Data.Settings.Theme.Info) {
        Write-Host "ℹ $Text" -ForegroundColor $script:Data.Settings.Theme.Info
    } else {
        $color = Get-ThemeProperty "Palette.InfoFG"
        if ($color) {
            Write-Host "$(Apply-PSStyle -Text "ℹ $Text" -FG $color)"
        } else {
            Write-Host "ℹ $Text" -ForegroundColor Blue
        }
    }
}

#endregion

#region Theme Configuration

function global:Edit-ThemeSettings {
    Write-Header "Theme Settings"
    
    # Ensure $script:Data.Settings.Theme exists and has defaults if not fully set
    if (-not $script:Data.Settings.Theme) {
        # This assumes Get-DefaultSettings is available and returns the legacy theme structure
        $script:Data.Settings.Theme = (Get-DefaultSettings).Theme 
    }
    # Ensure all expected keys exist in $script:Data.Settings.Theme, falling back to defaults from Get-DefaultSettings
    $defaultLegacyThemeColors = (Get-DefaultSettings).Theme # Get-DefaultSettings must be loaded
    foreach($colorKey in $defaultLegacyThemeColors.Keys){
        if(-not $script:Data.Settings.Theme.ContainsKey($colorKey) -or !$script:Data.Settings.Theme[$colorKey]){
            $script:Data.Settings.Theme[$colorKey] = $defaultLegacyThemeColors[$colorKey]
            Write-Warning "Theme color '$colorKey' was missing, reset to default '$($defaultLegacyThemeColors[$colorKey])'."
        }
    }

    Write-Host "Current theme colors (legacy ConsoleColor names):" -ForegroundColor Yellow
    Write-Host "  Header:  " -NoNewline; Write-Host "Sample Header Text" -ForegroundColor $script:Data.Settings.Theme.Header
    Write-Host "  Success: " -NoNewline; Write-Host "✓ Sample Success Text" -ForegroundColor $script:Data.Settings.Theme.Success
    Write-Host "  Warning: " -NoNewline; Write-Host "⚠ Sample Warning Text" -ForegroundColor $script:Data.Settings.Theme.Warning
    Write-Host "  Error:   " -NoNewline; Write-Host "✗ Sample Error Text" -ForegroundColor $script:Data.Settings.Theme.Error
    Write-Host "  Info:    " -NoNewline; Write-Host "ℹ Sample Info Text" -ForegroundColor $script:Data.Settings.Theme.Info
    Write-Host "  Accent:  " -NoNewline; Write-Host "Sample Accent Text" -ForegroundColor $script:Data.Settings.Theme.Accent
    Write-Host "  Subtle:  " -NoNewline; Write-Host "Sample Subtle Text" -ForegroundColor $script:Data.Settings.Theme.Subtle
    
    Write-Host "`nAvailable ConsoleColor names:" -ForegroundColor Gray
    Write-Host ([System.Enum]::GetNames([System.ConsoleColor]) -join ", ")
    
    Write-Host "`nLeave empty to keep current color" -ForegroundColor Gray
    
    $colorsToEdit = @("Header", "Success", "Warning", "Error", "Info", "Accent", "Subtle") # Renamed variable
    foreach ($colorType in $colorsToEdit) {
        $newColorName = Read-Host "$colorType color ($($script:Data.Settings.Theme[$colorType]))" # Renamed variable
        if (-not [string]::IsNullOrWhiteSpace($newColorName)) {
            # Validate if it's a known ConsoleColor
            try {
                $null = [System.Enum]::Parse([System.ConsoleColor], $newColorName, $true) # $true for case-insensitive
                $script:Data.Settings.Theme[$colorType] = $newColorName
            } catch {
                Write-Warning "Invalid color name '$newColorName'. Keeping current."
            }
        }
    }
    
    Save-UnifiedData # Save-UnifiedData must be global
    Initialize-ThemeSystem # Re-initialize to apply changes to $script:CurrentTheme (modern hex-based theme)
    Write-Success "Theme settings updated! Restart may be needed for full effect if legacy settings were used by Write-Host directly."
}

#endregion