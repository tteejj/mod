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

function Initialize-ThemeSystem {
    # Load theme from settings or use default
    if ($script:Data -and $script:Data.Settings -and $script:Data.Settings.Theme) {
        # For now, just use the legacy theme colors
        $script:CurrentTheme = @{
            Name = "Legacy"
            Description = "Legacy theme from settings"
            Palette = @{
                PrimaryFG = "#FFFFFF"
                SecondaryFG = "#808080"
                AccentFG = $(Get-ConsoleColorHex $script:Data.Settings.Theme.Accent)
                SuccessFG = $(Get-ConsoleColorHex $script:Data.Settings.Theme.Success)
                ErrorFG = $(Get-ConsoleColorHex $script:Data.Settings.Theme.Error)
                WarningFG = $(Get-ConsoleColorHex $script:Data.Settings.Theme.Warning)
                InfoFG = $(Get-ConsoleColorHex $script:Data.Settings.Theme.Info)
                HeaderFG = $(Get-ConsoleColorHex $script:Data.Settings.Theme.Header)
                SubtleFG = $(Get-ConsoleColorHex $script:Data.Settings.Theme.Subtle)
            }
            DataTable = $script:DefaultTheme.DataTable
        }
    } else {
        $script:CurrentTheme = $script:DefaultTheme
    }
}

function Get-ConsoleColorHex {
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
    
    if ($colorMap.ContainsKey($ColorName)) {
        return $colorMap[$ColorName]
    }
    return "#FFFFFF"
}

function Get-BorderStyleChars {
    param(
        [string]$Style = "Single"
    )
    
    if ($Global:borderStyles.ContainsKey($Style)) {
        return $Global:borderStyles[$Style]
    }
    return $Global:borderStyles.Single
}

function Get-ThemeProperty {
    param(
        [string]$Path
    )
    
    $parts = $Path -split '\.'
    $current = $script:CurrentTheme
    
    foreach ($part in $parts) {
        if ($current -is [hashtable] -and $current.ContainsKey($part)) {
            $current = $current[$part]
        } else {
            return $null
        }
    }
    
    return $current
}

#endregion

#region PSStyle Support

function Get-PSStyleValue {
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

function Apply-PSStyle {
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

function Write-Header {
    param([string]$Text)
    
    # Use legacy color if available, otherwise use theme
    if ($script:Data -and $script:Data.Settings -and $script:Data.Settings.Theme) {
        Write-Host "`n$Text" -ForegroundColor $script:Data.Settings.Theme.Header
        Write-Host ("=" * $Text.Length) -ForegroundColor DarkCyan
    } else {
        $headerColor = Get-ThemeProperty "Palette.HeaderFG"
        if ($headerColor) {
            Write-Host "`n$(Apply-PSStyle -Text $Text -FG $headerColor -Bold)"
            Write-Host ("=" * $Text.Length) -ForegroundColor DarkCyan
        } else {
            Write-Host "`n$Text" -ForegroundColor Cyan
            Write-Host ("=" * $Text.Length) -ForegroundColor DarkCyan
        }
    }
}

function Write-Success {
    param([string]$Text)
    
    if ($script:Data -and $script:Data.Settings -and $script:Data.Settings.Theme) {
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

function Write-Warning {
    param([string]$Text)
    
    if ($script:Data -and $script:Data.Settings -and $script:Data.Settings.Theme) {
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

function Write-Error {
    param([string]$Text)
    
    if ($script:Data -and $script:Data.Settings -and $script:Data.Settings.Theme) {
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

function Write-Info {
    param([string]$Text)
    
    if ($script:Data -and $script:Data.Settings -and $script:Data.Settings.Theme) {
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

function Edit-ThemeSettings {
    Write-Header "Theme Settings"
    
    Write-Host "Current theme colors:" -ForegroundColor Yellow
    Write-Host "  Header:  " -NoNewline; Write-Host "Sample" -ForegroundColor $script:Data.Settings.Theme.Header
    Write-Host "  Success: " -NoNewline; Write-Host "Sample" -ForegroundColor $script:Data.Settings.Theme.Success
    Write-Host "  Warning: " -NoNewline; Write-Host "Sample" -ForegroundColor $script:Data.Settings.Theme.Warning
    Write-Host "  Error:   " -NoNewline; Write-Host "Sample" -ForegroundColor $script:Data.Settings.Theme.Error
    Write-Host "  Info:    " -NoNewline; Write-Host "Sample" -ForegroundColor $script:Data.Settings.Theme.Info
    Write-Host "  Accent:  " -NoNewline; Write-Host "Sample" -ForegroundColor $script:Data.Settings.Theme.Accent
    Write-Host "  Subtle:  " -NoNewline; Write-Host "Sample" -ForegroundColor $script:Data.Settings.Theme.Subtle
    
    Write-Host "`nAvailable colors:" -ForegroundColor Gray
    Write-Host "Black, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, Gray,"
    Write-Host "DarkGray, Blue, Green, Cyan, Red, Magenta, Yellow, White"
    
    Write-Host "`nLeave empty to keep current color" -ForegroundColor Gray
    
    $colors = @("Header", "Success", "Warning", "Error", "Info", "Accent", "Subtle")
    foreach ($colorType in $colors) {
        $newColor = Read-Host "$colorType color"
        if ($newColor) {
            $script:Data.Settings.Theme[$colorType] = $newColor
        }
    }
    
    Save-UnifiedData
    Initialize-ThemeSystem
    Write-Success "Theme settings updated!"
}

#endregion