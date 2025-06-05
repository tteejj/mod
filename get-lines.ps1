# Get specific lines from a file
param(
    [string]$FilePath = "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-data.ps1",
    [int]$StartLine = 600,
    [int]$EndLine = 610
)

$lines = Get-Content $FilePath
$lineCount = $lines.Count

Write-Host "Total lines in file: $lineCount" -ForegroundColor Yellow
Write-Host "Showing lines $StartLine to $EndLine:" -ForegroundColor Green
Write-Host "----------------------------------------"

for ($i = $StartLine - 1; $i -lt $EndLine -and $i -lt $lineCount; $i++) {
    $lineNum = $i + 1
    Write-Host "${lineNum}: $($lines[$i])"
}

# Also search for the problematic line patterns
Write-Host "`n`nSearching for problematic Sort-Object patterns..." -ForegroundColor Yellow
$lineNum = 0
foreach ($line in $lines) {
    $lineNum++
    if ($line -match 'Sort-Object.*UseCount.*Descending.*@\{Expression') {
        Write-Host "Found at line ${lineNum}: $line" -ForegroundColor Red
    }
}

# Search for the $[Math] pattern
Write-Host "`n`nSearching for problematic variable reference patterns..." -ForegroundColor Yellow
$lineNum = 0
foreach ($line in $lines) {
    $lineNum++
    if ($line -match '\$\[Math\]') {
        Write-Host "Found at line ${lineNum}: $line" -ForegroundColor Red
    }
}
