# Simple extraction script to get the first copy of core-data.ps1
# This just extracts without trying to fix syntax

$coreDataPath = "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-data.ps1"
$outputPath = "C:\Users\jhnhe\Documents\GitHub\pmc-terminal\modular\core-data-extracted.ps1"

if (-not (Test-Path $coreDataPath)) {
    Write-Host "core-data.ps1 not found!" -ForegroundColor Red
    exit
}

$lines = Get-Content $coreDataPath
Write-Host "Total lines: $($lines.Count)" -ForegroundColor Yellow

# Find all occurrences of the module header
$headerIndices = @()
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -eq "# Core Data Management Module") {
        $headerIndices += $i
        Write-Host "Found header at line $($i + 1)" -ForegroundColor Green
    }
}

if ($headerIndices.Count -gt 1) {
    Write-Host "`nExtracting first module (lines 1 to $($headerIndices[1]))..." -ForegroundColor Cyan
    
    # Extract first occurrence
    $extracted = $lines[0..($headerIndices[1] - 1)]
    
    # Save to new file
    $extracted | Out-File $outputPath -Encoding UTF8
    
    Write-Host "`nExtracted $($extracted.Count) lines to: $outputPath" -ForegroundColor Green
    Write-Host "`nNow you need to:" -ForegroundColor Yellow
    Write-Host "1. Open $outputPath in a text editor" -ForegroundColor White
    Write-Host "2. Fix line 605 - complete the Sort-Object statement" -ForegroundColor White
    Write-Host "3. Save and test with:" -ForegroundColor White
    Write-Host "   Copy-Item '$outputPath' '$coreDataPath'" -ForegroundColor Cyan
    Write-Host "   . .\test-setup.ps1" -ForegroundColor Cyan
} else {
    Write-Host "No duplication found in file" -ForegroundColor Yellow
}
