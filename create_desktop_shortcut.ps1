# PowerShell script to create a standalone Windows Desktop Shortcut for SEC 10-K Financial Analyzer
$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$distPortableExe = Join-Path $scriptDir "dist\portable\SEC10KAnalyzer.exe"
$distSingleExe = Join-Path $scriptDir "dist\SEC10KAnalyzer_portable.exe"
$distFolderExe = Join-Path $scriptDir "dist\SEC10KAnalyzer\SEC10KAnalyzer.exe"

$targetExe = $null
if (Test-Path $distPortableExe) {
    $targetExe = $distPortableExe
} elseif (Test-Path $distSingleExe) {
    $targetExe = $distSingleExe
} elseif (Test-Path $distFolderExe) {
    $targetExe = $distFolderExe
} else {
    Write-Error "Could not find SEC10KAnalyzer.exe in dist folder. Please build the application first."
    exit 1
}

$desktopPath = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktopPath "SEC 10-K Financial Analyzer.lnk"

Write-Host "Creating desktop shortcut..." -ForegroundColor Cyan
Write-Host "  Target Executable: $targetExe"
Write-Host "  Shortcut Location: $shortcutPath"

$wshShell = New-Object -ComObject WScript.Shell
$shortcut = $wshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $targetExe
$shortcut.WorkingDirectory = Split-Path -Parent $targetExe
$shortcut.Description = "SEC 10-K Financial Analyzer Desktop Application"

$iconPath = Join-Path $scriptDir "app_icon.ico"
if (Test-Path $iconPath) {
    $shortcut.IconLocation = "$iconPath,0"
} else {
    $shortcut.IconLocation = "$targetExe,0"
}

$shortcut.Save()

Write-Host "`n[SUCCESS] Desktop shortcut created successfully!" -ForegroundColor Green
Write-Host "You can now launch the SEC 10-K Financial Analyzer directly from your Windows Desktop." -ForegroundColor White
