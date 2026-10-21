$WshShell = New-Object -ComObject WScript.Shell
$desktop = [System.Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop "Table Editor & Exporter.lnk"
$oldShortcutPath = Join-Path $desktop "Tabulkovy editor.lnk"
if (Test-Path $oldShortcutPath) {
    Remove-Item -Force $oldShortcutPath
}
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$htmlPath = Join-Path $scriptDir "index.html"
$iconPath = Join-Path $scriptDir "app-icon.ico"

$shortcut = $WshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $htmlPath
$shortcut.WorkingDirectory = $scriptDir
if (Test-Path $iconPath) {
    $shortcut.IconLocation = "$iconPath,0"
}
$shortcut.Description = "Table Editor & Exporter - Professional Spreadsheet & Export Engine"
$shortcut.Save()

Write-Host "Odkaz na ploche bol uspesne vytvoreny: $shortcutPath"
