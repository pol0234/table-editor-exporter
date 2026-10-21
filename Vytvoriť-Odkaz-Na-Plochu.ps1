$WshShell = New-Object -ComObject WScript.Shell
$desktop = [System.Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop "Tabulkovy editor.lnk"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$iconPath = Join-Path $scriptDir "app-icon.ico"
$vbsPath = Join-Path $scriptDir "Spustiť-Aplikáciu.vbs"
$batPath = Join-Path $scriptDir "Spustiť-Aplikáciu.bat"
$wscriptPath = "C:\Windows\System32\wscript.exe"

$shortcut = $WshShell.CreateShortcut($shortcutPath)
if ((Test-Path $wscriptPath) -and (Test-Path $vbsPath)) {
    $shortcut.TargetPath = $wscriptPath
    $shortcut.Arguments = "`"$vbsPath`""
} else {
    $shortcut.TargetPath = $batPath
}
$shortcut.WorkingDirectory = $scriptDir
if (Test-Path $iconPath) {
    $shortcut.IconLocation = "$iconPath,0"
}
$shortcut.Description = "Tabuľkový editor - Samostatná aplikácia"
$shortcut.Save()

Write-Host "Odkaz na ploche bol uspesne vytvoreny: $shortcutPath"
