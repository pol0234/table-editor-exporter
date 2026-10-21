$WshShell = New-Object -ComObject WScript.Shell
$desktop = [System.Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop "Tabulkovy editor.lnk"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Check Chrome or Edge
$chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
if (-not (Test-Path $chromePath)) {
    $chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
}
if (-not (Test-Path $chromePath)) {
    $chromePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
}

$htmlPath = Join-Path $scriptDir "index.html"
$iconPath = Join-Path $scriptDir "app-icon.ico"

$shortcut = $WshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $chromePath
$shortcut.Arguments = "--app=`"file:///$($htmlPath.Replace('\', '/'))`""
$shortcut.WorkingDirectory = $scriptDir
if (Test-Path $iconPath) {
    $shortcut.IconLocation = "$iconPath,0"
}
$shortcut.Description = "Tabulkovy editor nutricnych hodnot a potravin"
$shortcut.Save()

Write-Host "Odkaz na ploche bol uspesne vytvoreny: $shortcutPath"
