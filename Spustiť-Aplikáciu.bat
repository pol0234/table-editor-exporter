@echo off
title Tabuľkový editor
cd /d "%~dp0"
where node >nul 2>&1
if %errorlevel% equ 0 (
    start "" /b node launch.js
    exit
)
set "USER_DATA=%LOCALAPPDATA%\TabulkovyEditorProfile"
set "APP_URL=file:///%~dp0index.html"
set "APP_URL=%APP_URL:\=/%"
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
    start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --app="%APP_URL%" --user-data-dir="%USER_DATA%" --no-first-run --no-default-browser-check
    exit
)
if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" (
    start "" "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --app="%APP_URL%" --user-data-dir="%USER_DATA%" --no-first-run --no-default-browser-check
    exit
)
if exist "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" (
    start "" "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --app="%APP_URL%" --user-data-dir="%USER_DATA%" --no-first-run --no-default-browser-check
    exit
)
start "" "%~dp0index.html"
exit
