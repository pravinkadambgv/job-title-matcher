@echo off
:: -------------------------------------------------------
::  Email Notifier — Stop
::  Author: Pravin Kadam
::
::  Run this only if you can't right-click the tray icon.
:: -------------------------------------------------------

taskkill /F /IM powershell.exe /FI "WINDOWTITLE eq Email Notifier" >nul 2>&1
if %errorlevel% == 0 (
    echo Email Notifier stopped.
) else (
    echo No running instance found.
)
pause
