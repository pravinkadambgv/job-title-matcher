@echo off
:: -------------------------------------------------------
::  Email Notifier — Launcher
::  Author: Pravin Kadam
::
::  HOW TO USE:
::    1. Make sure Outlook is open and logged in
::    2. Double-click this file
::    3. A tray icon appears — you'll get a popup
::       alert whenever a new email arrives
::    4. Right-click the tray icon to exit
:: -------------------------------------------------------

title Email Notifier

:: Check that OutlookEmailNotifier.ps1 is in the same folder
if not exist "%~dp0OutlookEmailNotifier.ps1" (
    echo ERROR: OutlookEmailNotifier.ps1 not found in the same folder.
    echo Please keep both files together.
    pause
    exit /b 1
)

:: Launch silently (no console window) using Windows PowerShell
powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0OutlookEmailNotifier.ps1"

exit /b 0
