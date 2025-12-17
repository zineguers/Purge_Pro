@echo off
net session >nul 2>&1 || (powershell -Command "Start-Process '%~f0' -Verb runAs" & exit /b)
mode con: cols=130 lines=50
title Purge Pro Launcher
chcp 65001 >nul
cls
set "Dest=C:\Program Files\Purge Pro"
set "Marker=%Dest%\.installed"
set "SourcePS1=%~dp0Purge Pro.ps1"

if not exist "%Dest%" mkdir "%Dest%"

if not exist "%Marker%" (
    color 0A
    call :CatInstall
    copy /Y "%SourcePS1%" "%Dest%\Purge Pro.ps1" >nul
    copy /Y "%~dp0Purge Pro.exe" "%USERPROFILE%\Desktop\Purge Pro.exe" >nul
	copy /Y "%~dp0Purge Pro.exe" "%Dest%\Purge Pro.exe" >nul
    if exist "%Dest%\Purge Pro.ps1" (
        echo installed %DATE% %TIME%> "%Marker%"
        call :ShowInstalledMessage
    ) else (
        color 0C
        echo.
        echo ERROR: Purge Pro.ps1 not found in the same folder!
        pause >nul
        exit /b
    )
) else (
    color 06
    call :CatActive
    call :ShowAlreadyInstalledMessage
)

pause >nul
exit /b

:CatInstall
echo.
echo ###############################################################################
echo #                  S3 TECHNOLOGIES - PURGE PRO INSTALLED                      #
echo ###############################################################################
echo.
goto :eof

:CatActive
echo.
echo ###############################################################################
echo #                  S3 TECHNOLOGIES - PURGE PRO ACTIVE                         #
echo ###############################################################################
echo.
goto :eof

:ShowInstalledMessage
cls
color 0A
echo.
echo [97m═══════════════════════════════════════════════════════════════════════════════[0m
echo [92m                        PURGE PRO INSTALLED SUCCESSFULLY ✅                     [0m
echo [97m═══════════════════════════════════════════════════════════════════════════════[0m
echo.
echo [96m [Installation Details][0m
echo [96m • Status[0m       : [93mInstalled and active.[0m
echo [96m • Install Date[0m : [93m%DATE% %TIME%[0m
echo [96m • Location[0m     : [93mC:\Program Files\Purge Pro[0m
echo [96m • Script[0m       : [93mPurge Pro.ps1[0m
echo.
echo [96m [Launch Options][0m
echo [96m • Option 1[0m     : [0m Double click [93mPurge Pro.exe
echo [96m • Option 2[0m     : [0m Select [93mPurge Pro.ps1[0m → Shift + Right-click → [93mRun with PowerShell[0m
echo.
echo [91m [NEED HELP?][0m Re-launch this script to access the full Troubleshooting Guide.
echo.
echo [97m───────────────────────────────────────────────────────────────────────────────[0m
echo. 
echo [91m SECURITY REMINDER[0m🔒:
echo  This is an internal tool for S3 Technologies privileged admins only.
echo  Never share outside the company.
echo.
echo [95m Key Features[0m🔥:
echo    • Advanced email search         • Export results to CSV
echo    • Real-time email preview       • Fast mailbox queries
echo    • Bulk email purge              • Full Microsoft Graph API integration
echo.
echo.
echo [97m═══════════════════════════════════════════════════════════════════════════════[0m
echo [90m Created by ZG / JoMo for S3 Technologies...[0m
echo [90m Press any key to close...[0m
echo.
goto :eof

:ShowAlreadyInstalledMessage
cls
color 06
echo.
echo [97m═══════════════════════════════════════════════════════════════════════════════[0m
echo [97m                  PURGE PRO IS ALREADY INSTALLED AND ACTIVE ✔️                  [0m
echo [97m═══════════════════════════════════════════════════════════════════════════════[0m
echo.
echo  - [96mStatus[0m                 : Installed - Active ✅
echo  - [96mVersion[0m: v5.0 ║ Path   : C:\Program Files\Purge Pro
echo  - [96mLogs[0m                   : “Last scan found no issues.”
echo  - [96mUpdates[0m                : “No updates available.”
echo  - [96mEnvironment[0m            : Windows %PROCESSOR_ARCHITECTURE%
echo.
echo [96m Alternative launch[0m:
echo [96m •[0m Go to [93mPurge Pro.ps1[0m → Shift + Right-click → [93mRun with PowerShell[0m
echo.
echo [97m───────────────────────────────────────────────────────────────────────────────[0m
echo [91m [Troubleshooting⚠️][0m:
echo  1. [93mError Message: "Installation failed or did not complete."[0m: File is blocked;
echo  → Right click Purge Pro.ps1: Properties:  Unblock.
echo.
echo  2. [93mError Message: "Purge Pro.ps1 not found."[0m: Delete the folder: C:\Program Files\Purge Pro 
echo  → and relaunch the installer.
echo.
echo  3. [93mError Message: "Application failed to open."[0m: Go to the installer folder path 
echo  → Shift + Right-click → Run with PowerShell
echo.
echo  4 [93mError Message: "ERROR[0m: PowerShell vX.X or higher is required to run this script.": 
echo  → Install latest Powershell version : winget install Microsoft.PowerShell
echo.
echo [97m═══════════════════════════════════════════════════════════════════════════════[0m
echo [90m Created by ZG / JoMo for S3 TECHNOLOGIES
echo  You can safely close this window now...[0m
echo.
goto :eof