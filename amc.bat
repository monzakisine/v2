@echo off
REM ==========================================================
REM  amc.bat  -  AMC Automation launcher
REM
REM  Double-click  -> opens a friendly menu (for the nurse)
REM  Called with args from a terminal -> runs that command directly
REM      amc scms
REM      amc all -DryRun
REM      amc altamimi -NoArchive
REM
REM  If you move the project, update SCRIPTDIR below.
REM ==========================================================

setlocal

set "SCRIPTDIR=C:\AMC-Automation\Scripts"
set "MENU=%SCRIPTDIR%\Run-Menu.ps1"
set "ENGINE=%SCRIPTDIR%\Update-Tracker.ps1"

if not exist "%MENU%" (
    echo.
    echo  [ERROR] Cannot find Run-Menu.ps1 at:
    echo          %MENU%
    echo.
    echo  Edit amc.bat and fix SCRIPTDIR.
    echo.
    pause
    exit /b 1
)

if "%~1"=="" (
    REM Double-clicked: launch the menu
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%MENU%"
) else (
    REM Called with args: run the engine directly
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%ENGINE%" %*
)

exit /b %ERRORLEVEL%
