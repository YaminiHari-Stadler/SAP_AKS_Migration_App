@echo off
rem ---------------------------------------------------------------------------
rem  AKS Migration App
rem  Convert legacy AKS lists to the AKS V2 format
rem
rem  Double-click this file. Nothing else is needed.
rem  The first start installs everything the application needs, inside this
rem  folder only. Later starts open the application straight away.
rem ---------------------------------------------------------------------------

title AKS Migration App
cd /d "%~dp0"

powershell.exe -NoProfile -NoLogo -ExecutionPolicy Bypass -File "%~dp0install_and_run.ps1" %*
set "EXITCODE=%ERRORLEVEL%"

if not "%EXITCODE%"=="0" (
    echo.
    echo ---------------------------------------------------------------
    echo  The AKS Migration App could not be started.
    echo  The reason is written above and in the logs folder.
    echo ---------------------------------------------------------------
    echo.
    pause
)

exit /b %EXITCODE%
