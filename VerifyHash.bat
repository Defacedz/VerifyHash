@echo off
REM ---------------------------------------------------------------------------
REM  VerifyHash - double-click launcher
REM
REM  Runs the PowerShell script with an execution policy bypass, so it works on
REM  a stock Windows install and on a file still tagged as a web download. With
REM  no argument the script opens its installer window.
REM ---------------------------------------------------------------------------

title VerifyHash

powershell.exe -Sta -NoProfile -ExecutionPolicy Bypass -File "%~dp0VerifyHash.ps1"

if errorlevel 1 (
    echo.
    echo VerifyHash exited with an error.
    pause
)
