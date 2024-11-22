@ECHO OFF
@setlocal enableextensions
@cd /d "%~dp0"

REM --> Check for permissions using net session
net session >nul 2>&1
if %errorlevel% NEQ 0 (
    echo Requesting administrative privileges...
    if "%*" == "" (
        powershell.exe -Command "Start-Process -FilePath '%~0' -Verb RunAs"
    ) else (
        powershell.exe -Command "Start-Process -FilePath '%~0' -ArgumentList '%*' -Verb RunAs"
    )
    exit /b
)

REM --> We have admin privileges
:gotAdmin
cls

REM --> Ensure PowerShell script exists
if not exist "%~dp0pyrevit-Install-Main.ps1" (
    echo PowerShell script not found.
    echo Press any key to exit.
    pause >nul
    goto :EOF
)

REM --> Run the PowerShell script
powershell.exe -executionpolicy bypass -Command "& '%~dp0pyrevit-Install-Main.ps1'"
