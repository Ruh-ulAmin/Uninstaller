@echo off
echo Running the uninstallation script...
echo.

REM Check if the script is running with administrative privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo This script requires administrative privileges. Please run it as an administrator.
    pause
    exit
)

REM Uninstall software using the product name
echo Uninstalling Software1...
wmic product where "name='Software1'" call uninstall /nointeractive

echo Uninstalling Software2...
wmic product where "name='Software2'" call uninstall /nointeractive

REM Additional configurations or commands can be added here

REM Uncomment the following lines if PowerShell is available on the system
REM echo Uninstalling Software3...
REM powershell.exe -Command "Start-Process 'msiexec.exe' -ArgumentList '/x {ProductCode} /passive' -Verb RunAs"

echo Uninstallation completed.
pause
