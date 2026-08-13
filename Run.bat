@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PORT=8790"

if not "%~1"=="" set "PORT=%~1"

echo Starting DocInspector on port %PORT% ...
echo Press Ctrl+C in this window to stop the server.
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%DocInspector.ps1" -Port %PORT%

echo.
echo Server stopped.
pause
