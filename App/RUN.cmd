@echo off
setlocal
cd /d "%~dp0"
title Repair Calculator
echo Starting...
"%~dp0node_modules\.bin\electron.cmd" "%~dp0."
echo.
echo Exit code: %ERRORLEVEL%
pause
endlocal
