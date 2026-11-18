@echo off
setlocal
cd /d "%~dp0"
echo ==========================================
echo  Repair Calculator - Desktop
echo ==========================================
echo.
where node >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Node.js not found. Install from https://nodejs.org
  pause
  exit /b 1
)
echo [OK] Node:
node -v
echo.

if not exist "node_modules\xlsx\package.json" call npm install --no-audit --no-fund
if not exist "node_modules\electron\package.json" call npm install --no-audit --no-fund
if exist "node_modules\electron\package.json" if not exist "node_modules\electron\dist\electron.exe" (
  echo [SETUP] Installing Electron binary...
  node node_modules\electron\install.js
)

echo [START] Launching...
echo.
if exist "node_modules\electron\dist\electron.exe" (
  "node_modules\electron\dist\electron.exe" .
) else (
  call npx --yes electron .
)
set EXITCODE=%ERRORLEVEL%
echo.
if "%EXITCODE%"=="0" (echo [DONE] App closed.) else (echo [ERROR] Exit code %EXITCODE%)
pause
endlocal
