@echo off
setlocal
cd /d "%~dp0"
echo Starting Repair Calculator...
echo.
if exist "App\start.bat" (
  cd /d "%~dp0App"
  call start.bat
) else (
  echo [ERROR] App\start.bat not found.
  echo Trying direct launch...
  cd /d "%~dp0App"
  if exist "node_modules\electron\dist\electron.exe" (
    "node_modules\electron\dist\electron.exe" .
  ) else (
    call npx --yes electron .
  )
  pause
)
endlocal
