@echo off
setlocal EnableExtensions
title Email Automation V2 Launcher

set "V2_PROJECT=C:\Users\DELL\Documents\Codex\email automation v2\EmailAutomation"
set "V2_EXE=%V2_PROJECT%\outputs\EmailAutomationDeployment\Email Automation.exe"
set "V2_BUILD_SCRIPT=%V2_PROJECT%\build.ps1"
set "LOCALAPPDATA=%V2_PROJECT%\.runtime\LocalAppData"

echo Email Automation V2
echo Project: "%V2_PROJECT%"
echo Executable: "%V2_EXE%"
echo Runtime Config: "%LOCALAPPDATA%"
echo.

if not exist "%V2_PROJECT%" (
  echo ERROR: Version 2 project folder was not found.
  pause
  exit /b 1
)

cd /d "%V2_PROJECT%" || (
  echo ERROR: Could not switch to the Version 2 project folder.
  pause
  exit /b 1
)

if not exist "%LOCALAPPDATA%" mkdir "%LOCALAPPDATA%"

if not exist "%V2_EXE%" (
  echo Version 2 executable was not found. Building now...
  if not exist "%V2_BUILD_SCRIPT%" (
    echo ERROR: Build script was not found:
    echo "%V2_BUILD_SCRIPT%"
    pause
    exit /b 1
  )
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%V2_BUILD_SCRIPT%"
  if errorlevel 1 (
    echo ERROR: Version 2 build failed.
    pause
    exit /b 1
  )
)

if not exist "%V2_EXE%" (
  echo ERROR: Version 2 executable still does not exist after build.
  pause
  exit /b 1
)

echo Starting Email Automation Version 2...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; try { $p = Start-Process -FilePath $env:V2_EXE -WorkingDirectory $env:V2_PROJECT -PassThru -Wait; exit $p.ExitCode } catch { Write-Host 'ERROR: Failed to start Email Automation V2.'; Write-Host $_.Exception.Message; exit 1 }"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
  echo.
  echo ERROR: Email Automation V2 exited with code %EXIT_CODE%.
  echo The console will remain open so this error can be reviewed.
  pause
  exit /b %EXIT_CODE%
)

endlocal
exit /b 0
