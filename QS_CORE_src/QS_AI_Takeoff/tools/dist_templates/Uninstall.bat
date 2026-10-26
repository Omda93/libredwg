@echo off
setlocal EnableExtensions
title QS CORE - Uninstall
rem Usage: Uninstall.bat ["install folder"]   (default %LOCALAPPDATA%\DWG_QS_Tools)
set "TARGET=%LOCALAPPDATA%\DWG_QS_Tools"
if not "%~1"=="" set "TARGET=%~1"

rem When started from inside the install folder, re-launch from TEMP so the folder can be deleted.
if /i "%~dp0"=="%TARGET%\" (
    copy /y "%~f0" "%TEMP%\DWG_QS_Tools_uninstall.bat" >nul
    start "" cmd /c ""%TEMP%\DWG_QS_Tools_uninstall.bat" "%TARGET%""
    exit /b 0
)

echo.
echo  This removes QS CORE from:
echo     %TARGET%
echo  including its settings (config) and logs. Your drawings and Excel reports
echo  saved elsewhere are NOT touched.
echo.
if not "%QS_NO_PAUSE%"=="1" (
    choice /c YN /m "  Continue"
    if errorlevel 2 exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$d=[Environment]::GetFolderPath('Desktop'); $p=[Environment]::GetFolderPath('Programs'); $m=(Join-Path $p 'QS CORE'),(Join-Path $p 'DWG QS Tools');" ^
  "Remove-Item -LiteralPath (Join-Path $d 'DWG Measure Extractor V6.lnk'),(Join-Path $d 'Plaster and Paint Takeoff.lnk'),(Join-Path $d 'QS CORE.lnk'),(Join-Path $d 'QS CORE Walls and Rooms.lnk') -ErrorAction SilentlyContinue;" ^
  "Remove-Item -LiteralPath $m -Recurse -Force -ErrorAction SilentlyContinue"

if exist "%TARGET%" rmdir /s /q "%TARGET%"
if exist "%TARGET%" (
    echo  [WARN] Some files are in use. Close the programs and run Uninstall again.
) else (
    echo  QS CORE removed.
)
if not "%QS_NO_PAUSE%"=="1" pause
exit /b 0
