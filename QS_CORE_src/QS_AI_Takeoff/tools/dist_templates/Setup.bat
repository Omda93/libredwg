@echo off
setlocal EnableExtensions EnableDelayedExpansion
title QS CORE - Setup
cd /d "%~dp0"

rem ---------------------------------------------------------------------------
rem  QS CORE installer (no admin rights needed)
rem  Usage:  Setup.bat                -> installs to %LOCALAPPDATA%\DWG_QS_Tools
rem          Setup.bat "D:\Tools\QS"  -> installs to the given folder
rem  Set QS_NO_SHORTCUTS=1 to skip Desktop / Start-menu shortcuts.
rem  Set QS_NO_PAUSE=1 for unattended runs.
rem ---------------------------------------------------------------------------
set "SRC=%~dp0app"
set "WHEELS=%~dp0wheels"
set "TARGET=%LOCALAPPDATA%\DWG_QS_Tools"
if not "%~1"=="" set "TARGET=%~1"
set "LOG=%TEMP%\DWG_QS_Tools_setup.log"
echo QS CORE setup %DATE% %TIME% > "%LOG%"

echo.
echo  ==========================================================
echo     QS CORE  -  Setup
echo     QS CORE  -  the all-in-one QS platform
echo  ==========================================================
echo   Install folder : %TARGET%
echo   Log file       : %LOG%
echo.

if not exist "%SRC%\DWG_Measure_Extractor_V6.py" (
    echo  [ERROR] "app" folder not found next to Setup.bat. Extract the whole ZIP first.
    goto :fail
)

rem ---------------------------------------------------------------- 1. Python
echo  [1/5] Looking for Python 3.12 (64-bit) ...
set "PY="
py -3.12 -c "import sys,tkinter;sys.exit(0 if sys.maxsize>2**32 else 1)" >nul 2>&1 && set "PY=py -3.12"
if not defined PY (
    python -c "import sys,tkinter;sys.exit(0 if sys.version_info[:2]==(3,12) and sys.maxsize>2**32 else 1)" >nul 2>&1 && set "PY=python"
)
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PY="%LOCALAPPDATA%\Programs\Python\Python312\python.exe""
)
if not defined PY (
    set "PYINST="
    for %%F in ("%~dp0python\python-3.12*-amd64.exe") do set "PYINST=%%~fF"
    if not defined PYINST (
        echo  [ERROR] Python 3.12 is not installed and the bundled installer is missing.
        echo          Install Python 3.12 64-bit from https://www.python.org and run Setup again.
        goto :fail
    )
    echo         Python 3.12 not found - installing bundled Python for this user only ...
    "!PYINST!" /passive InstallAllUsers=0 PrependPath=1 Include_launcher=1 Include_tcltk=1 Include_test=0 Shortcuts=0 >>"%LOG%" 2>&1
    if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
        set "PY="%LOCALAPPDATA%\Programs\Python\Python312\python.exe""
    ) else (
        echo  [ERROR] Python installation failed. See %LOG%
        goto :fail
    )
)
rem version via temp file: safe for quoted python paths with spaces (no for /f quoting issues)
set "PYVER=unknown"
%PY% -c "import sys;print(sys.version.split()[0])" > "%TEMP%\qs_pyver.txt" 2>nul
if exist "%TEMP%\qs_pyver.txt" set /p PYVER=<"%TEMP%\qs_pyver.txt"
del "%TEMP%\qs_pyver.txt" >nul 2>&1
echo         Using Python %PYVER%  (%PY%)
echo Python: %PY% %PYVER% >> "%LOG%"

rem ---------------------------------------------------------------- 2. copy files
echo  [2/5] Copying program files ...
if not exist "%TARGET%" mkdir "%TARGET%"
rem code is always refreshed; user settings (config\*.json) are kept if they already exist
robocopy "%SRC%" "%TARGET%" /E /R:1 /W:1 /NFL /NDL /NJH /NJS /NP /XD "__pycache__" "config" >>"%LOG%" 2>&1
if errorlevel 8 ( echo  [ERROR] Copy failed. See %LOG% & goto :fail )
robocopy "%SRC%\QS_AI_Takeoff\config" "%TARGET%\QS_AI_Takeoff\config" /E /XC /XN /XO /R:1 /W:1 /NFL /NDL /NJH /NJS /NP >>"%LOG%" 2>&1
if errorlevel 8 ( echo  [ERROR] Config copy failed. See %LOG% & goto :fail )
copy /y "%~dp0Uninstall.bat" "%TARGET%\Uninstall.bat" >nul

rem ---------------------------------------------------------------- 3. venv
echo  [3/5] Creating private Python environment ...
set "VPY=%TARGET%\.venv\Scripts\python.exe"
if not exist "%VPY%" (
    %PY% -m venv "%TARGET%\.venv" >>"%LOG%" 2>&1
    if errorlevel 1 ( echo  [ERROR] Could not create the virtual environment. See %LOG% & goto :fail )
)

rem ---------------------------------------------------------------- 4. packages
echo  [4/5] Installing packages (offline first) ...
"%VPY%" -m pip install --disable-pip-version-check --no-index --find-links "%WHEELS%" -r "%SRC%\requirements-runtime.txt" >>"%LOG%" 2>&1
if errorlevel 1 (
    echo         Offline install incomplete - trying the internet ...
    "%VPY%" -m pip install --disable-pip-version-check -r "%SRC%\requirements-runtime.txt" >>"%LOG%" 2>&1
    if errorlevel 1 ( echo  [ERROR] Package installation failed. See %LOG% & goto :fail )
)
"%VPY%" -c "import shapely, ezdxf, openpyxl, numpy, tkinter, pythoncom, win32com.client; print('imports OK')" >>"%LOG%" 2>&1
if errorlevel 1 ( echo  [ERROR] Package check failed. See %LOG% & goto :fail )
echo         Packages OK.

rem ---------------------------------------------------------------- 5. shortcuts
echo  [5/5] Creating shortcuts ...
if "%QS_NO_SHORTCUTS%"=="1" (
    echo         skipped ^(QS_NO_SHORTCUTS=1^)
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%TARGET%\make_shortcuts.ps1" -Target "%TARGET%" >>"%LOG%" 2>&1
    if errorlevel 1 ( echo         [WARN] Shortcuts could not be created - use the .bat files in %TARGET% ) else ( echo         Desktop + Start menu shortcuts created. )
)

echo.
echo  ==========================================================
echo   Installation complete.
echo   Start:  Desktop shortcut "QS CORE"
echo           (Walls ^& Rooms opens from its tab, or its own shortcut)
echo   Folder: %TARGET%
echo   AutoCAD LISP files: %TARGET%\QS_AI_Takeoff\autocad
echo  ==========================================================
echo SUCCESS >> "%LOG%"
if not "%QS_NO_PAUSE%"=="1" pause
exit /b 0

:fail
echo FAILED >> "%LOG%"
echo.
echo  Setup did not complete. Send the log file to the tool owner:
echo     %LOG%
if not "%QS_NO_PAUSE%"=="1" pause
exit /b 1
