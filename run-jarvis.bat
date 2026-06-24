@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "JARVIS_DIR=%SCRIPT_DIR%Jarvis"
set "SCRIPT_PATH=%JARVIS_DIR%\jarvis.py"
set "PYTHON_EXE=C:\Users\asust\AppData\Local\Programs\Python\Python311\python.exe"

if not exist "%SCRIPT_PATH%" (
    echo Could not find Jarvis script at:
    echo %SCRIPT_PATH%
    pause
    exit /b 1
)

if exist "%PYTHON_EXE%" (
    pushd "%JARVIS_DIR%"
    "%PYTHON_EXE%" "%SCRIPT_PATH%"
    set "EXIT_CODE=%ERRORLEVEL%"
    popd
    pause
    exit /b %EXIT_CODE%
)

where py >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    pushd "%JARVIS_DIR%"
    py -3.11 "%SCRIPT_PATH%"
    set "EXIT_CODE=%ERRORLEVEL%"
    popd
    pause
    exit /b %EXIT_CODE%
)

echo Python 3.11 was not found.
echo Install Python 3.11 or update run-jarvis.bat with your Python path.
pause
exit /b 1