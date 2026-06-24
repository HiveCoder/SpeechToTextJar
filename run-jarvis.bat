@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "JARVIS_DIR=%SCRIPT_DIR%Jarvis"
set "SCRIPT_PATH=%JARVIS_DIR%\jarvis.py"
set "ROOT_REQUIREMENTS=%SCRIPT_DIR%requirements.txt"
set "PREFERRED_PYTHON=C:\Users\asust\AppData\Local\Programs\Python\Python311\python.exe"
set "PYTHON_EXE="
set "PYTHON_ARGS="

if not exist "%SCRIPT_PATH%" (
    echo Could not find Jarvis script at:
    echo %SCRIPT_PATH%
    pause
    exit /b 1
)

call :resolve_python
if errorlevel 1 goto :python_not_found

if /I "%~1"=="run" goto :run_assistant
if /I "%~1"=="install" goto :install_deps
if /I "%~1"=="check" goto :check_env
if /I "%~1"=="help" goto :show_help
if /I "%~1"=="" goto :menu

echo Unknown option: %~1
echo.
goto :show_help

:menu
cls
echo ======================================
echo           Jarvis Launcher
echo ======================================
echo.
echo [1] Run Jarvis
echo [2] Install or repair dependencies
echo [3] Check Python and required packages
echo [4] Show help
echo [5] Quit
echo.
choice /c 12345 /n /m "Select an option: "

if errorlevel 5 exit /b 0
if errorlevel 4 goto :show_help
if errorlevel 3 goto :check_env
if errorlevel 2 goto :install_deps
if errorlevel 1 goto :run_assistant

:run_assistant
call :check_imports
if errorlevel 1 (
    echo.
    echo Jarvis dependencies are missing or broken.
    choice /c YN /n /m "Install dependencies now? [Y/N]: "
    if errorlevel 2 goto :end
    goto :install_deps
)

pushd "%JARVIS_DIR%"
call :run_python "%SCRIPT_PATH%"
set "EXIT_CODE=%ERRORLEVEL%"
popd

echo.
if not "%EXIT_CODE%"=="0" (
    echo Jarvis exited with code %EXIT_CODE%.
)
pause
exit /b %EXIT_CODE%

:install_deps
echo.
echo Installing Jarvis dependencies...
pushd "%SCRIPT_DIR%"
call :run_python -m pip install -r "%ROOT_REQUIREMENTS%"
if errorlevel 1 (
    set "EXIT_CODE=%ERRORLEVEL%"
    popd
    echo.
    echo Failed to install requirements.
    pause
    exit /b %EXIT_CODE%
)
call :run_python -m pip install pyaudio pywin32
set "EXIT_CODE=%ERRORLEVEL%"
popd

echo.
if "%EXIT_CODE%"=="0" (
    echo Dependencies installed successfully.
) else (
    echo Some dependencies failed to install. Exit code: %EXIT_CODE%
)
pause
exit /b %EXIT_CODE%

:check_env
echo.
echo Python command:
echo   %PYTHON_EXE% %PYTHON_ARGS%
echo.
call :run_python --version
echo.
call :check_imports
set "EXIT_CODE=%ERRORLEVEL%"
if "%EXIT_CODE%"=="0" (
    echo All required imports are available.
) else (
    echo One or more required imports are missing.
)
echo.
if defined OPENAI_API_KEY (
    echo OpenAI API key detected. AI chat fallback is enabled.
) else (
    echo OPENAI_API_KEY is not set. Jarvis will keep voice commands, but AI chat fallback is disabled.
)
echo.
pause
exit /b %EXIT_CODE%

:show_help
echo.
echo Usage:
echo   run-jarvis.bat run
echo   run-jarvis.bat install
echo   run-jarvis.bat check
echo   run-jarvis.bat help
echo.
echo With no arguments, the launcher opens an interactive menu.
echo Set OPENAI_API_KEY before running Jarvis to enable ChatGPT-like fallback replies.
echo.
pause
exit /b 0

:resolve_python
if exist "%SCRIPT_DIR%.venv\Scripts\python.exe" (
    call :python_supported "%SCRIPT_DIR%.venv\Scripts\python.exe"
    if not errorlevel 1 (
    set "PYTHON_EXE=%SCRIPT_DIR%.venv\Scripts\python.exe"
    set "PYTHON_ARGS="
    exit /b 0
    )
)

if exist "%PREFERRED_PYTHON%" (
    call :python_supported "%PREFERRED_PYTHON%"
    if not errorlevel 1 (
    set "PYTHON_EXE=%PREFERRED_PYTHON%"
    set "PYTHON_ARGS="
    exit /b 0
    )
)

where py >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_EXE=py"
    set "PYTHON_ARGS=-3.11"
    exit /b 0
)

where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    call :python_supported python
    if not errorlevel 1 (
    set "PYTHON_EXE=python"
    set "PYTHON_ARGS="
    exit /b 0
    )
)

exit /b 1

:python_supported
"%~1" -c "import sys; raise SystemExit(0 if sys.version_info[:2] == (3, 11) else 1)" >nul 2>nul
exit /b %ERRORLEVEL%

:check_imports
call :run_python -c "import pyttsx3, speech_recognition, wikipedia, pyautogui, pyjokes, win32com.client, openai"
exit /b %ERRORLEVEL%

:run_python
"%PYTHON_EXE%" %PYTHON_ARGS% %*
exit /b %ERRORLEVEL%

:python_not_found
echo Python 3.11 was not found.
echo Install Python 3.11 or update run-jarvis.bat with your Python path.
echo.
pause
exit /b 1

:end
exit /b 0