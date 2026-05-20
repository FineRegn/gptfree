@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"
set "SERVICE_DIR=%~dp0outlookEmailPlus"
set "PYTHON_EXE=%SERVICE_DIR%\.venv\Scripts\python.exe"
set "START_SCRIPT=%SERVICE_DIR%\start.py"
set "SHOULD_PAUSE=1"

if /i "%~1"=="--no-pause" set "SHOULD_PAUSE=0"

if not exist "%SERVICE_DIR%" (
    echo [ERROR] OutlookEmailPlus directory not found:
    echo %SERVICE_DIR%
    goto :finish_error
)

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Python virtual environment not found:
    echo %PYTHON_EXE%
    echo.
    echo Please install the OutlookEmailPlus environment first.
    goto :finish_error
)

if not exist "%START_SCRIPT%" (
    echo [ERROR] start.py not found:
    echo %START_SCRIPT%
    goto :finish_error
)

set "PYTHONIOENCODING=utf-8"
cd /d "%SERVICE_DIR%"

echo Starting OutlookEmailPlus email service...
echo Working directory: %CD%
echo Python: %PYTHON_EXE%
echo.

"%PYTHON_EXE%" "%START_SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"
goto :finish

:finish_error
set "EXIT_CODE=1"

:finish
echo.
if not "%EXIT_CODE%"=="0" echo [ERROR] OutlookEmailPlus exited with code %EXIT_CODE%.
if "%SHOULD_PAUSE%"=="1" pause
exit /b %EXIT_CODE%
