@echo off
setlocal
cd /d "%~dp0"

set "PYTHON_EXE=%~dp0outlookEmailPlus\.venv\Scripts\python.exe"
if not exist "%PYTHON_EXE%" (
    echo [ERROR] Python virtual environment not found: %PYTHON_EXE%
    echo Please install the OutlookEmailPlus environment first.
    pause
    exit /b 1
)

"%PYTHON_EXE%" "%~dp0release_outlook_email_pool.py"
set "EXIT_CODE=%ERRORLEVEL%"
pause
exit /b %EXIT_CODE%
