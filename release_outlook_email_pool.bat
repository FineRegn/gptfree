@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"
set "PYTHONIOENCODING=utf-8"
set "PYTHON_EXE=%~dp0outlookEmailPlus\.venv\Scripts\python.exe"
set "SCRIPT=%~dp0release_outlook_email_pool.py"
set "SHOULD_PAUSE=1"
set "SCRIPT_ARGS="

:parse_args
if "%~1"=="" goto :after_parse_args
if /i "%~1"=="--no-pause" goto :no_pause_arg
set SCRIPT_ARGS=%SCRIPT_ARGS% "%~1"
shift /1
goto :parse_args

:no_pause_arg
set "SHOULD_PAUSE=0"
shift /1
goto :parse_args

:after_parse_args

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Python virtual environment not found:
    echo %PYTHON_EXE%
    echo.
    echo Please install the OutlookEmailPlus environment first.
    goto :finish_error
)

if not exist "%SCRIPT%" (
    echo [ERROR] Release script not found:
    echo %SCRIPT%
    goto :finish_error
)

"%PYTHON_EXE%" "%SCRIPT%" %SCRIPT_ARGS%
set "EXIT_CODE=%ERRORLEVEL%"
goto :finish

:finish_error
set "EXIT_CODE=1"

:finish
echo.
if "%SHOULD_PAUSE%"=="1" pause
exit /b %EXIT_CODE%
