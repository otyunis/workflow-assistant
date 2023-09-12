@echo off

:: Initialize Log File
echo Log started on %date% %time% > log.txt

:: Set Python Path
set PYTHON_PATH=C:\Users\omart\AppData\Local\Programs\Python\Python311\python.exe

:: Check for Python Version
for /f "tokens=* delims= " %%i in ('%PYTHON_PATH% --version') do (
    set PYTHON_VERSION=%%i
)
echo Checked Python version: %PYTHON_VERSION% >> log.txt

:: Validate Python Version
if not "%PYTHON_VERSION%"=="Python 3.11.1" (
    echo Incorrect Python version found. Required is 3.11.1. >> log.txt
    exit /b 1
)

:: Create virtual environment if doesn't already exist
if not exist ".venv\" (
    %PYTHON_PATH% -m venv .venv
    echo Created virtual environment. >> log.txt
)

:: Activate virtual environment
call .venv\Scripts\activate
echo Activated virtual environment. >> log.txt

:: Install Dependencies
pip install -r requirements.txt >> log.txt
echo Installed dependencies from requirements.txt. >> log.txt

:: Run Application
python application.py >> log.txt 2>&1
echo Ran application.py. >> log.txt

:: Deactivate virtual environment and Exit
deactivate
echo Deactivated virtual environment. >> log.txt
exit /b 0
