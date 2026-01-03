@echo off
echo Starting Weekly Report Generator...
echo ----------------------------------

REM Navigate to the script directory
cd /d "%~dp0"

REM Check for virtual environment
if not exist ".venv" (
    echo Error: .venv virtual environment not found.
    echo Please ensure you have set up the project correctly.
    pause
    exit /b
)

REM Activate venv and run script
echo Running script...
call .venv\Scripts\activate.bat
python generate_report.py

echo.
echo ----------------------------------
echo Done.
pause
