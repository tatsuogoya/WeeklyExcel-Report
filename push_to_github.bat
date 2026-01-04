@echo off
echo Backing up to GitHub...
echo.

cd /d "%~dp0"

git add .
if %errorlevel% neq 0 (
    echo ERROR: Failed to add files
    pause
    exit /b
)

echo Files staged successfully
echo.
echo Enter your commit message (what changed?):
set /p message=

if "%message%"=="" (
    echo ERROR: Commit message cannot be empty
    pause
    exit /b
)

git commit -m "%message%"
if %errorlevel% neq 0 (
    echo ERROR: Failed to commit
    pause
    exit /b
)

echo.
echo Pushing to GitHub...
git push
if %errorlevel% neq 0 (
    echo ERROR: Failed to push
    pause
    exit /b
)

echo.
echo ==============================
echo   SUCCESS! Backed up to GitHub
echo ==============================
echo.
pause

