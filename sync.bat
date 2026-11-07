@echo off
setlocal
cd /d "%~dp0" || goto fail

echo Current path: %cd%

rem Keep the local Excel file out of the commit.
git update-index --assume-unchanged option_activity_log.xlsx
if errorlevel 1 goto fail

git add .
if errorlevel 1 goto fail

git reset -- option_activity_log.xlsx >nul 2>&1
rem A missing/untracked Excel file is fine.

git diff --cached --quiet
if errorlevel 2 goto fail
if errorlevel 1 (
    git commit -m "sync on %date% %time%"
    if errorlevel 1 goto fail
) else (
    echo No new changes to commit.
)

git pull origin main --rebase
if errorlevel 1 goto fail

git push origin main
if errorlevel 1 goto fail

echo Sync successful.
exit /b 0

:fail
echo.
echo ERROR: Sync failed. Review the messages above.
pause
exit /b 1
