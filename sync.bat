@echo off
cd /d "%~dp0"
echo ✅ 当前路径：%cd%

:: 提交当前修改（除了 .xlsx）
git add *.py
git add *.yml
git add *.txt
git add sync.bat
if exist .gitignore git add .gitignore
git commit -m "sync on %date% %time%" 2>nul

:: 如果本地 .xlsx 有改动，先 stash 防止被覆盖
git diff --quiet option_activity_log.xlsx
if errorlevel 1 (
    echo ⚠️ 检测到本地修改的 Excel，自动 stash...
    git stash push -m "stash-xlsx" option_activity_log.xlsx
)

:: 拉远程代码（现在目录干净，不会报错）
git pull origin main --rebase

:: 恢复 Excel（如果有 stash）
git stash pop >nul 2>&1

:: 最后把 .xlsx 加进来一起提交
git add option_activity_log.xlsx
git commit -m "update excel %date% %time%" 2>nul
git push origin main

pause
