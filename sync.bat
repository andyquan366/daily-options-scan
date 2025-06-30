@echo off
cd /d "%~dp0"
echo ✅ 当前路径：%cd%

:: 如果 Excel 有修改，先 stash（避免 pull 冲突）
git diff --quiet option_activity_log.xlsx
if errorlevel 1 (
    echo ⚠️ 检测到本地修改的 Excel，自动 stash...
    git stash push -m "temp-stash-xlsx" option_activity_log.xlsx
)

:: 拉取远程更新（不会丢你 stash 的文件）
git pull origin main --rebase

:: 恢复你刚才 stash 的 Excel（如果有）
git stash pop >nul 2>&1

:: 添加并提交其他代码
git add *.py
git add *.yml
git add *.txt
git add sync.bat
git add option_activity_log.xlsx
if exist .gitignore git add .gitignore

git commit -m "sync on %date% %time%" 2>nul
git push origin main

pause
