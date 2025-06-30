@echo off
cd /d "%~dp0"
echo ✅ 当前路径：%cd%

:: 恢复未提交的 Excel 文件（跳过 sync.bat 自身）
git restore option_activity_log.xlsx 2>nul

:: 拉取远程代码前，先提交当前 sync.bat 变更
git add sync.bat
git add *.py
git add *.yml
git add *.txt
if exist .gitignore git add .gitignore
git commit -m "sync on %date% %time%" 2>nul

:: 拉取远程更新
git pull origin main --rebase

:: 推送
git push origin main

pause
