@echo off
cd /d E:\Investment\daily_top_options

:: 自动拉取远程代码并尝试合并
git pull origin main --rebase

:: 自动添加和提交更改
git add .
git commit -m "sync on %date% %time%" 2>nul

:: 自动推送到 GitHub
git push origin main

pause
