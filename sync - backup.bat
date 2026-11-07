@echo off
cd /d "%~dp0"
echo ✅ 当前路径：%cd%

:: ✅ 告诉 Git：永远别再管 option_activity_log.xlsx（只需执行一次）
git update-index --assume-unchanged option_activity_log.xlsx

:: ✅ 提交所有文件，排除 Excel（reset 确保不会意外加入）
git add .
git reset option_activity_log.xlsx >nul 2>&1
git commit -m "sync on %date% %time%" 2>nul

:: ✅ 拉远程代码，不影响本地 Excel
git pull origin main --rebase

:: ✅ 推送（不包含 Excel）
git push origin main

pause