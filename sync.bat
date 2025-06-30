@echo off
cd /d E:\Investment\daily_top_options
git add .
git commit -m "sync on %date% %time%"
git push origin main
pause
