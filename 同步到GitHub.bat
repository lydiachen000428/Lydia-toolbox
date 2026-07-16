@echo off
chcp 65001 >nul
cd /d "D:\user\lydia_chen\Documents\packing-app"
echo.
echo Syncing to GitHub...
echo.
git add .
set /p msg=Commit message (e.g. fix CBP):
git commit -m "%msg%"
git push origin HEAD:main
echo.
echo Done! Press any key to close.
pause >nul
