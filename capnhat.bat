@echo off
chcp 65001 >nul
rem ==== CAP NHAT DU LIEU CSKH ====
rem Cach dung: keo tha file Excel moi vao file nay (hoac nhay dup de ma hoa lai file trong data\)
cd /d "%~dp0"
where node >nul 2>nul || (echo Chua cai Node.js - tai tai https://nodejs.org & pause & exit /b 1)
node tools\update.js %1
echo.
pause
