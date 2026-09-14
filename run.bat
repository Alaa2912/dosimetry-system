@echo off
chcp 65001 >nul
title سيرفر منظومة قياس الوضحات OSL
echo ========================================================
echo   جاري تشغيل خادم منظومة الرصد الإشعاعي OSL...
echo ========================================================

rem Free port 8000 if previously occupied
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8000"') do (
    taskkill /f /pid %%a >nul 2>&1
)

set PYTHONPATH=%cd%
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
pause
