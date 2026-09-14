@echo off
chcp 65001 >nul
echo جاري إيقاف سيرفر النظام...
taskkill /f /im python.exe >nul 2>&1
echo تم إيقاف السيرفر بنجاح.
timeout /t 3
