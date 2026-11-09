@echo off
title 关闭投行工作台
setlocal enabledelayedexpansion
set FOUND=0
echo ============================================
echo   停止投行工作台 + 外网访问
echo ============================================
echo.
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5000 " ^| findstr "LISTENING"') do (
    taskkill /F /PID %%a >nul 2>&1
    set FOUND=1
)
taskkill /F /IM cpolar.exe >nul 2>&1
netstat -ano | findstr ":5000 " | findstr "LISTENING" >nul 2>nul
if %errorlevel%==0 (
    echo [警告] 5000 端口仍被占用，请检查是否有其他程序占用
) else (
    echo [完成] 工作台已停止，外网隧道已关闭
    echo        浏览器如仍显示旧页面，请按 Ctrl+F5 强制刷新
)
echo.
pause
