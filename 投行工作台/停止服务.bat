@echo off
title 关闭投行工作台
setlocal enabledelayedexpansion
set FOUND=0
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5000 " ^| findstr "LISTENING"') do (
    taskkill /F /PID %%a >nul 2>&1
    set FOUND=1
)
if "!FOUND!"=="1" (
    echo [完成] 工作台已停止。
    echo        注意: cpolar 外网隧道不受影响，仍在运行。
) else (
    echo [提示] 工作台本来就没有在运行。
)
echo.
pause
