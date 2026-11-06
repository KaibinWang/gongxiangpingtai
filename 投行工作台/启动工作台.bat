@echo off
title TouHang Workbench
cd /d "%~dp0"

rem ---- 检查 5000 端口是否已被占用（避免重复启动两个实例）----
netstat -ano | findstr ":5000 " | findstr "LISTENING" >nul 2>nul
if %errorlevel%==0 (
    echo [提示] 5000 端口已被占用，工作台可能已经在运行。
    echo        请直接访问  http://127.0.0.1:5000
    echo        如需重启，请先关闭已运行的工作台窗口。
    echo.
    pause
    exit /b 0
)

rem ---- 检查虚拟环境 ----
if not exist ".venv\Scripts\python.exe" (
    echo [错误] 未找到虚拟环境 .venv
    echo        请先双击「一键安装并启动.bat」完成首次安装。
    echo.
    pause
    exit /b 1
)

rem ---- 打开浏览器并启动服务 ----
start "" http://127.0.0.1:5000
echo 工作台启动中，浏览器访问 http://127.0.0.1:5000
echo 关闭本窗口即停止服务。
echo.
".venv\Scripts\python.exe" app.py
pause
