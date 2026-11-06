@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo [错误] 未找到虚拟环境 .venv
    echo        请先双击「一键安装并启动.bat」完成首次安装。
    echo.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" 启动并穿透.py
pause
