@echo off
title 启动投行工作台
cd /d "%~dp0"
netstat -ano | findstr ":5000 " | findstr "LISTENING" >nul 2>nul
if %errorlevel%==0 (
    echo [提示] 工作台已在运行: http://127.0.0.1:5000
    goto END
)
if not exist ".venv\Scripts\pythonw.exe" (
    echo [错误] 未找到虚拟环境，请先运行「一键安装并启动.bat」
    goto END
)
start "" ".venv\Scripts\pythonw.exe" app.py
echo 正在后台启动，请稍候...
ping -n 5 127.0.0.1 >nul
netstat -ano | findstr ":5000 " | findstr "LISTENING" >nul 2>nul
if %errorlevel%==0 (
    echo [成功] 工作台已启动: http://127.0.0.1:5000
    echo         外网固定地址: https://testh.cpolar.cn
) else (
    echo [失败] 未检测到 5000 端口，可用「启动工作台.bat」前台启动查看报错
)
:END
echo.
pause
