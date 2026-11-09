@echo off
title 启动投行工作台
cd /d "%~dp0"
echo ============================================
echo   启动投行工作台 + 外网固定域名
echo ============================================
echo.
if not exist ".venv\Scripts\pythonw.exe" (
    echo [错误] 未找到虚拟环境，请先运行「一键安装并启动.bat」
    goto END
)
netstat -ano | findstr ":5000 " | findstr "LISTENING" >nul 2>nul
if %errorlevel%==0 (
    echo [1/2] 工作台已在运行
) else (
    start "" ".venv\Scripts\pythonw.exe" app.py
    echo [1/2] 工作台后台启动中...
    ping -n 5 127.0.0.1 >nul
)
tasklist /FI "IMAGENAME eq cpolar.exe" | findstr /i "cpolar.exe" >nul 2>nul
if %errorlevel%==0 (
    echo [2/2] 外网隧道已在运行
) else (
    if exist "..\cpolar\cpolar.exe" (
        start "" "..\cpolar\cpolar.exe" start-all
        echo [2/2] 外网隧道启动中...
        ping -n 8 127.0.0.1 >nul
    ) else (
        echo [2/2] 未找到 cpolar.exe，外网地址暂不可用
    )
)
netstat -ano | findstr ":5000 " | findstr "LISTENING" >nul 2>nul
if %errorlevel%==0 (
    echo.
    echo [成功] 本机: http://127.0.0.1:5000
    echo         外网: https://testh.cpolar.cn  (隧道约需 20-30 秒生效)
) else (
    echo.
    echo [失败] 5000 端口未就绪，请用「启动工作台.bat」前台启动查看报错
)
:END
echo.
pause
