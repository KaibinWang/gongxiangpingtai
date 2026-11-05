@echo off
chcp 936 >nul
cd /d "%~dp0"
echo ==========================================
echo    投行部 数智化工作台 - 一键安装并启动
echo ==========================================
echo.
echo [1/4] 检查 Python 环境 ...
where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo [X] 未检测到 Python
    echo     请先安装 Python 3.9 或更高版本 (推荐 3.13)
    echo     安装时务必勾选 "Add Python to PATH"
    echo     下载地址 https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)
python --version
echo.
echo [2/4] 创建虚拟环境 .venv ...
if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
)
if not exist ".venv\Scripts\python.exe" (
    echo [X] 虚拟环境创建失败
    pause
    exit /b 1
)
echo.
echo [3/4] 安装依赖 ...
if exist "_wheels" (
    echo     优先使用离线包 _wheels (无需外网)
    ".venv\Scripts\python.exe" -m pip install --no-index --find-links=_wheels -r requirements.txt
    if errorlevel 1 (
        echo.
        echo     [!] 离线包与当前 Python 版本不匹配, 改为联网安装 ...
        ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    )
) else (
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
)
".venv\Scripts\python.exe" -c "import flask, openpyxl, docx" >nul 2>nul
if errorlevel 1 (
    echo.
    echo [X] 依赖安装失败, 请查看上方错误信息
    echo     若机器无外网, 建议使用 Python 3.13 后重新运行本脚本
    pause
    exit /b 1
)
echo     依赖校验通过
echo.
echo [4/4] 启动工作台 ...
start "" http://127.0.0.1:5000
echo.
echo 服务已启动 !  浏览器访问 http://127.0.0.1:5000
echo 关闭本窗口即停止服务
echo.
".venv\Scripts\python.exe" app.py
pause
