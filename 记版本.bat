@echo off
cd /d "%~dp0"
echo ============================================
echo   给当前项目记一个版本（可回退 / 可追溯）
echo ============================================
echo.
set /p MSG=这轮改了什么（直接回车用默认说明）: 
if "%MSG%"=="" set MSG=手动快照
echo.
"%~dp0投行工作台\.venv\Scripts\python.exe" "%~dp0版本快照.py" -m "%MSG%"
echo.
pause
