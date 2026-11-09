@echo off
cd /d "E:\WorkBuddy\投行项目\投行工作台"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" "cpolar_start.py"
) else (
    py -3 "cpolar_start.py"
)
echo.
pause
