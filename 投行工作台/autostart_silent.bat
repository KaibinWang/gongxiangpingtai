@echo off
cd /d "E:\WorkBuddy\投行项目\投行工作台"
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "cpolar_start.py" --no-open
) else (
    start "" py -3 "cpolar_start.py" --no-open
)
exit
