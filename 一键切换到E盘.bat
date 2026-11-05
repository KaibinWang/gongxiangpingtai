@echo off
chcp 936 >nul
cd /d "%~dp0"
title 切换到 E 盘工作区

echo ============================================================
echo   把 WorkBuddy 的工作区从 C 盘切到 E 盘
echo   C: 旧目录（103 MB）空闲出来
echo ============================================================
echo.
echo 即将关闭 WorkBuddy，改写工作区路径后自动重新打开。
echo 关闭前请确认没有正在跑的任务。
echo.
pause

echo.
echo [1/4] 关闭 WorkBuddy ...
taskkill /IM WorkBuddy.exe /F >nul 2>&1
timeout /t 3 /nobreak >nul

echo [2/4] 等待进程完全退出 ...
set /a n=0
:waitloop
tasklist /FI "IMAGENAME eq WorkBuddy.exe" 2>nul | find /I "WorkBuddy.exe" >nul
if not errorlevel 1 (
  set /a n+=1
  if %n% GEQ 15 goto continue
  timeout /t 1 /nobreak >nul
  goto waitloop
)
:continue

echo [3/4] 改写工作区路径 ...
"C:/Users/10949/.workbuddy/binaries/python/versions/3.13.12/python.exe" "%~dp0切换工作区.py"
if errorlevel 1 (
  echo.
  echo [X] 改写失败，WorkBuddy 将以原路径启动。
  pause
)

echo.
echo [4/4] 重新打开 WorkBuddy ...
start "" "C:/Users/10949/AppData/Local/Programs/WorkBuddy/WorkBuddy.exe"
timeout /t 3 /nobreak >nul

echo.
echo 完成。WorkBuddy 重新打开后，本对话应在 E:/WorkBuddy/投行项目 下继续。
echo 窗口可以关闭了。
timeout /t 8 >nul
