@echo off
chcp 936 >nul
set "PROJ=%~dp0"
set "DST=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\投行工作台自启.bat"
> "%DST%" echo @echo off
>> "%DST%" echo cd /d "%PROJ%"
>> "%DST%" echo start "" /MIN ".venv\Scripts\python.exe" app.py
echo.
echo 已设置开机自启, 脚本位置:
echo   %DST%
echo.
echo 以后开机登录后会自动启动, 直接访问 http://127.0.0.1:5000
echo 如需取消, 删除上面那个 bat 文件即可
echo.
pause
