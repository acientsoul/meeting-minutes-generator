@echo off
chcp 65001 > nul 2>&1
cd /d "%~dp0"

python gui_stable.py
pause

if errorlevel 1 (
    echo.
    echo ❌ 오류가 발생했습니다!
    echo 위의 메시지를 확인하세요.
    echo.
)

pause
