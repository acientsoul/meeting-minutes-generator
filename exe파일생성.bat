@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo ================================================
echo 회의록 자동 생성 프로그램 - exe 파일 생성 중
echo ================================================
echo.

cd /d "%~dp0"

echo 1/3: PyInstaller 설치 중...
python -m pip install pyinstaller --quiet
if !errorlevel! neq 0 (
    echo 오류: PyInstaller 설치 실패
    pause
    exit /b 1
)
echo ✓ PyInstaller 설치 완료

echo.
echo 2/3: exe 파일 생성 중... (이 과정은 1-2분 소요)
timeout /t 2 /nobreak
python -m PyInstaller --onefile --windowed --name "회의록자동생성" main.py

if !errorlevel! neq 0 (
    echo 오류: exe 파일 생성 실패
    pause
    exit /b 1
)
echo ✓ exe 파일 생성 완료

echo.
echo 3/3: 정리 중...
if exist "build" rmdir /s /q "build" >nul 2>&1
if exist "회의록자동생성.spec" del "회의록자동생성.spec" >nul 2>&1

echo.
echo ================================================
echo ✓ 완료!
echo ================================================
echo.
echo 생성된 파일: dist\회의록자동생성.exe
echo 이 exe 파일을 더블클릭하면 프로그램이 실행됩니다.
echo.
pause
