@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
echo ============================================
echo 스터디 요약 자동화 실행
echo ============================================
echo.

cd /d "%~dp0"

REM Python 가상환경 활성화 (있는 경우)
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
)

REM 메인 스크립트 실행
python main.py

echo.
echo ============================================
echo 실행 완료
echo ============================================

REM 오류 발생 시 일시 정지 (수동 실행 시 확인용)
if %ERRORLEVEL% NEQ 0 (
    echo 오류가 발생했습니다. 로그를 확인하세요.
    pause
)
