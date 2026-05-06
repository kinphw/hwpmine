@echo off
chcp 65001 >nul
setlocal

REM ─────────────────────────────────────────────────────────────
REM  HWP Mine — 설치 스크립트
REM  같은 폴더에 있는 hwpmine-*.whl 파일을 pip으로 설치합니다.
REM ─────────────────────────────────────────────────────────────

set "HERE=%~dp0"
set "WHL="

for %%F in ("%HERE%hwpmine-*-py3-none-any.whl") do set "WHL=%%F"

if not defined WHL (
    echo [오류] 이 폴더에서 hwpmine-*.whl 파일을 찾을 수 없습니다.
    echo        install.bat 과 같은 위치에 wheel 파일을 두고 다시 실행하세요.
    pause
    exit /b 1
)

where python >nul 2>nul
if errorlevel 1 (
    echo [오류] Python 이 설치되어 있지 않거나 PATH 에 없습니다.
    echo        https://www.python.org 에서 Python 3.10 이상을 먼저 설치하세요.
    pause
    exit /b 1
)

echo 설치 파일: %WHL%
echo.

python -m pip install --force-reinstall --upgrade "%WHL%"
if errorlevel 1 (
    echo.
    echo [오류] 설치 실패. 위 메시지를 확인하세요.
    pause
    exit /b 1
)

echo.
echo ==============================================================
echo   설치 완료.
echo.
echo   다음 단계:
echo     1) .env.example 을 .env 로 복사 후 DB 접속정보 등을 채우세요.
echo     2) 실행: hwpmine       (대화형 메뉴 — 1/2/3/4 선택)
echo        예)   hwpmine 3     (검색 GUI 바로 실행)
echo ==============================================================
pause
endlocal
