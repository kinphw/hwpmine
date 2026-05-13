@echo off
chcp 65001 >nul
setlocal EnableExtensions

REM ─────────────────────────────────────────────────────────────
REM  DocMine — 설치 스크립트 (단일 실행파일 버전)
REM
REM  같은 폴더에 있는 docmine.exe 와 .env.example 을
REM  %LOCALAPPDATA%\Programs\docmine\ 에 복사하고
REM  사용자 PATH 에 등록합니다.
REM ─────────────────────────────────────────────────────────────

set "HERE=%~dp0"
set "EXE=%HERE%docmine.exe"
set "ENV_EXAMPLE=%HERE%.env.example"
set "TARGET=%LOCALAPPDATA%\Programs\docmine"

if not exist "%EXE%" (
    echo [오류] 같은 폴더에서 docmine.exe 를 찾을 수 없습니다.
    echo        install.bat 과 같은 위치에 docmine.exe 를 두고 다시 실행하세요.
    pause
    exit /b 1
)

echo 설치 위치: %TARGET%
echo 실행파일 : %EXE%
echo.

if not exist "%TARGET%" mkdir "%TARGET%" >nul 2>nul
if errorlevel 1 (
    echo [오류] 설치 폴더를 생성하지 못했습니다: %TARGET%
    pause
    exit /b 1
)

copy /Y "%EXE%" "%TARGET%\docmine.exe" >nul
if errorlevel 1 (
    echo [오류] docmine.exe 복사에 실패했습니다. 다른 docmine.exe 가 실행 중인지 확인하세요.
    pause
    exit /b 1
)

if exist "%ENV_EXAMPLE%" (
    copy /Y "%ENV_EXAMPLE%" "%TARGET%\.env.example" >nul
)

REM ── 사용자 PATH 에 설치 폴더 추가 (이미 있으면 건너뜀) ──
set "INPATH="
for %%P in ("%PATH:;=";"%") do (
    if /I "%%~P"=="%TARGET%" set "INPATH=1"
)

if defined INPATH (
    echo PATH 에 이미 등록되어 있습니다.
) else (
    REM 현재 사용자 PATH 만 안전하게 읽어서 setx 로 갱신
    for /f "tokens=2*" %%A in ('reg query "HKCU\Environment" /v Path 2^>nul ^| findstr /R /C:"REG_SZ" /C:"REG_EXPAND_SZ"') do set "USER_PATH=%%B"
    if not defined USER_PATH (
        setx Path "%TARGET%" >nul
    ) else (
        setx Path "%USER_PATH%;%TARGET%" >nul
    )
    if errorlevel 1 (
        echo [경고] 사용자 PATH 등록에 실패했습니다. 수동으로 추가하세요: %TARGET%
    ) else (
        echo 사용자 PATH 에 추가했습니다 (새 콘솔에서 적용됨).
    )
)

echo.
echo ==============================================================
echo   설치 완료.
echo.
echo   다음 단계:
echo     1) 사용자 .env 위치 — 권장: %APPDATA%\docmine\.env
echo        (또는 docmine 을 실행할 작업 폴더의 .env)
echo        설치 폴더의 .env.example 을 복사해 값을 채우세요.
echo.
echo     2) 새 콘솔(또는 탐색기)을 열고 실행:
echo          docmine          ^(대화형 메뉴^)
echo          docmine 3        ^(검색 GUI 바로 실행^)
echo ==============================================================
pause
endlocal
