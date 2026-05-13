@echo off
REM ============================================================
REM  DocMine — 빌드 + 압축 (개발자용)
REM ------------------------------------------------------------
REM  더블클릭하거나 프로젝트 루트에서 `build` 라고 치면
REM  PyInstaller 로 dist\docmine.exe 를 새로 만들고 곧장
REM  release\docmine_v<버전>\ 와 docmine_v<버전>.zip 까지 생성한다.
REM
REM  버전은 src\docmine\__init__.py 의 __version__ 을 그대로 사용.
REM ============================================================

setlocal
cd /d "%~dp0"

echo.
echo === [1/2] PyInstaller 빌드 ===
pyinstaller docmine.spec --clean --noconfirm
if errorlevel 1 (
    echo.
    echo [실패] PyInstaller 빌드 오류.
    exit /b 1
)

echo.
echo === [2/2] 릴리즈 zip 패키징 ===
python make_release.py
if errorlevel 1 (
    echo.
    echo [실패] make_release.py 오류.
    exit /b 1
)

echo.
echo === 완료 ===
endlocal
