@echo off
chcp 65001 > nul
echo ============================================
echo  NotebookLM 구글 로그인
echo ============================================
echo.
set PLAYWRIGHT_BROWSERS_PATH=C:\playwright-browsers
set NOTEBOOKLM_HOME=C:\notebooklm
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
echo 브라우저 경로: %PLAYWRIGHT_BROWSERS_PATH%
echo 데이터 경로: %NOTEBOOKLM_HOME%
echo.
echo 잠시 후 Chromium 브라우저가 열립니다.
echo 구글 계정으로 로그인해 주세요.
echo.
notebooklm login
echo.
if %ERRORLEVEL% EQU 0 (
    echo ===== 로그인 성공! =====
) else (
    echo ===== 로그인 실패. 에러 코드: %ERRORLEVEL% =====
)
echo.
pause
