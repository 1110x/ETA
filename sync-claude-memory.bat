@echo off
chcp 65001 >nul
setlocal

REM ── ETA Claude Code 메모리 동기화 스크립트 ──
REM 이 스크립트를 ETA 프로젝트 폴더에서 한번 실행하면
REM .claude\memory\ 파일들이 Claude Code 메모리 경로로 복사됩니다.

set "PROJECT_DIR=%cd%"
set "MEMORY_SRC=%PROJECT_DIR%\.claude\memory"

REM 프로젝트 경로를 Claude 메모리 경로 형식으로 변환 (C:\Users\foo\ETA → -C-Users-foo-ETA)
set "CONVERTED=%PROJECT_DIR:\=-%"
set "CONVERTED=-%CONVERTED::=%"

set "MEMORY_DST=%USERPROFILE%\.claude\projects\%CONVERTED%\memory"

echo.
echo [ETA] Claude Code 메모리 동기화
echo ─────────────────────────────────
echo  소스: %MEMORY_SRC%
echo  대상: %MEMORY_DST%
echo.

if not exist "%MEMORY_SRC%" (
    echo [오류] .claude\memory 폴더가 없습니다. ETA 프로젝트 루트에서 실행하세요.
    pause
    exit /b 1
)

if not exist "%MEMORY_DST%" (
    mkdir "%MEMORY_DST%"
    echo [생성] %MEMORY_DST%
)

xcopy /Y /Q "%MEMORY_SRC%\*" "%MEMORY_DST%\" >nul
echo [완료] 메모리 파일 %MEMORY_SRC% → %MEMORY_DST% 복사 완료
echo.
echo 이제 이 폴더에서 Claude Code를 실행하면 메모리가 적용됩니다.
echo.
pause
