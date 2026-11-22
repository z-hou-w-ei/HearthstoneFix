@echo off
chcp 936 >nul
setlocal
cd /d "%~dp0"

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%~dp0HearthstoneFix.ps1" (
    echo [错误] 找不到 HearthstoneFix.ps1，请确认它和本文件在同一个目录。
    pause
    exit /b 1
)

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0HearthstoneFix.ps1" %*
set "RC=%ERRORLEVEL%"

if "%RC%"=="0" goto ok
if "%RC%"=="2" goto ok

echo.
echo [未完成] 返回码 %RC%，窗口保留，按任意键关闭。
pause >nul
exit /b %RC%

:ok
echo.
echo [完成] 窗口将在 3 秒后自动关闭。
timeout /t 3 /nobreak >nul
exit /b %RC%
