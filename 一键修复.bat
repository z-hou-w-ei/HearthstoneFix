@echo off
chcp 936 >nul
title 炉石一键修复
setlocal
cd /d "%~dp0"

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" goto nops
if not exist "%~dp0HearthstoneFix.ps1" goto nosec

fltmc.exe >nul 2>&1
if not errorlevel 1 goto run

echo [提权] 正在请求管理员权限，请在弹窗中选“是”...
set "ARGS=%*"
if "%ARGS%"=="" goto elev_bare
"%PS%" -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -WorkingDirectory '%~dp0' -Verb RunAs -ArgumentList '%ARGS%'" >nul 2>&1
goto elev_done

:elev_bare
"%PS%" -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -WorkingDirectory '%~dp0' -Verb RunAs" >nul 2>&1

:elev_done
if errorlevel 1 goto noelev
exit /b 0

:noelev
echo.
echo [提示] 未获得管理员权限，改用普通权限继续（重启战网服务、SYSTEM 强杀会跳过）。
echo.

:run
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

:nops
echo [错误] 找不到 Windows PowerShell，无法运行本脚本。
pause
exit /b 1

:nosec
echo [错误] 找不到 HearthstoneFix.ps1，请确认它和本文件在同一个目录。
pause
exit /b 1
