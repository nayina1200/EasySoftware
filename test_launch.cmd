@echo off
chcp 65001 >nul
set "EASYSOFTWARE_ROOT=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%EASYSOFTWARE_ROOT%启动EasySoftware.ps1" %*
echo.
echo 处理窗口即将关闭，按任意键退出。
pause >nul
