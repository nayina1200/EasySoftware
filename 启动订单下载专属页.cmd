@echo off
chcp 65001 >nul
setlocal
set "ORDER_DOWNLOAD_SCRIPT=%~dp0scripts\order_material_downloader_gui.ps1"

if not exist "%ORDER_DOWNLOAD_SCRIPT%" (
    echo 找不到订单资料下载页面：%ORDER_DOWNLOAD_SCRIPT%
    echo.
    pause
    exit /b 2
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%ORDER_DOWNLOAD_SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" echo 订单下载页面异常退出，退出代码：%EXIT_CODE%
echo.
echo 订单资料下载专属页已关闭。按任意键退出。
pause >nul
exit /b %EXIT_CODE%
