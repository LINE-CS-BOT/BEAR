@echo off
:: 立即手動同步 Ecount 庫存情況 -> data/available.json
:: （需要 Chrome 以 open_chrome_debug.bat 開啟，或 Chrome 完全關閉）

set PYTHON=C:\Users\bear\AppData\Local\Programs\Python\Python312\python.exe
set WORKDIR=C:\Users\bear\Desktop\code\line-cs-bot

echo ======================================
echo  Ecount 庫存情況同步
echo ======================================
cd /d %WORKDIR%
%PYTHON% scripts\auto_sync_unfulfilled.py
echo.
if errorlevel 1 (
    echo [失敗] 同步失敗，請確認：
    echo   1. Chrome 是否以 open_chrome_debug.bat 開啟
    echo   2. 或者關閉 Chrome 再執行
) else (
    echo [完成] available.json 已更新
)
echo.
pause
