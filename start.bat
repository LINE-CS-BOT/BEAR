@echo off
set PYTHON=C:\Users\bear\AppData\Local\Programs\Python\Python312\python.exe
set WORKDIR=C:\Users\bear\Desktop\code\line-cs-bot
set CADDY=%WORKDIR%\caddy.exe

echo ===================================
echo   LINE 客服 Bot 啟動中...
echo ===================================
echo.

:: 啟動 FastAPI server（新視窗）
start "LINE Bot Server" cmd /k "cd /d %WORKDIR% && %PYTHON% main.py"

:: 等待 server 啟動
timeout /t 3 > nul

echo Bot Server 已啟動在 http://localhost:8000
echo Webhook URL: https://xmnline.duckdns.org/webhook
echo.
echo ===================================
echo   啟動 Caddy HTTPS...
echo ===================================

:: 啟動 Caddy（新視窗）
start "Caddy HTTPS" cmd /k "cd /d %WORKDIR% && %CADDY% run --config Caddyfile"
