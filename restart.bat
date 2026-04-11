@echo off
chcp 65001 > nul
set PYTHON=C:\Users\bear\AppData\Local\Programs\Python\Python312\python.exe
set WORKDIR=C:\Users\bear\Desktop\code\line-cs-bot

echo === 關閉舊的 Bot 程序 ===
taskkill /F /IM python.exe /T > nul 2>&1
timeout /t 2 > nul

echo === 啟動 Bot Server ===
start "LINE Bot Server" cmd /k "cd /d %WORKDIR% && %PYTHON% main.py"

timeout /t 3 > nul

echo 完成！Bot 已啟動。
echo Webhook URL: https://xmnline.duckdns.org/webhook
