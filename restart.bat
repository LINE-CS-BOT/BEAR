@echo off
chcp 65001 > nul
set PYTHON=C:\Users\bear\AppData\Local\Programs\Python\Python312\python.exe
set WORKDIR=C:\Users\bear\Desktop\code\line-cs-bot
set CHROME="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

echo === 關閉舊的 Bot 程序 ===
taskkill /F /IM python.exe /T > nul 2>&1
timeout /t 2 > nul

echo === 啟動 Bot Server ===
start "LINE Bot Server" cmd /k "cd /d %WORKDIR% && %PYTHON% main.py"

timeout /t 3 > nul

:: 確認 LINE OA Chrome 是否在跑（port 9223）
curl -s http://127.0.0.1:9223/json/version >nul 2>&1
if errorlevel 1 (
    echo === 啟動 LINE OA Chrome ===
    start "" %CHROME% --remote-debugging-port=9223 --user-data-dir="%WORKDIR%\data\line_chrome_session" --no-first-run --disable-default-apps --window-position=-32000,-32000 --window-size=1400,900 "https://chat.line.biz/"
) else (
    echo LINE OA Chrome 已在執行中
)

echo 完成！Bot 已啟動。
echo Webhook URL: https://xmnline.duckdns.org/webhook
