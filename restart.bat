@echo off
chcp 65001 > nul
set PYTHON=C:\Users\bear\AppData\Local\Programs\Python\Python312\python.exe
set WORKDIR=C:\Users\bear\Desktop\code\line-cs-bot
set CHROME="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

:: ── 程式碼健檢（--skip-check 可略過）─────────────────
if /I "%~1"=="--skip-check" goto :skip_check
echo === 程式碼健檢 (compileall + ruff F821) ===
cd /d %WORKDIR%
%PYTHON% -m compileall -q . > nul
if errorlevel 1 (
    echo.
    echo ❌ 語法錯誤！未重啟 server。修完再跑或加 --skip-check
    %PYTHON% -m compileall .
    pause
    exit /b 1
)
%PYTHON% -m ruff check --select F821 . > nul
if errorlevel 1 (
    echo.
    echo ❌ 有未定義名稱！未重啟 server。以下是詳細位置：
    %PYTHON% -m ruff check --select F821 .
    pause
    exit /b 1
)
echo ✅ 程式碼健檢通過
:skip_check

echo === 關閉舊的 Bot 程序 ===
:: 建 lock 讓 tray.py 的 watchdog 暫停 respawn（避免兩個 python 打架）
:: 用 copy /b 更新 mtime，也保證 lock 存在
echo restart_in_progress > "%WORKDIR%\data\.restart_in_progress.lock"
taskkill /F /IM python.exe /T > nul 2>&1
timeout /t 2 > nul

echo === 啟動 Bot Server ===
start "LINE Bot Server" cmd /k "cd /d %WORKDIR% && %PYTHON% main.py"

timeout /t 3 > nul
:: 新 server 應該已經起來，釋放 lock 讓 watchdog 恢復監看
del "%WORKDIR%\data\.restart_in_progress.lock" > nul 2>&1

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
