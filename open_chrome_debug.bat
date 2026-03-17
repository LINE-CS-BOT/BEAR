@echo off
:: 以 remote debugging port 9222 啟動 Chrome
:: 執行後 auto_sync_unfulfilled.py 可以連接正在執行的 Chrome

echo 關閉現有 Chrome...
taskkill /F /IM chrome.exe >nul 2>&1
timeout /t 2 >nul

echo 啟動 Chrome (remote debugging port 9222)...
start "" "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" ^
  --remote-debugging-port=9222 ^
  --user-data-dir="C:\Users\bear\Desktop\code\line-cs-bot\data\chrome_ecount_session" ^
  https://loginib.ecount.com/ec5/view/erp

echo.
echo Chrome 已啟動，remote debugging port: 9222
echo 現在可以執行：python scripts/auto_sync_unfulfilled.py
echo.
