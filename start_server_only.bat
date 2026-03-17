@echo off
chcp 65001 > nul
set WORKDIR=C:\Users\bear\Desktop\code\line-cs-bot

echo 正在背景啟動 LINE Bot Server...
wscript.exe "%WORKDIR%\start_server_bg.vbs"
echo 完成！Server 已在背景執行，不會出現在工作列。
echo 日誌：%WORKDIR%\server.log
timeout /t 3 > nul
