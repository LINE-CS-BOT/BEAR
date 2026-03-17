@echo off
set PYTHON=C:\Users\bear\AppData\Local\Programs\Python\Python312\pythonw.exe
set WORKDIR=C:\Users\bear\Desktop\code\line-cs-bot

cd /d %WORKDIR%
start "" %PYTHON% tray.py
