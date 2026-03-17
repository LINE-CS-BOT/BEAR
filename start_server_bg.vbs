' start_server_bg.vbs
' 在背景啟動 LINE Bot server，完全不顯示任何視窗
Dim WshShell
Dim PYTHON, WORKDIR, CMD

PYTHON  = "C:\Users\bear\AppData\Local\Programs\Python\Python312\python.exe"
WORKDIR = "C:\Users\bear\Desktop\code\line-cs-bot"

Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = WORKDIR

' 先殺掉佔用 8000 port 的舊程序
WshShell.Run "cmd /c for /f ""tokens=5"" %a in ('netstat -ano ^| findstr :8000 ^| findstr LISTENING') do taskkill /F /PID %a >nul 2>&1", 0, True

' 等 1 秒確保 port 釋放
WScript.Sleep 1000

' 用 PowerShell 隱藏視窗啟動（最穩定）
CMD = "powershell.exe -NoProfile -WindowStyle Hidden -Command """ & _
      "Start-Process -FilePath '" & PYTHON & "'" & _
      " -ArgumentList '-m uvicorn main:app --host 0.0.0.0 --port 8000'" & _
      " -WorkingDirectory '" & WORKDIR & "'" & _
      " -WindowStyle Hidden" & _
      """"

WshShell.Run CMD, 0, False
