$PYTHON  = "C:\Users\bear\AppData\Local\Programs\Python\Python312\python.exe"
$WORKDIR = "C:\Users\bear\Desktop\code\line-cs-bot"
$PORT    = 8000

Write-Host "[1] Stop old server..."
$conns = Get-NetTCPConnection -LocalPort $PORT -ErrorAction SilentlyContinue
$killed = @()
foreach ($c in $conns) {
    if ($c.OwningProcess -gt 0 -and $c.OwningProcess -notin $killed) {
        Stop-Process -Id $c.OwningProcess -Force -ErrorAction SilentlyContinue
        Write-Host "    killed PID $($c.OwningProcess)"
        $killed += $c.OwningProcess
    }
}
Start-Sleep 3

# 只檢查 LISTEN / ESTABLISHED（忽略 TIME_WAIT / FIN_WAIT）
$active = Get-NetTCPConnection -LocalPort $PORT -ErrorAction SilentlyContinue |
          Where-Object { $_.State -in @('Listen','Established') }
if ($active) {
    Write-Host "[!] Port $PORT still active. Waiting 5s more..."
    Start-Sleep 5
    $active = Get-NetTCPConnection -LocalPort $PORT -ErrorAction SilentlyContinue |
              Where-Object { $_.State -in @('Listen','Established') }
    if ($active) {
        Write-Host "[!] Port $PORT still in use. Check manually."
        exit 1
    }
}

Write-Host "[2] Start new server..."
Start-Process -FilePath $PYTHON `
    -ArgumentList "-m uvicorn main:app --host 0.0.0.0 --port $PORT" `
    -WorkingDirectory $WORKDIR `
    -WindowStyle Hidden

Write-Host "    Waiting 6s..."
Start-Sleep 6

Write-Host "[3] Verify..."
try {
    $r = Invoke-RestMethod -Uri "http://localhost:$PORT/health" -ErrorAction Stop
    Write-Host "    OK  pid=$($r.pid)  started_at=$($r.started_at)"
} catch {
    Write-Host "    FAIL - server not responding, check logs"
}
