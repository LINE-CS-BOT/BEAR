@echo off
chcp 65001 > nul
cd /d C:\Users\bear\Desktop\code\line-cs-bot

:: Check if there are changes
git diff --quiet --exit-code 2>nul
if %errorlevel%==0 (
    git diff --cached --quiet --exit-code 2>nul
    if %errorlevel%==0 (
        :: No changes, skip
        exit /b 0
    )
)

:: Stage all tracked files (not untracked)
git add -u

:: Commit with timestamp
for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set d=%%a-%%b-%%c
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set t=%%a:%%b
git commit -m "auto-backup %d% %t%" >nul 2>&1

:: Push
git push origin master >nul 2>&1
