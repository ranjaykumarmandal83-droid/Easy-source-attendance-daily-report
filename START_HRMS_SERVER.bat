@echo off
title EasySource HRMS Server
color 0A
cls

echo.
echo  ████████████████████████████████████████████████
echo  █                                              █
echo  █    EasySource HRMS Enterprise Server        █
echo  █    Laptop Server + Mobile Access            █
echo  █                                              █
echo  ████████████████████████████████████████████████
echo.

:: Get local IP
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4 Address"') do (
    set LOCAL_IP=%%a
    goto :gotip
)
:gotip
set LOCAL_IP=%LOCAL_IP: =%

echo  [1/4] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python not found!
    echo  Download from: https://python.org
    pause & exit /b
)
echo  ✓ Python OK

echo  [2/4] Installing dependencies...
pip install flask --quiet
echo  ✓ Flask OK

echo  [3/4] Starting HRMS Server...
cd /d "%~dp0"
start "HRMS Server" /MIN python app.py
timeout /t 3 /nobreak >nul
echo  ✓ Server Started

echo  [4/4] Checking ngrok for Internet access...
ngrok version >nul 2>&1
if errorlevel 1 (
    echo  ℹ ngrok not found - Only LAN access available
    echo  For Internet access, download ngrok from: https://ngrok.com/download
) else (
    echo  ✓ Starting ngrok tunnel...
    start "ngrok" /MIN ngrok http 5000
    timeout /t 4 /nobreak >nul
)

cls
echo.
echo  ████████████████████████████████████████████████
echo  █         SERVER IS RUNNING ✓                  █
echo  ████████████████████████████████████████████████
echo.
echo  ┌─────────────────────────────────────────────┐
echo  │  📍 LOCAL ACCESS (Same WiFi)               │
echo  │     http://%LOCAL_IP%:5000           │
echo  │     Mobile: http://%LOCAL_IP%:5000/mobile  │
echo  └─────────────────────────────────────────────┘
echo.
echo  ┌─────────────────────────────────────────────┐
echo  │  🌐 INTERNET ACCESS (ngrok)                │
echo  │     Check ngrok window for URL             │
echo  │     OR visit: http://127.0.0.1:4040        │
echo  └─────────────────────────────────────────────┘
echo.
echo  ┌─────────────────────────────────────────────┐
echo  │  🔑 LOGIN CREDENTIALS                      │
echo  │     Admin:  admin   / admin123             │
echo  │     User:   user1   / user123              │
echo  └─────────────────────────────────────────────┘
echo.
echo  📱 EMPLOYEES KO YE URL SHARE KAREIN:
echo     http://%LOCAL_IP%:5000/mobile
echo.
echo  Press any key to open browser...
pause >nul
start http://localhost:5000/dashboard

echo.
echo  Server running... Is window band mat karein!
echo  Band karne ke liye Ctrl+C dabayein.
echo.
pause
