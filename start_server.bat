@echo off
echo ============================================
echo   EasySource HRMS Enterprise Server
echo ============================================
echo.
cd /d "%~dp0"

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.8+
    pause
    exit /b
)

:: Install dependencies
echo Installing dependencies...
pip install flask --quiet

:: Start server
echo.
echo Starting HRMS Server...
echo Open your browser at: http://localhost:5000
echo.
echo Default Login:
echo   Admin: admin / admin123
echo   User:  user1 / user123
echo.
echo Press Ctrl+C to stop the server
echo ============================================
python app.py
pause
