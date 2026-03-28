@echo off
:: ============================================================
:: HRMS AUTO-START ON WINDOWS BOOT
:: Yeh script Windows startup mein add karti hai HRMS ko
:: ============================================================

echo.
echo EasySource HRMS – Windows Startup Setup
echo ========================================
echo.

set "HRMS_DIR=%~dp0"
set "STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "BAT_FILE=%HRMS_DIR%START_HRMS_SERVER.bat"
set "VBS_FILE=%STARTUP_FOLDER%\EasySource_HRMS.vbs"

echo Creating hidden startup script...

:: Create VBS wrapper (runs BAT hidden, no CMD window flash)
echo Set WshShell = CreateObject("WScript.Shell") > "%VBS_FILE%"
echo WshShell.Run """%BAT_FILE%""", 1, False >> "%VBS_FILE%"

echo.
echo ✓ Auto-start installed!
echo.
echo HRMS ab Windows start hone par automatically chalega.
echo.
echo Startup folder: %STARTUP_FOLDER%
echo.

:: Also create Desktop shortcut
set "DESKTOP=%USERPROFILE%\Desktop"
set "LNK_FILE=%DESKTOP%\HRMS Server.lnk"

powershell -Command "$s=(New-Object -COM WScript.Shell).CreateShortcut('%LNK_FILE%'); $s.TargetPath='%BAT_FILE%'; $s.WorkingDirectory='%HRMS_DIR%'; $s.Description='EasySource HRMS Server'; $s.Save()"

echo ✓ Desktop shortcut bana diya: "HRMS Server"
echo.
echo ========================================
echo Setup complete!
echo.
echo Abhi server start karne ke liye:
echo   Desktop par "HRMS Server" double-click karein
echo.
pause
