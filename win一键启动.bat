@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo CSV batch converter
echo Folder: %SCRIPT_DIR%
echo.
echo Opening the CSV batch converter window...
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 csv_batch_convert_gui.py
    goto done
)

where python >nul 2>nul
if %errorlevel%==0 (
    python csv_batch_convert_gui.py
    goto done
)

where python3 >nul 2>nul
if %errorlevel%==0 (
    python3 csv_batch_convert_gui.py
    goto done
)

echo Error: Python is not installed or not available in PATH.
echo Please install Python 3, then run this script again.

:done
echo.
echo Window closed. Press any key to close this window.
pause >nul
endlocal
