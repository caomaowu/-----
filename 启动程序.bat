@echo off
cd /d "%~dp0"
python gui_main.py
if %errorlevel% neq 0 (
    echo.
    echo Program exited with error code %errorlevel%.
    pause
)
