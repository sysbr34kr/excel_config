@echo off

:: Ensure the script runs in the same directory as the .bat file
cd /d %~dp0

:: Execute quotes.py
echo Running quotes.py...
python quotes.py
if errorlevel 1 (
    echo Failed to execute quotes.py.
    pause
    exit /b
)

echo quotes.py executed successfully!
pause
