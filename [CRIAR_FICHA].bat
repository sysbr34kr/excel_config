@echo off

:: Ensure the script runs in the same directory as the .bat file
cd /d %~dp0

:: Execute record.py
echo Running record.py...
python record.py
if errorlevel 1 (
    echo Failed to execute record.py.
    pause
    exit /b
)

echo record.py executed successfully!
pause
