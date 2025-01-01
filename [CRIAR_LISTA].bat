@echo off

:: Ensure the script runs in the same directory as the .bat file
cd /d %~dp0

:: Execute list.py
echo Running list.py...
python list.py
if errorlevel 1 (
    echo Failed to execute list.py.
    pause
    exit /b
)

echo list.py executed successfully!
pause
