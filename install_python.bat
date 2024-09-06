@echo off
set PYTHON_URL=https://www.python.org/ftp/python/3.9.9/python-3.9.9-amd64.exe
set INSTALLER_PATH=%TEMP%\python-installer.exe

rem Download Python installer
powershell -Command "Invoke-WebRequest -Uri %PYTHON_URL% -OutFile %INSTALLER_PATH%"

rem Install Python silently
%INSTALLER_PATH% /quiet InstallAllUsers=1 PrependPath=1

rem Clean up installer
del /f %INSTALLER_PATH%

rem Verify installation
python --version