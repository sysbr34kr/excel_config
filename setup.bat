@echo off

:: Verify Python installation
python --version >nul 2>&1
if errorlevel 1 (
    echo Python not installed.
    pause
    exit /b
)

:: Upgrade pip
echo Upgrading pip...
python -m ensurepip --upgrade >nul 2>&1
python -m pip install --upgrade pip >nul 2>&1
if errorlevel 1 (
    echo Failed to upgrade pip.
    pause
    exit /b
)

:: Install required Python packages
echo Installing required packages...
python -m pip install pandas openpyxl lxml
if errorlevel 1 (
    echo Failed to install required packages.
    pause
    exit /b
)

:: Verify lxml installation
echo Verifying lxml installation...
python -c "import lxml; print('lxml is installed, version:', lxml.__version__)" || (
    echo lxml installation failed. Please install it manually.
    pause
    exit /b
)

echo All dependencies installed successfully!
pause
