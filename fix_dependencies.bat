@echo off
chcp 65001 > nul
echo == DEPENDENCY REPAIR TOOL ==
echo This script will fix missing dependencies in the virtual environment
echo.

if not exist venv (
  echo Virtual environment not found!
  echo Please run setup_once.bat first to create the virtual environment.
  pause
  exit /b 1
)

call venv\Scripts\activate

echo Installing all required dependencies...

echo Upgrading pip...
venv\Scripts\python.exe -m pip install --upgrade pip

echo Installing Pillow (using wheel - no compilation needed)...
venv\Scripts\pip.exe install --only-binary=:all: Pillow

echo Installing psutil (using wheel - no compilation needed)...
venv\Scripts\pip.exe install --only-binary=:all: psutil

echo Installing PyGetWindow...
venv\Scripts\pip.exe install --only-binary=:all: PyGetWindow

echo Installing PyAutoGUI...
venv\Scripts\pip.exe install --only-binary=:all: PyAutoGUI

echo Installing keyboard...
venv\Scripts\pip.exe install --only-binary=:all: keyboard

echo Installing mouse...
venv\Scripts\pip.exe install --only-binary=:all: mouse

echo Installing pyperclip...
venv\Scripts\pip.exe install --only-binary=:all: pyperclip

echo Installing pywin32 (for window process detection)...
venv\Scripts\pip.exe install pywin32==306

echo.
echo Dependencies reinstalled. Checking if modules can be imported...
echo.

echo Testing imports:
venv\Scripts\python.exe -c "import PIL; print('✓ Pillow installed successfully')" 2>nul || echo ✗ Pillow import failed
venv\Scripts\python.exe -c "import psutil; print('✓ psutil installed successfully')" 2>nul || echo ✗ psutil import failed
venv\Scripts\python.exe -c "import pygetwindow; print('✓ PyGetWindow installed successfully')" 2>nul || echo ✗ PyGetWindow import failed
venv\Scripts\python.exe -c "import pyautogui; print('✓ PyAutoGUI installed successfully')" 2>nul || echo ✗ PyAutoGUI import failed
venv\Scripts\python.exe -c "import keyboard; print('✓ keyboard installed successfully')" 2>nul || echo ✗ keyboard import failed
venv\Scripts\python.exe -c "import mouse; print('✓ mouse installed successfully')" 2>nul || echo ✗ mouse import failed
venv\Scripts\python.exe -c "import pyperclip; print('✓ pyperclip installed successfully')" 2>nul || echo ✗ pyperclip import failed
venv\Scripts\python.exe -c "import win32gui; print('✓ pywin32 installed successfully')" 2>nul || echo ✗ pywin32 import failed

echo.
echo Dependency repair complete.
echo If some modules still failed to import, try:
echo 1. Running this script as administrator
echo 2. Checking your internet connection
echo 3. Deleting the venv folder and running setup_once.bat again
echo.

call deactivate
pause 