@echo off
chcp 65001 > nul
echo Starting user behavior collector...

REM Check if virtual environment exists
if not exist venv (
  echo Error: Virtual environment not found.
  echo Please run setup_once.bat first to create the environment and install dependencies.
  pause
  exit /b 1
)

REM Activate virtual environment
call venv\Scripts\activate

REM Check and install missing dependencies before running
echo Checking dependencies...
venv\Scripts\python.exe -c "import PIL" 2>nul
if %errorlevel% neq 0 (
  echo Installing missing dependency: Pillow...
  venv\Scripts\pip.exe install --only-binary=:all: Pillow
)

venv\Scripts\python.exe -c "import psutil" 2>nul
if %errorlevel% neq 0 (
  echo Installing missing dependency: psutil...
  venv\Scripts\pip.exe install --only-binary=:all: psutil
)

venv\Scripts\python.exe -c "import pygetwindow" 2>nul
if %errorlevel% neq 0 (
  echo Installing missing dependency: PyGetWindow...
  venv\Scripts\pip.exe install --only-binary=:all: PyGetWindow
)

venv\Scripts\python.exe -c "import pyautogui" 2>nul
if %errorlevel% neq 0 (
  echo Installing missing dependency: PyAutoGUI...
  venv\Scripts\pip.exe install --only-binary=:all: PyAutoGUI
)

venv\Scripts\python.exe -c "import keyboard" 2>nul
if %errorlevel% neq 0 (
  echo Installing missing dependency: keyboard...
  venv\Scripts\pip.exe install --only-binary=:all: keyboard
)

venv\Scripts\python.exe -c "import mouse" 2>nul
if %errorlevel% neq 0 (
  echo Installing missing dependency: mouse...
  venv\Scripts\pip.exe install --only-binary=:all: mouse
)

venv\Scripts\python.exe -c "import pyperclip" 2>nul
if %errorlevel% neq 0 (
  echo Installing missing dependency: pyperclip...
  venv\Scripts\pip.exe install --only-binary=:all: pyperclip
)

venv\Scripts\python.exe -c "import win32gui" 2>nul
if %errorlevel% neq 0 (
  echo Installing missing dependency: pywin32...
  venv\Scripts\pip.exe install pywin32==306
)

echo All dependencies checked and installed.

REM Run the application
echo Running program with virtual environment Python...
venv\Scripts\python.exe main.py

REM Deactivate environment when done
call deactivate

echo Program has exited
pause 