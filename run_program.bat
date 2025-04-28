@echo off
REM Change directory to the script's location
pushd "%~dp0"

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
echo Activating virtual environment...
call venv\Scripts\activate

REM Optional: Check if activation was successful by checking the Python executable path
REM python -c "import sys; print(sys.executable)" | findstr /I /C:"\\venv\\Scripts\\python.exe" > nul
REM if %errorlevel% neq 0 (
REM   echo Error: Failed to activate the virtual environment correctly.
REM   pause
REM   exit /b 1
REM )

REM Run the application
echo Running program with virtual environment Python...
venv\Scripts\python.exe main.py

REM Deactivate environment when done (optional, script will exit anyway)
REM call deactivate

echo.
echo Program has exited.

REM Return to the original directory (optional but good practice)
popd

pause
exit /b 0 