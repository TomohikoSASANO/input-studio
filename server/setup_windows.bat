@echo off
REM Windows setup script for Input Studio Web Server

echo Creating virtual environment...
python -m venv .venv

echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo Installing dependencies...
python -m pip install --upgrade pip
python -m pip install -r server\requirements.txt

echo Setup complete!
echo.
echo To start the server, run:
echo   .venv\Scripts\activate.bat
echo   python server\main.py

pause
