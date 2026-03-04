@echo off
REM Windows startup script for Input Studio Web Server

echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo Starting server...
python server\main.py

pause
