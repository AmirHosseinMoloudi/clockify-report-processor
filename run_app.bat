@echo off
echo Clockify Report Processor
echo ========================
echo Upgrading pip...
python -m pip install --upgrade pip
echo Installing dependencies...
pip install -r requirements.txt
echo Starting the application...
echo Press Ctrl+C to exit the application

:: Run Python with SIGINT handler enabled
python -W ignore::DeprecationWarning src/main.py

echo Application closed
pause 