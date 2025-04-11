@echo off
echo Clockify Report Processor
echo ========================
echo Starting the application...
echo Press Ctrl+C to exit the application

python -W ignore::DeprecationWarning src/main.py

echo Application closed 