@echo off
title Annual Statements Automation
echo ===================================================
echo Starting Annual Statement Downloader...
echo ===================================================
echo.

:: Navigate to the folder where this batch file and script live
cd /d "%~dp0"

:: Activate virtual environment if you use one (uncomment the line below if applicable)
:: call venv\Scripts\activate

:: Run the Python script
python annual_statements.py

echo.
echo ===================================================
echo Process complete!
echo ===================================================
pause