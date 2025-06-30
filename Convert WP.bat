@echo off
REM Windows launcher for WP Converter
REM Changes to the script's directory and runs the Python application

cd /d "%~dp0"

REM Try different Python executables in order of preference
python wpd_to_docx.py 2>nul && goto :success
python3 wpd_to_docx.py 2>nul && goto :success
py wpd_to_docx.py 2>nul && goto :success

REM If we get here, Python wasn't found
echo Error: Python not found. Please install Python from python.org
echo Make sure Python is added to your PATH during installation.
pause
exit /b 1

:success
pause