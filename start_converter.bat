@echo off
REM This batch file starts the File to PDF Converter application.

REM Get the directory of the batch file
SET SCRIPT_DIR=%~dp0

REM Change to the script's directory to ensure relative paths work correctly
cd /D "%SCRIPT_DIR%"

echo Starting File to PDF Converter...
python converter.py

echo Application closed.
pause 