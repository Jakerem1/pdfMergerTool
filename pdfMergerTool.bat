@echo off
REM Get the folder where this BAT file lives
set BASEDIR=%~dp0

REM Activate the virtual environment
call "%BASEDIR%venv\Scripts\activate.bat"

REM Run your Python script
python "%BASEDIR%pdfMergerTool.py"

REM Keep window open if something crashes
pause