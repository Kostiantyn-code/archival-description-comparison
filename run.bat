@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creating a local Python environment...
    where py >nul 2>nul
    if not errorlevel 1 (
        py -3 -m venv .venv
    ) else (
        python -m venv .venv
    )
    if errorlevel 1 goto :error
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    if errorlevel 1 goto :error
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 goto :error
)

".venv\Scripts\python.exe" main.py
if errorlevel 1 goto :error
echo.
echo Analysis completed.
pause
exit /b 0

:error
echo.
echo The script stopped with an error. Review the message above.
pause
exit /b 1
