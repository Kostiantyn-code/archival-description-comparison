@echo off
setlocal
cd /d "%~dp0"

if not exist "main.py" goto :missing_files
if not exist "requirements.txt" goto :missing_files

rem Keep dependencies across extractions of the project ZIP.
if defined LOCALAPPDATA (
    set "VENV_DIR=%LOCALAPPDATA%\archival-description-comparison\venv"
) else (
    set "VENV_DIR=%~dp0.venv"
)
rem Continue using an environment already created next to this project.
if exist ".venv\Scripts\python.exe" set "VENV_DIR=%~dp0.venv"
set "PYTHON=%VENV_DIR%\Scripts\python.exe"
echo Python environment: "%VENV_DIR%"

if exist "%PYTHON%" goto :check_requirements
echo Creating a Python environment (first launch only)...
where py >nul 2>nul
if errorlevel 1 goto :create_with_python
py -3 -m venv "%VENV_DIR%"
if errorlevel 1 goto :error
goto :check_requirements

:create_with_python
python -m venv "%VENV_DIR%"
if errorlevel 1 goto :error

:check_requirements
if not exist "%PYTHON%" goto :error
if not exist "%VENV_DIR%\installed-requirements.txt" goto :install
fc /b "requirements.txt" "%VENV_DIR%\installed-requirements.txt" >nul 2>nul
if errorlevel 1 goto :install
echo Dependencies already installed.
goto :run

:install
echo Installing dependencies (first launch or changed requirements)...
"%PYTHON%" -m pip install -r "requirements.txt"
if errorlevel 1 goto :error
copy /y "requirements.txt" "%VENV_DIR%\installed-requirements.txt" >nul
if errorlevel 1 goto :error

:run
"%PYTHON%" main.py %*
if errorlevel 1 goto :error
echo.
echo Analysis completed.
pause
exit /b 0

:missing_files
echo Extract the entire ZIP archive before running run.bat.
goto :error

:error
echo.
echo The script stopped with an error. Review the message above.
pause
exit /b 1
