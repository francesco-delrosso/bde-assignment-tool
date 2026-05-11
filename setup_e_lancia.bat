@echo off
REM Setup e lancio BDE Assignment Tool (Windows)
REM Doppio click su questo file per lanciare

echo ================================
echo  BDE Assignment Tool - Setup
echo ================================

set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%.venv

if not exist "%VENV_DIR%" (
    echo Creo virtual environment...
    python -m venv "%VENV_DIR%"
    echo Installo dipendenze...
    "%VENV_DIR%\Scripts\pip" install --quiet pandas openpyxl
    echo Setup completato!
) else (
    echo Virtual environment trovato.
)

echo.
echo Lancio BDE Assignment Tool...
echo Si apre nel browser: http://127.0.0.1:8787
echo Per chiudere: Ctrl+C
echo.

"%VENV_DIR%\Scripts\python" "%SCRIPT_DIR%auto_bde_gui.py"
pause
