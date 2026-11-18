@echo off
chcp 65001 >nul
title BO SWI Generador
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo.
  echo  No encontre Python. Instalalo desde https://www.python.org/downloads/
  echo  y marca la casilla "Add Python to PATH".
  echo.
  pause
  exit /b 1
)

python -c "import openpyxl, xlrd" >nul 2>nul
if errorlevel 1 (
  echo Instalando componentes por primera vez...
  python -m pip install --quiet -r requirements.txt
)

python app.py
pause
