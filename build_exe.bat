@echo off
chcp 65001 >nul
title Crear SICAG.exe
echo ============================================================
echo  SICAG - Sistema Integrado de Control de Avances Geologicos: crear ejecutable (.exe)
echo ============================================================
where python >nul 2>nul
if errorlevel 1 (
  echo.
  echo No se encontro Python. Instala Python 3.10 o superior desde https://www.python.org/downloads/
  echo y marca la casilla "Add python.exe to PATH" durante la instalacion.
  pause
  exit /b 1
)
echo.
echo [1/3] Creando entorno e instalando librerias (puede tardar unos minutos)...
python -m venv .venv_build
call .venv_build\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller
if errorlevel 1 ( echo Fallo la instalacion de librerias. & pause & exit /b 1 )
echo.
echo [2/3] Generando el ejecutable...
pyinstaller --noconfirm --clean --onefile --windowed --name SICAG --icon logo.ico --collect-all docx --collect-all pptx --hidden-import openpyxl app_sicag.py
if errorlevel 1 ( echo Fallo la creacion del ejecutable. & pause & exit /b 1 )
echo.
echo [3/3] Listo.
echo El ejecutable esta en:  %CD%\dist\SICAG.exe
echo Puedes copiarlo a cualquier PC con Windows; no necesita Python instalado.
pause
