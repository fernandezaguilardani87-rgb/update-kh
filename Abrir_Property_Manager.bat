@echo off
title Property Manager - KH Inmobiliaria

cd /d "%~dp0"

set PYTHON_EXE=

python --version >/dev/null 2>&1
if not errorlevel 1 (
    set PYTHON_EXE=python
    goto :found
)

python3 --version >/dev/null 2>&1
if not errorlevel 1 (
    set PYTHON_EXE=python3
    goto :found
)

for %%V in (313 312 311 310 39 38) do (
    if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
        set PYTHON_EXE="%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
        goto :found
    )
    if exist "C:\Python%%V\python.exe" (
        set PYTHON_EXE="C:\Python%%V\python.exe"
        goto :found
    )
    if exist "C:\Program Files\Python%%V\python.exe" (
        set PYTHON_EXE="C:\Program Files\Python%%V\python.exe"
        goto :found
    )
)

echo.
echo  Python no encontrado en tu ordenador.
echo.
echo  Sigue estos pasos:
echo  1. Ve a https://www.python.org/downloads/
echo  2. Descarga la ultima version de Python 3
echo  3. Ejecuta el instalador
echo  4. MUY IMPORTANTE: activa "Add Python to PATH"
echo  5. Haz clic en "Install Now"
echo  6. Reinicia el ordenador y vuelve a abrir este archivo
echo.
pause
exit /b 1

:found
echo Python encontrado: %PYTHON_EXE%
echo.
echo Instalando/verificando dependencias...
%PYTHON_EXE% -m pip install pdfplumber openpyxl pandas pymupdf pytesseract pillow --quiet
echo.
echo Abriendo la aplicacion...
%PYTHON_EXE% property_manager.py

if errorlevel 1 (
    echo.
    echo  La aplicacion se cerro con un error. Revisa el mensaje de arriba.
    pause
)
