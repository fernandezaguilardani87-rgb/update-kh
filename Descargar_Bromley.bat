@echo off
title Descargador Bromley Estates

cd /d "%~dp0"

echo.
echo  Descargador de precios - Bromley Estates
echo  ==========================================
echo.

:: Buscar Python
set PYTHON_EXE=
python --version >/dev/null 2>&1
if not errorlevel 1 (set PYTHON_EXE=python && goto :run)
for %%V in (313 312 311 310 39 38) do (
    if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
        set PYTHON_EXE="%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
        goto :run
    )
)
echo Python no encontrado. Abre Abrir_Property_Manager.bat primero.
pause
exit /b 1

:run
echo Instalando dependencias...
%PYTHON_EXE% -m pip install playwright --quiet
%PYTHON_EXE% -m playwright install chromium --quiet
echo.
echo Iniciando descarga...
echo.
%PYTHON_EXE% descargar_bromley.py

if errorlevel 1 (
    echo.
    echo  Error. Revisa el mensaje de arriba.
    pause
)
