@echo off
REM ============================================================
REM  build.bat  —  Genera el ejecutable Rep_Infra.exe
REM  Requiere: pip install pyinstaller
REM ============================================================

echo.
echo  Verificando PyInstaller...
py -m pip install pyinstaller --quiet

echo.
echo  Construyendo Rep_Infra.exe ...
echo  (esto puede tardar un par de minutos)
echo.

py -m PyInstaller ^
    --noconfirm ^
    --onedir ^
    --windowed ^
    --name "Rep_Infra" ^
    --add-data "infra_logic.py;." ^
    --hidden-import pandas ^
    --hidden-import numpy ^
    --hidden-import openpyxl ^
    Rep_Infra_App.py

echo.
if exist dist\Rep_Infra\Rep_Infra.exe (
    echo  ============================================================
    echo   Listo! El ejecutable se encuentra en:
    echo   dist\Rep_Infra\Rep_Infra.exe
    echo.
    echo   Para distribuir la app, comparte la carpeta completa:
    echo   dist\Rep_Infra\
    echo  ============================================================
) else (
    echo  [ERROR] No se encontro el ejecutable. Revisa los mensajes de arriba.
)
echo.
pause
