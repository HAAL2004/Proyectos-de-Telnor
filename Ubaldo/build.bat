@echo off
REM ============================================================
REM  build.bat  —  Genera el ejecutable Reportes.exe
REM  Requiere: pip install pyinstaller
REM ============================================================

echo.
echo  Verificando PyInstaller...
py -m pip install pyinstaller --quiet

echo.
echo  Construyendo Reportes.exe ...
echo  (esto puede tardar un par de minutos)
echo.

py -m PyInstaller ^
    --noconfirm ^
    --onedir ^
    --windowed ^
    --name "Reportes" ^
    --add-data "scraper.py;." ^
    --add-data "ESCALADAS 2026_03_20.xlsx;." ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    --collect-all selenium ^
    --collect-all webdriver_manager ^
    Reportes.py

echo.
if exist dist\Reportes\Reportes.exe (
    echo  ============================================================
    echo   Listo! El ejecutable se encuentra en:
    echo   dist\Reportes\Reportes.exe
    echo.
    echo   Para distribuir la app, comparte la carpeta completa:
    echo   dist\Reportes\
    echo  ============================================================
) else (
    echo  [ERROR] No se encontro el ejecutable. Revisa los mensajes de arriba.
)
echo.
pause
