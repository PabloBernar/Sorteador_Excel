@echo off
title Compilador - Sorteador de Base de Dados
color 0A
echo.
echo  ============================================
echo   Sorteador de Base de Dados - Build Script
echo  ============================================
echo.

REM Verifica Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado. Instale em https://python.org
    pause
    exit /b 1
)

echo [1/3] Instalando dependencias...
pip install pyinstaller pandas openpyxl xlrd pillow --quiet
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    pause
    exit /b 1
)

echo [2/3] Compilando para .exe com icone...
pyinstaller --onefile --windowed ^
  --name "Sorteador_BaseDados" ^
  --icon "icone.ico" ^
  --add-data "icone.ico;." ^
  --add-data "icone.png;." ^
  --hidden-import=openpyxl ^
  --hidden-import=openpyxl.cell._writer ^
  --hidden-import=openpyxl.styles ^
  --hidden-import=openpyxl.styles.fills ^
  --hidden-import=openpyxl.styles.fonts ^
  --hidden-import=openpyxl.styles.borders ^
  --hidden-import=openpyxl.styles.alignment ^
  --hidden-import=openpyxl.styles.numbers ^
  --hidden-import=openpyxl.utils ^
  --hidden-import=openpyxl.utils.dataframe ^
  --hidden-import=openpyxl.reader.excel ^
  --hidden-import=openpyxl.writer.excel ^
  --hidden-import=openpyxl.workbook ^
  --hidden-import=openpyxl.worksheet ^
  --hidden-import=openpyxl.worksheet._writer ^
  --hidden-import=openpyxl.worksheet.worksheet ^
  --hidden-import=xlrd ^
  --hidden-import=pandas._libs.tslibs.base ^
  --hidden-import=pandas._libs.tslibs.np_datetime ^
  --hidden-import=pandas._libs.tslibs.nattype ^
  --hidden-import=pandas._libs.tslibs.timedeltas ^
  --hidden-import=pandas._libs.tslibs.timestamps ^
  --hidden-import=pandas._libs.skiplist ^
  --collect-all openpyxl ^
  --collect-all xlrd ^
  --clean ^
  sorteador.py

if errorlevel 1 (
    echo [ERRO] Falha na compilacao.
    pause
    exit /b 1
)

echo [3/3] Limpando arquivos temporarios...
rmdir /s /q build 2>nul
del /q *.spec 2>nul

echo.
echo  ============================================
echo   PRONTO!
echo   Executavel: dist\Sorteador_BaseDados.exe
echo  ============================================
echo.
pause
