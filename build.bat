@echo off
cd /d "%~dp0"
echo.
echo ============================================================
echo   BUILD - GestaoCapsulas.exe
echo ============================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Python nao encontrado.
    pause
    exit /b 1
)

echo A instalar dependencias...
pip install -q -r requirements.txt
pip install -q pyinstaller

:: Usar pasta local fora do OneDrive para evitar conflitos de permissao
set BUILD_WORK=%TEMP%\pyi_GestaoCapsulas_work
set BUILD_DIST=%TEMP%\pyi_GestaoCapsulas_dist

echo.
echo A criar o .exe (pode demorar 3-5 minutos)...
echo.

pyinstaller app.py ^
    --name GestaoCapsulas ^
    --onefile ^
    --windowed ^
    --collect-all ttkbootstrap ^
    --collect-all openpyxl ^
    --hidden-import pandas ^
    --hidden-import sqlite3 ^
    --workpath "%BUILD_WORK%" ^
    --distpath "%BUILD_DIST%" ^
    --clean ^
    --noconfirm

echo.
if exist "%BUILD_DIST%\GestaoCapsulas.exe" (
    :: Copiar o .exe para a pasta dist\ do projeto
    if not exist "dist" mkdir dist
    copy /Y "%BUILD_DIST%\GestaoCapsulas.exe" "dist\GestaoCapsulas.exe" >nul

    echo ============================================================
    echo   BUILD CONCLUIDO COM SUCESSO!
    echo ============================================================
    echo.
    echo   Ficheiro gerado: dist\GestaoCapsulas.exe
    echo.
    echo   Copie para a pasta partilhada:
    echo     1. dist\GestaoCapsulas.exe
    echo     2. capsulas.db  (se ja tiver dados)
    echo.
    echo   Os colegas fazem duplo-clique em GestaoCapsulas.exe
    echo   Abre diretamente como uma aplicacao Windows normal.
    echo.
) else (
    echo   ERRO no build. Verifique as mensagens acima.
)

pause
