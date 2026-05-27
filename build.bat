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
    echo   Distribua o .exe pelos utilizadores:
    echo     dist\GestaoCapsulas.exe  ->  Desktop (ou pasta local) de cada PC
    echo.
    echo   A base de dados e os ficheiros Excel ficam na rede:
    echo     \\sidel.com\emea\pt-smf\groups\STORAGE\Maquinas\Producao\caps
    echo.
    echo   A app verifica a ligacao a rede automaticamente ao abrir.
    echo.
) else (
    echo   ERRO no build. Verifique as mensagens acima.
)

pause
