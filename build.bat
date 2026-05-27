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

echo A preparar o build (pode demorar 3-5 minutos)...
echo.

pyinstaller launcher.py ^
    --name GestaoCapsulas ^
    --onefile ^
    --collect-all streamlit ^
    --collect-all openpyxl ^
    --add-data "app.py;." ^
    --clean ^
    --noconfirm

echo.
if exist "dist\GestaoCapsulas.exe" (
    echo ============================================================
    echo   BUILD CONCLUIDO COM SUCESSO!
    echo ============================================================
    echo.
    echo   Ficheiro gerado: dist\GestaoCapsulas.exe
    echo.
    echo   Para a pasta partilhada copie:
    echo     1. dist\GestaoCapsulas.exe
    echo     2. capsulas.db  (se ja tiver dados)
    echo.
    echo   Os colegas fazem duplo-clique em GestaoCapsulas.exe
    echo   Uma janela preta aparece - NAO fechar - e o browser abre sozinho.
    echo.
) else (
    echo   ERRO no build. Verifique as mensagens acima.
)

pause
