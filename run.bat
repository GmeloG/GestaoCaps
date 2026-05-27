@echo off
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo Python nao encontrado. Execute install.bat primeiro.
    pause
    exit /b 1
)

echo Instalando dependencias...
pip install -q -r requirements.txt

echo.
echo A abrir a aplicacao...
python app.py

pause
