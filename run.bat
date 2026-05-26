@echo off
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo Python nao encontrado. Execute install.bat
    pause
    exit /b 1
)

echo Instalando dependencias...
pip install -q -r requirements.txt 2>nul

echo.
echo Abrindo app em http://localhost:8501
echo.
timeout /t 2

start http://localhost:8501 2>nul

python -m streamlit run app.py --server.port 8501

pause
