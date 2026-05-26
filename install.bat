@echo off

echo.
echo Instalador - Python para Gestao de Stock
echo.

python --version
if not errorlevel 1 (
    echo Python ja esta instalado!
    pause
    exit /b 0
)

echo Python nao encontrado.
echo.
echo Descarregando Python 3.11...
echo.

set "INSTALLER=%TEMP%\python-installer.exe"
curl -L -o "%INSTALLER%" "https://www.python.org/ftp/python/3.11.7/python-3.11.7-amd64.exe"

if not exist "%INSTALLER%" (
    echo Erro ao descarregar. Instale manualmente de:
    echo https://www.python.org/downloads/
    pause
    exit /b 1
)

echo.
echo Instalando Python...
"%INSTALLER%" /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1

echo.
echo Instalacao concluida!
echo Proximo passo: execute run.bat
echo.
pause
