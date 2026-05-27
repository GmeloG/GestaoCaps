@echo off
echo.
echo ============================================================
echo   INSTALACAO - Ambiente de Desenvolvimento
echo ============================================================
echo.
echo Este ficheiro e apenas para o administrador que vai gerar o .exe.
echo Os utilizadores finais usam GestaoCapsulas.exe diretamente.
echo.

python --version >nul 2>&1
if not errorlevel 1 (
    echo Python ja esta instalado!
    echo.
    echo A instalar dependencias...
    pip install -r requirements.txt
    pip install pyinstaller
    echo.
    echo Pronto! Pode agora correr build.bat para gerar o .exe
    pause
    exit /b 0
)

echo Python nao encontrado.
echo.
echo Descarregando Python 3.11...
set "INSTALLER=%TEMP%\python-installer.exe"
curl -L -o "%INSTALLER%" "https://www.python.org/ftp/python/3.11.7/python-3.11.7-amd64.exe"

if not exist "%INSTALLER%" (
    echo Erro ao descarregar. Instale manualmente em:
    echo https://www.python.org/downloads/
    pause
    exit /b 1
)

echo.
echo Instalando Python...
"%INSTALLER%" /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1

echo.
echo Python instalado! Abra um novo terminal e execute install.bat novamente.
echo.
pause
