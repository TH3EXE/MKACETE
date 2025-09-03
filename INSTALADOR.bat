@echo off
setlocal
color 0E

echo ======================================================
echo  Bem-vindo ao instalador automatico do MKACETE
echo ======================================================

REM Variaveis de configuracao
set PYTHON_VERSION=3.10.11
set PYTHON_INSTALLER=python-%PYTHON_VERSION%-amd64.exe
set PYTHON_DOWNLOAD_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/%PYTHON_INSTALLER%

echo.
echo [1/3] Verificando a instalacao do Python...

REM Verifica se o Python ja esta instalado
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo.
    echo %PYTHON_VERSION% ja esta instalado. Pulando a instalacao do Python.
) else (
    echo.
    echo Python nao detectado. Iniciando a instalacao...
    echo Baixando o instalador do Python...
    
    powershell -Command "Invoke-WebRequest -Uri '%PYTHON_DOWNLOAD_URL%' -OutFile '%PYTHON_INSTALLER%'"
    
    if not exist "%PYTHON_INSTALLER%" (
        echo.
        echo ERRO: Falha ao baixar o instalador do Python. Verifique sua conexao com a internet.
        goto end
    )
    
    echo.
    echo Executando instalacao silenciosa do Python. Por favor, aguarde...
    echo Esta etapa pode demorar alguns minutos.
    
    REM Executa a instalacao silenciosa
    "%PYTHON_INSTALLER%" /quiet InstallAllUsers=1 PrependPath=1
    
    if %errorlevel% neq 0 (
        echo.
        echo ERRO: A instalacao do Python falhou.
        goto end
    )
    
    echo.
    echo Instalacao do Python concluida com sucesso!
    echo.
)

REM Adiciona um pequeno atraso para que o PATH seja atualizado
timeout /t 5 >nul

echo [2/3] Instalando as dependencias do codigo...

REM Instala as bibliotecas usando o pip.
pip install pandas
pip install openpyxl
pip install unidecode
pip install colorama
pip install pyperclip

echo.
echo [3/3] Verificando a instalacao das dependencias...

REM Verifica se as dependencias foram instaladas corretamente
pip show pandas >nul 2>&1 || (echo ERRO: Falha ao instalar pandas. & goto end)
pip show openpyxl >nul 2>&1 || (echo ERRO: Falha ao instalar openpyxl. & goto end)
pip show unidecode >nul 2>&1 || (echo ERRO: Falha ao instalar unidecode. & goto end)
pip show colorama >nul 2>&1 || (echo ERRO: Falha ao instalar colorama. & goto end)
pip show pyperclip >nul 2>&1 || (echo ERRO: Falha ao instalar pyperclip. & goto end)

echo.
echo ======================================================
echo  Instalacao completa!
echo  Voce ja pode executar o seu script.
echo ======================================================

:end
echo.
pause
