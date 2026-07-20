@echo off
chcp 65001 >nul
title Importacao de Colaboradores - Grupo Umuarama

:: ============================================================
::  CONFIGURACOES
:: ============================================================
set SERVIDOR=\\10.56.43.28\Users\pedro.lima\Desktop\IMPORT
set PASTA_LOCAL=C:\Users\Pedro Lima\Documents\IMPORT_COLABORADORES
set SCRIPT_GERAR=%PASTA_LOCAL%\gerar_importar_usuarios_uap.py
set SCRIPT_IMPORTAR=%PASTA_LOCAL%\ImportarUsuariosUmuarama.py
:: ============================================================

echo.
echo ============================================================
echo   Importacao de Colaboradores - Grupo Umuarama
echo ============================================================
echo.

:: -- Verifica Python ------------------------------------------
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado.
    echo Instale em: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo Python OK.

:: -- Instala dependencias -------------------------------------
echo Verificando dependencias...
python -m pip install chardet pandas openpyxl --quiet --disable-pip-version-check --no-input 2>nul
echo       OK.

:: -- Verifica scripts -----------------------------------------
if not exist "%SCRIPT_GERAR%" (
    echo [ERRO] Script nao encontrado: %SCRIPT_GERAR%
    pause
    exit /b 1
)
if not exist "%SCRIPT_IMPORTAR%" (
    echo [ERRO] Script nao encontrado: %SCRIPT_IMPORTAR%
    pause
    exit /b 1
)
echo Scripts OK.

:: -- Etapa 1: Conecta ao servidor (usa cmdkey implicito) ------
echo.
echo [1/3] Conectando ao servidor %SERVIDOR%...
net use "%SERVIDOR%" /persistent:no >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Nao foi possivel conectar ao servidor.
    echo Execute CADASTRAR_CREDENCIAIS.bat para salvar as credenciais.
    pause
    exit /b 1
)
echo       OK - Servidor acessivel.

:: -- Etapa 2: Busca o FPRE111 mais recente --------------------
echo.
echo [2/3] Buscando arquivo FPRE111 mais recente...
set ARQUIVO_FPRE=

for /f "delims=" %%F in ('dir /b /o-d "%SERVIDOR%\FPRE111-*.CSV" 2^>nul') do (
    if not defined ARQUIVO_FPRE set ARQUIVO_FPRE=%%F
)

if not defined ARQUIVO_FPRE (
    echo [ERRO] Nenhum arquivo FPRE111-*.CSV encontrado em %SERVIDOR%
    net use "%SERVIDOR%" /delete >nul 2>&1
    pause
    exit /b 1
)
echo       OK - Arquivo: %ARQUIVO_FPRE%

:: -- Determina caminho do CSV de saida ------------------------
set NOME_SEM_EXT=%ARQUIVO_FPRE:~0,-4%
set DATA_ARQUIVO=%NOME_SEM_EXT:~8%
set DIA=%DATA_ARQUIVO:~0,2%
set MES=%DATA_ARQUIVO:~2,2%
set PASTA_SAIDA=%PASTA_LOCAL%\%DIA%-%MES%
set CSV_ADMISSOES=%PASTA_SAIDA%\importar-usuarios-%DATA_ARQUIVO%.csv

:: -- Verifica se CSV ja existe e esta bloqueado ---------------
if exist "%CSV_ADMISSOES%" (
    2>nul (>>"%CSV_ADMISSOES%" echo off) || (
        echo.
        echo [ERRO] O arquivo de saida esta aberto em outro programa:
        echo        %CSV_ADMISSOES%
        echo.
        echo Feche o arquivo no Excel ou em outro programa e tente novamente.
        net use "%SERVIDOR%" /delete >nul 2>&1
        pause
        exit /b 1
    )
    echo       Aviso: CSV anterior sera sobrescrito.
)

:: -- Etapa 3: Gera o CSV de admissoes -------------------------
echo.
echo [3/3] Gerando CSV de admissoes a partir de %ARQUIVO_FPRE%...
echo ------------------------------------------------------------
set CAMINHO_FPRE=%SERVIDOR%\%ARQUIVO_FPRE%
python "%SCRIPT_GERAR%" "%CAMINHO_FPRE%"
set RESULTADO_GERAR=%errorlevel%
echo ------------------------------------------------------------

if %RESULTADO_GERAR% neq 0 (
    echo [ERRO] Falha ao gerar CSV. Codigo: %RESULTADO_GERAR%
    echo Verifique se o arquivo de saida nao esta aberto no Excel.
    net use "%SERVIDOR%" /delete >nul 2>&1
    pause
    exit /b 1
)

if not exist "%CSV_ADMISSOES%" (
    echo [ERRO] CSV nao encontrado: %CSV_ADMISSOES%
    net use "%SERVIDOR%" /delete >nul 2>&1
    pause
    exit /b 1
)
echo       OK - CSV gerado: %CSV_ADMISSOES%

:: -- Encerra conexao ------------------------------------------
net use "%SERVIDOR%" /delete >nul 2>&1

:: -- Abre GUI para pre-visualizacao e importacao manual -------
echo.
echo ============================================================
echo   Abrindo interface para pre-visualizacao e importacao...
echo   Informe suas credenciais AD na tela e clique Importar.
echo ============================================================
echo.
python "%SCRIPT_IMPORTAR%" --csv-preload "%CSV_ADMISSOES%"