@echo off
chcp 65001 >nul
title Importacao de Colaboradores - UAC

echo.
echo ============================================================
echo   IMPORTACAO DE COLABORADORES
echo ============================================================
echo.

set SERVIDOR=\\10.56.43.28\Users\pedro.lima\Desktop\IMPORT
set USUARIO=pedro.lima@umuarama.local
if not defined SENHA set /p SENHA=Digite a senha do usuario %USUARIO%: 
set PASTA_LOCAL=C:\Users\Pedro Lima\Documents\IMPORT_COLABORADORES
set SCRIPT_PYTHON=%PASTA_LOCAL%\gerar_importar_usuarios_uap.py

echo [1/4] Conectando ao servidor...
net use "%SERVIDOR%" "%SENHA%" /user:"%USUARIO%" >nul 2>&1
if errorlevel 1 (
    echo ERRO: Nao foi possivel conectar ao servidor.
    echo Verifique a conexao de rede e as credenciais.
    echo.
    pause
    exit /b 1
)
echo       OK - Servidor acessivel.

echo.
echo [2/4] Buscando arquivo mais recente...
set ARQUIVO_MAIS_RECENTE=

for /f "delims=" %%F in ('dir /b /o-d "%SERVIDOR%\FPRE111-*.CSV" 2^>nul') do (
    if not defined ARQUIVO_MAIS_RECENTE set ARQUIVO_MAIS_RECENTE=%%F
)

if not defined ARQUIVO_MAIS_RECENTE (
    echo ERRO: Nenhum arquivo FPRE111-*.CSV encontrado no servidor.
    net use "%SERVIDOR%" /delete >nul 2>&1
    pause
    exit /b 1
)

set CAMINHO_ARQUIVO=%SERVIDOR%\%ARQUIVO_MAIS_RECENTE%
echo       OK - Arquivo encontrado: %ARQUIVO_MAIS_RECENTE%

echo.
echo [3/4] Processando %ARQUIVO_MAIS_RECENTE%...
echo ============================================================
python "%SCRIPT_PYTHON%" "%CAMINHO_ARQUIVO%"
set RESULTADO=%errorlevel%
echo ============================================================
echo.

if %RESULTADO% neq 0 (
    echo ERRO: O script Python retornou um erro.
    net use "%SERVIDOR%" /delete >nul 2>&1
    pause
    exit /b 1
)

echo [4/4] Abrindo arquivo de saida...

set NOME_SEM_EXT=%ARQUIVO_MAIS_RECENTE:~0,-4%
set DATA_ARQUIVO=%NOME_SEM_EXT:~7%
set DIA=%DATA_ARQUIVO:~0,2%
set MES=%DATA_ARQUIVO:~2,2%

set PASTA_SAIDA=%PASTA_LOCAL%\%DIA%-%MES%
set CSV_SAIDA=%PASTA_SAIDA%\importar-usuarios-%DATA_ARQUIVO%.csv

if exist "%CSV_SAIDA%" (
    echo       OK - Abrindo: %CSV_SAIDA%
    start "" "%CSV_SAIDA%"
) else (
    echo       AVISO: Arquivo de saida nao encontrado em %CSV_SAIDA%
)

net use "%SERVIDOR%" /delete >nul 2>&1

echo.
echo ============================================================
echo   CONCLUIDO COM SUCESSO
echo ============================================================
echo.
pause