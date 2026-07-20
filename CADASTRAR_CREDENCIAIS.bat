@echo off
chcp 65001 >nul
title Cadastrar Credenciais - Grupo Umuarama

echo.
echo ============================================================
echo   Cadastro de Credencial do Servidor - Grupo Umuarama
echo   Execute este arquivo apenas uma vez por maquina.
echo ============================================================
echo.

echo Credencial de acesso ao servidor de arquivos (\\10.56.43.28)
echo.
set /p SERV_USUARIO=   Usuario (ex: pedro.lima@umuarama.local): 
set /p SERV_SENHA=     Senha: 

cmdkey /add:10.56.43.28 /user:%SERV_USUARIO% /pass:%SERV_SENHA%
if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Falha ao salvar credencial.
    pause
    exit /b 1
)

set SERV_SENHA=
set SERV_USUARIO=

echo.
echo ============================================================
echo   Credencial salva com sucesso!
echo   Voce ja pode usar o EXECUTAR_IMPORT_COMPLETO.bat
echo ============================================================
echo.
pause