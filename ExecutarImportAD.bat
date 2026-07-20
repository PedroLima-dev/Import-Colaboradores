@echo off
chcp 65001 >nul
title Importar Usuarios AD - Grupo Umuarama

echo.
echo ============================================================
echo   Importar Usuarios AD - Grupo Umuarama
echo ============================================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado.
    echo.
    echo Instale em: https://www.python.org/downloads/
    echo Marque "Add Python to PATH" durante a instalacao.
    echo.
    pause
    exit /b 1
)

echo Verificando dependencias...
python -m pip install chardet --quiet --disable-pip-version-check >nul 2>&1

set "SCRIPT=%~dp0ImportarUsuariosUmuarama.py"

if not exist "%SCRIPT%" (
    echo [ERRO] Arquivo nao encontrado: %SCRIPT%
    echo.
    echo Coloque ImportarUsuariosUmuarama.py na mesma pasta deste .bat
    echo.
    pause
    exit /b 1
)

echo Iniciando...
echo.
python "%SCRIPT%"

if %errorlevel% neq 0 (
    echo.
    echo [AVISO] Aplicacao encerrou com erro %errorlevel%.
    pause
)