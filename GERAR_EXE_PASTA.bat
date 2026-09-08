@echo off
chcp 65001 >nul
title Gerar Executavel (Modo Pasta) - Importar Usuarios AD

echo.
echo ============================================================
echo   GERADOR DE EXECUTAVEL - MODO PASTA (--onedir)
echo   Importar Usuarios AD - Grupo Umuarama
echo ============================================================
echo.

set PASTA_PROJETO=%~dp0
cd /d "%PASTA_PROJETO%"

:: -- 1. Verifica Python ----------------------------------------
echo [1/4] Verificando instalacao do Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao foi encontrado no sistema.
    echo Por favor, instale o Python e marque 'Add to PATH'.
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo       OK - %%v

:: -- 2. Instala dependencias necessarias ----------------------
echo.
echo [2/4] Verificando e instalando dependencias (PyInstaller, pandas, openpyxl, chardet)...
python -m pip install --upgrade pip --quiet --disable-pip-version-check 2>nul
python -m pip install pyinstaller pandas openpyxl chardet --quiet --disable-pip-version-check --no-input
if %errorlevel% neq 0 (
    echo [ERRO] Falha ao instalar dependencias com pip.
    pause
    exit /b 1
)
echo       OK - Dependencias prontas.

:: -- 3. Compila o projeto em modo PASTA ------------------------
echo.
echo [3/4] Compilando com PyInstaller (Modo Pasta)...
echo       Aguarde, este processo pode levar de 30 a 60 segundos...
echo.

if exist "dist\ImportarUsuariosUmuarama" (
    echo Limpando build anterior...
    rmdir /s /q "dist\ImportarUsuariosUmuarama" 2>nul
)

python -m PyInstaller --noconsole ^
                       --onedir ^
                       --add-data "ou_map.json;." ^
                       --name "ImportarUsuariosUmuarama" ^
                       --clean ^
                       --noconfirm ^
                       ImportarUsuariosUmuarama.py

if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Ocorreu uma falha durante a geracao do executavel.
    pause
    exit /b 1
)

:: -- 4. Copia ou_map.json externo para a pasta gerada ----------
echo.
echo [4/4] Finalizando e copiando arquivos de configuracao...
if exist "ou_map.json" (
    copy /y "ou_map.json" "dist\ImportarUsuariosUmuarama\ou_map.json" >nul 2>&1
    echo       OK - ou_map.json copiado para a pasta do executavel.
)

echo.
echo ============================================================
echo   EXECUTAVEL GERADO COM SUCESSO!
echo ============================================================
echo.
echo Pasta de saida:
echo   %PASTA_PROJETO%dist\ImportarUsuariosUmuarama\
echo.
echo Executavel principal:
echo   %PASTA_PROJETO%dist\ImportarUsuariosUmuarama\ImportarUsuariosUmuarama.exe
echo.
echo Voce pode mover toda a pasta 'ImportarUsuariosUmuarama' para onde quiser.
echo O arquivo ou_map.json nesta pasta pode ser editado diretamente.
echo ============================================================
echo.

explorer "%PASTA_PROJETO%dist\ImportarUsuariosUmuarama"
pause
