@echo off
chcp 65001 >nul
title Gerar Executavel (Arquivo Unico) - Importar Usuarios AD

echo.
echo ============================================================
echo   GERADOR DE EXECUTAVEL - ARQUIVO UNICO (--onefile)
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

:: -- 3. Compila o projeto em ARQUIVO UNICO ---------------------
echo.
echo [3/4] Compilando com PyInstaller (Arquivo Unico)...
echo       Aguarde, este processo pode levar de 30 a 60 segundos...
echo.

if exist "dist\ImportarUsuariosUmuarama.exe" (
    echo Removendo executavel anterior...
    del /f /q "dist\ImportarUsuariosUmuarama.exe" 2>nul
)

python -m PyInstaller --noconsole ^
                       --onefile ^
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

:: -- 4. Copia ou_map.json para a pasta dist --------------------
echo.
echo [4/5] Finalizando...
if exist "ou_map.json" (
    copy /y "ou_map.json" "dist\ou_map.json" >nul 2>&1
    echo       OK - Copia de ou_map.json salva em dist\ para consulta/edicao externa.
)

:: -- 5. Publica automaticamente no Portal de Servicos (se existir)
if exist "..\portal\downloads" (
    copy /y "dist\ImportarUsuariosUmuarama.exe" "..\portal\downloads\ImportarUsuariosUmuarama.exe" >nul 2>&1
    copy /y "ou_map.json" "..\portal\downloads\ou_map.json" >nul 2>&1
    echo [5/5] Atualizado automaticamente no Portal de Servicos (portal\downloads\).
)

echo.
echo ============================================================
echo   EXECUTAVEL UNICO GERADO COM SUCESSO!
echo ============================================================
echo.
echo Arquivo gerado:
echo   %PASTA_PROJETO%dist\ImportarUsuariosUmuarama.exe
echo.
echo O ou_map.json ja esta embutido dentro do .exe.
echo (Se voce colocar um 'ou_map.json' na mesma pasta do .exe, ele
echo  tera prioridade, permitindo customizacao sem recompilar).
echo ============================================================
echo.

explorer "%PASTA_PROJETO%dist"
pause
