@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ================================================
echo   AutoMail - Verificacao de Dependencias
echo ================================================
echo.

:: ─── 1. Python ──────────────────────────────────────────────────────────────
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Python nao encontrado. Tentando instalar via winget...
    winget install --id Python.Python.3.12 --source winget --silent
    if %errorlevel% neq 0 (
        echo.
        echo [ERRO] Nao foi possivel instalar o Python automaticamente.
        echo        Baixe e instale manualmente em: https://www.python.org/downloads/
        echo        Marque a opcao "Add Python to PATH" durante a instalacao.
        echo.
        pause
        exit /b 1
    )
    echo [OK] Python instalado.
    echo [!] Feche e reabra este terminal, depois rode o script novamente.
    echo.
    pause
    exit /b 0
)

for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PY_VER=%%i
echo [OK] !PY_VER! encontrado.

:: ─── 2. pip ─────────────────────────────────────────────────────────────────
python -m pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] pip nao encontrado. Instalando...
    python -m ensurepip --upgrade
    if %errorlevel% neq 0 (
        echo [ERRO] Nao foi possivel instalar o pip.
        pause
        exit /b 1
    )
)
echo [OK] pip disponivel.

:: ─── 3. Atualizar pip ───────────────────────────────────────────────────────
echo [*] Atualizando pip...
python -m pip install --upgrade pip --quiet --no-warn-script-location
echo [OK] pip atualizado.

:: ─── 4. pywin32 ─────────────────────────────────────────────────────────────
echo.
echo [*] Verificando pywin32...
python -c "import win32com.client" >nul 2>&1
if %errorlevel% neq 0 (
    echo [*] Instalando pywin32...
    python -m pip install pywin32 --quiet --no-warn-script-location
    if %errorlevel% neq 0 (
        echo [ERRO] Falha ao instalar pywin32.
        pause
        exit /b 1
    )

    :: Post-install obrigatorio para registrar DLLs do pywin32
    echo [*] Configurando pywin32 ^(registrando DLLs^)...
    for /f "tokens=*" %%p in ('python -c "import sys; print(sys.prefix)"') do set PY_PREFIX=%%p
    python "!PY_PREFIX!\Scripts\pywin32_postinstall.py" -install >nul 2>&1
    if %errorlevel% neq 0 (
        :: Fallback: tenta via modulo direto
        python -c "import pywin32_postinstall; pywin32_postinstall.install()" >nul 2>&1
    )
    echo [OK] pywin32 instalado e configurado.
) else (
    echo [OK] pywin32 ja instalado.
)

:: ─── 5. py7zr ───────────────────────────────────────────────────────────────
echo.
echo [*] Verificando py7zr...
python -c "import py7zr" >nul 2>&1
if %errorlevel% neq 0 (
    echo [*] Instalando py7zr...
    python -m pip install py7zr --quiet --no-warn-script-location
    if %errorlevel% neq 0 (
        echo [ERRO] Falha ao instalar py7zr.
        pause
        exit /b 1
    )
    echo [OK] py7zr instalado.
) else (
    echo [OK] py7zr ja instalado.
)

:: ─── 6. Validacao final ──────────────────────────────────────────────────────
echo.
echo [*] Validando instalacao...
python -c "import imaplib, email, json, zipfile, re, csv, os, sys; import win32com.client; import py7zr; print('OK')" >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Validacao falhou. Algum modulo nao foi carregado corretamente.
    echo        Tente rodar este script como Administrador.
    echo.
    pause
    exit /b 1
)

echo.
echo ================================================
echo   Tudo pronto! Todas as dependencias estao
echo   instaladas. Rode o AutoMail.bat para iniciar.
echo ================================================
echo.
pause
