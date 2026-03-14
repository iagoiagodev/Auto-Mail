@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ================================================
echo   AutoMail - Gerador de Executavel (.exe)
echo ================================================
echo.

:: ─── Python ─────────────────────────────────────────────────────────────────
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado. Execute install-dependencies.bat primeiro.
    pause & exit /b 1
)
for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PY_VER=%%i
echo [OK] !PY_VER!

:: ─── Dependencias de build ───────────────────────────────────────────────────
echo [*] Verificando PyInstaller...
python -m pip install pyinstaller pywin32 py7zr --quiet --no-warn-script-location
echo [OK] Dependencias prontas.

:: ─── Limpeza ─────────────────────────────────────────────────────────────────
echo [*] Limpando build anterior...
if exist "build"           rmdir /s /q "build"
if exist "dist\AutoMail"   rmdir /s /q "dist\AutoMail"
if exist "AutoMail.spec"   del /q "AutoMail.spec"

:: ─── Compilar ────────────────────────────────────────────────────────────────
echo [*] Compilando... (pode demorar alguns minutos)
echo.

python -m PyInstaller ^
    --onedir ^
    --console ^
    --name AutoMail ^
    --hidden-import=win32com ^
    --hidden-import=win32com.client ^
    --hidden-import=win32com.client.dynamic ^
    --hidden-import=win32com.client.gencache ^
    --hidden-import=win32com.server ^
    --hidden-import=win32com.server.util ^
    --hidden-import=pywintypes ^
    --hidden-import=win32api ^
    --hidden-import=win32con ^
    --hidden-import=win32timezone ^
    --hidden-import=tkinter ^
    --hidden-import=tkinter.ttk ^
    --hidden-import=tkinter.filedialog ^
    --hidden-import=tkinter.messagebox ^
    --collect-all=py7zr ^
    --collect-all=multivolumefile ^
    --exclude-module=matplotlib ^
    --exclude-module=numpy ^
    --exclude-module=pandas ^
    --exclude-module=PIL ^
    --exclude-module=scipy ^
    --exclude-module=IPython ^
    --exclude-module=notebook ^
    main.py

if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Falha na compilacao. Veja o log acima.
    pause & exit /b 1
)

:: ─── Copiar arquivos necessarios para dist ───────────────────────────────────
echo.
echo [*] Preparando pasta de distribuicao...
if exist "config.example.json" copy /y "config.example.json" "dist\AutoMail\config.example.json" >nul

:: Cria AutoMail.bat dentro da pasta para facil execucao
(
    echo @echo off
    echo cd /d "%%~dp0"
    echo AutoMail.exe
    echo pause
) > "dist\AutoMail\Executar AutoMail.bat"

:: ─── Resultado ───────────────────────────────────────────────────────────────
echo.
echo ================================================
echo   Executavel gerado em: dist\AutoMail\
echo.
echo   Para distribuir para outro computador:
echo   1. Copie a pasta dist\AutoMail\ inteira
echo   2. Crie config.json (baseado em config.example.json)
echo      OU deixe em branco — a GUI preenche tudo
echo   3. Execute "Executar AutoMail.bat"
echo.
echo   Requisito minimo no destino: Outlook instalado
echo   (Python NAO e necessario)
echo ================================================
echo.
pause
