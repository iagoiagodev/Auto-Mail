# Install-AutoMail.ps1
#
# Comando para instalação rápida via PowerShell:
# iwr -useb https://raw.githubusercontent.com/iagoiagodev/Auto-Mail/main/Install-AutoMail.ps1 | iex

$ErrorActionPreference = "Stop"

# Força o console a usar codificação UTF-8 para exibir os acentos corretamente
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "   AutoMail - Instalador e Atualizador Automático  " -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Configurações
$repoUser = "iagoiagodev"
$repoName = "Auto-Mail"
$branch = "main"

$repoZipUrl = "https://github.com/$repoUser/$repoName/archive/refs/heads/$branch.zip"

Write-Host "`nPor favor, escolha a pasta onde deseja instalar o AutoMail na janela que abriu (pode estar minimizada)." -ForegroundColor Yellow

# Picker moderno estilo Explorer (IFileOpenDialog — mesmo usado pelo Chrome, Outlook, etc.)
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class ModernFolderPicker {
    private const uint FOS_PICKFOLDERS     = 0x00000020;
    private const uint FOS_FORCEFILESYSTEM = 0x00000040;
    private const uint SIGDN_FILESYSPATH   = 0x80058000;

    [ComImport, Guid("42F85136-DB7E-439C-85F1-E4075D135FC8"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOpenDialog {
        [PreserveSig] int Show(IntPtr hwnd);
        void SetFileTypes(uint c, IntPtr filters);
        void SetFileTypeIndex(uint i);
        void GetFileTypeIndex(out uint i);
        void Advise(IntPtr pfde, out uint cookie);
        void Unadvise(uint cookie);
        void SetOptions(uint fos);
        void GetOptions(out uint fos);
        void SetDefaultFolder(IShellItem psi);
        void SetFolder(IShellItem psi);
        void GetFolder(out IShellItem ppsi);
        void GetCurrentSelection(out IShellItem ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string name);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string name);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string title);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string text);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string label);
        void GetResult(out IShellItem ppsi);
        void AddPlace(IShellItem psi, int fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string ext);
        void Close(int hr);
        void SetClientGuid(ref Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr filter);
        void GetResults(out IntPtr ppenum);
        void GetSelectedItems(out IntPtr ppsai);
    }

    [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdn, [MarshalAs(UnmanagedType.LPWStr)] out string name);
        void GetAttributes(uint mask, out uint attribs);
        void Compare(IShellItem psi, uint hint, out int order);
    }

    [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
    private class FileOpenDialogCoClass { }

    public static string Show(string title) {
        var dialog = (IFileOpenDialog)new FileOpenDialogCoClass();
        dialog.SetOptions(FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
        dialog.SetTitle(title);
        int hr = dialog.Show(IntPtr.Zero);
        if (hr != 0) return null;
        dialog.GetResult(out IShellItem item);
        item.GetDisplayName(SIGDN_FILESYSPATH, out string path);
        return path;
    }
}
'@

$basePath = [ModernFolderPicker]::Show("Selecione a pasta onde deseja instalar o AutoMail")

if ($basePath) {
    if ($basePath -match "Auto[\-\s]?Mail$") {
        $installDir = $basePath
    } else {
        $installDir = Join-Path $basePath "AutoMail"
    }
} else {
    Write-Host "`n[!] Instalação cancelada pelo usuário. O script será encerrado." -ForegroundColor Red
    exit 1
}

$tempZip = "$env:TEMP\AutoMail.zip"

Write-Host "Pasta de Destino: $installDir" -ForegroundColor DarkGray
Write-Host "Baixando da Branch: $branch" -ForegroundColor DarkGray
Write-Host ""

# 2. Verificar/Instalar Python
Write-Host "[*] Verificando se o Python está instalado..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    Write-Host "[OK] Python encontrado: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "[!] Python não encontrado. Tentando instalar via winget..." -ForegroundColor Red
    try {
        Start-Process -FilePath "winget" -ArgumentList "install --id Python.Python.3.12 --source winget --silent" -Wait -NoNewWindow
        Write-Host "[OK] Python instalado com sucesso." -ForegroundColor Green
        
        # Recarregar PATH para a sessão atual
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
        
    } catch {
        Write-Host "[ERRO] Não foi possível instalar o Python automaticamente." -ForegroundColor Red
        Write-Host "Baixe e instale manualmente em: https://www.python.org/downloads/" -ForegroundColor Red
        exit 1
    }
}

# 3. Baixar e extrair o projeto
Write-Host "`n[*] Baixando a versão mais recente do GitHub..." -ForegroundColor Yellow
try {
    Invoke-WebRequest -Uri $repoZipUrl -OutFile $tempZip -UseBasicParsing
    Write-Host "[OK] Download concluído." -ForegroundColor Green
} catch {
    Write-Host "[ERRO] Falha ao baixar o repositório." -ForegroundColor Red
    Write-Host "    - Verifique sua conexão com a internet." -ForegroundColor Red
    Write-Host "    - O script está configurado para baixar de: $repoZipUrl" -ForegroundColor Red
    Write-Host "    - Certifique-se de que o repositório está público ou acessível." -ForegroundColor Red
    exit 1
}

Write-Host "`n[*] Extraindo arquivos para $installDir..." -ForegroundColor Yellow
if (-not (Test-Path $installDir)) {
    New-Item -ItemType Directory -Force -Path $installDir | Out-Null
}

$extractDir = "$env:TEMP\AutoMail_Extract"
if (Test-Path $extractDir) { Remove-Item -Recurse -Force $extractDir }
Expand-Archive -Path $tempZip -DestinationPath $extractDir -Force

# Mover arquivos da pasta extraída (que geralmente se chama RepoName-branch)
$extractedFolder = Get-ChildItem -Path $extractDir | Select-Object -First 1
Copy-Item -Path "$($extractedFolder.FullName)\*" -Destination $installDir -Recurse -Force

# Limpeza temp
Remove-Item $tempZip -Force
Remove-Item -Recurse -Force $extractDir

Write-Host "[OK] Extração e atualização concluídas." -ForegroundColor Green

# 4. Checagem de Configurações
if (-not (Test-Path "$installDir\config.json")) {
    if (Test-Path "$installDir\config.example.json") {
        Copy-Item "$installDir\config.example.json" "$installDir\config.json"
        Write-Host "`n[!] O arquivo 'config.json' foi criado a partir do modelo." -ForegroundColor DarkYellow
        Write-Host "[!] AVISO: É necessário editá-lo com suas credenciais de email antes da primeira execução." -ForegroundColor Red
    }
}

# 5. Instalar Dependências
Write-Host "`n[*] Verificando e Instalando dependências do Python..." -ForegroundColor Yellow
Set-Location -Path $installDir

python -m pip install --upgrade pip --quiet --no-warn-script-location
python -m pip install pywin32 py7zr --quiet --no-warn-script-location

# pywin32 post-install workaround para rodar silencioso
try {
    $pyPrefix = python -c "import sys; print(sys.prefix)"
    python "$pyPrefix\Scripts\pywin32_postinstall.py" -install *>$null
} catch {}

Write-Host "[OK] Dependências instaladas e configuradas." -ForegroundColor Green

# 6. Criar Atalho no Desktop
$desktopPath = [Environment]::GetFolderPath('Desktop')
$shortcutPath = "$desktopPath\AutoMail.lnk"
if (-not (Test-Path $shortcutPath)) {
    Write-Host "`n[*] Criando atalho na Área de Trabalho..." -ForegroundColor Yellow
    if (-not (Test-Path $desktopPath)) {
        Write-Host "[!] Área de Trabalho não encontrada em '$desktopPath'. Atalho não criado." -ForegroundColor DarkYellow
    } else {
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath)
        $Shortcut.TargetPath = "$installDir\AutoMail.bat"
        $Shortcut.WorkingDirectory = $installDir
        $Shortcut.IconLocation = "$installDir\AutoMail.bat"
        $Shortcut.Save()
        Write-Host "[OK] Atalho 'AutoMail' criado na Área de Trabalho." -ForegroundColor Green
    }
}

Write-Host "`n=================================================" -ForegroundColor Cyan
Write-Host "   Instalação/Atualização Concluída com Sucesso! " -ForegroundColor Green
Write-Host "   Acesse a pasta da sua aplicação em: " -ForegroundColor Cyan
Write-Host "   $installDir" -ForegroundColor White
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "Pressione qualquer tecla para sair..." -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
