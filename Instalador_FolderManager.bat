@echo off
title Instalador FolderManager - Sistema Inteligente
color 0A
cls

echo.
echo ╔══════════════════════════════════════════════════════════════════════════════╗
echo ║                    🚀 INSTALADOR FOLDER MANAGER 🚀                          ║
echo ║                        Sistema Inteligente v1.0                             ║
echo ╚══════════════════════════════════════════════════════════════════════════════╝
echo.

REM ═══════════════════════════════════════════════════════════════════════════════
REM DETECÇÃO AUTOMÁTICA DO SISTEMA
REM ═══════════════════════════════════════════════════════════════════════════════

echo 🔍 Detectando sistema...
echo.

REM Detectar arquitetura do processador
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    set "ARCH=64-bit"
    set "ARCH_TYPE=x64"
) else if "%PROCESSOR_ARCHITECTURE%"=="x86" (
    set "ARCH=32-bit"
    set "ARCH_TYPE=x86"
) else (
    set "ARCH=Desconhecida"
    set "ARCH_TYPE=unknown"
)

REM Detectar versão do Windows
for /f "tokens=4-5 delims=. " %%i in ('ver') do set VERSION=%%i.%%j
if "%version%" == "10.0" (
    set "WIN_VER=Windows 10/11"
    set "COMPATIBLE=SIM"
) else if "%version%" == "6.3" (
    set "WIN_VER=Windows 8.1"
    set "COMPATIBLE=PROVAVEL"
) else if "%version%" == "6.1" (
    set "WIN_VER=Windows 7"
    set "COMPATIBLE=LIMITADA"
) else (
    set "WIN_VER=Versão Antiga"
    set "COMPATIBLE=NAO"
)

echo ✅ Sistema detectado:
echo    📊 Arquitetura: %ARCH% (%ARCH_TYPE%)
echo    🖥️  Sistema: %WIN_VER%
echo    ✔️  Compatibilidade: %COMPATIBLE%
echo.

REM ═══════════════════════════════════════════════════════════════════════════════
REM VERIFICAÇÃO DE COMPATIBILIDADE
REM ═══════════════════════════════════════════════════════════════════════════════

if "%COMPATIBLE%"=="NAO" (
    echo ❌ AVISO: Seu sistema pode não ser compatível!
    echo    💡 Recomendamos atualizar para Windows 10 ou superior.
    echo.
    echo    Deseja continuar mesmo assim? (S/N)
    set /p "CONTINUE="
    if /i not "%CONTINUE%"=="S" (
        echo.
        echo 🚪 Instalação cancelada pelo usuário.
        pause
        exit /b 1
    )
)

if "%ARCH_TYPE%"=="x86" (
    echo ⚠️  ATENÇÃO: Sistema 32-bit detectado!
    echo    📦 Será usado o executável de compatibilidade.
    echo.
)

REM ═══════════════════════════════════════════════════════════════════════════════
REM VERIFICAÇÃO DE ARQUIVOS
REM ═══════════════════════════════════════════════════════════════════════════════

echo 🔍 Verificando arquivos necessários...

if not exist "dist\FolderManager_ANFAVEA.exe" (
    echo.
    echo    ❌ Erro: Executável ANFAVEA não encontrado!
    echo    📁 Procurando: dist\FolderManager_ANFAVEA.exe
    echo.
    echo    💡 Solução:
    echo       1. Execute: pyinstaller --onefile --windowed --icon=icone.ico --name="FolderManager_ANFAVEA" --add-data "empresas.csv;." apk.py
    echo       2. Ou use o executável básico: dist\FolderManager.exe (requer empresas.csv separado)
    echo.
    pause
    exit /b 1
)

REM O executável completo já inclui o CSV embutido - não precisa verificar empresas.csv separadamente

echo ✅ Todos os arquivos necessários encontrados!
echo.

REM ═══════════════════════════════════════════════════════════════════════════════
REM INSTALAÇÃO
REM ═══════════════════════════════════════════════════════════════════════════════

echo 📦 Iniciando instalação...
echo.

REM Criar pasta de destino
set "INSTALL_DIR=%USERPROFILE%\Desktop\FolderManager"
if not exist "%INSTALL_DIR%" (
    mkdir "%INSTALL_DIR%"
    echo ✅ Pasta criada: %INSTALL_DIR%
)

REM Copiar arquivos
echo 📋 Copiando arquivos...
copy "dist\FolderManager_ANFAVEA.exe" "%INSTALL_DIR%\" >nul 2>&1
if errorlevel 1 (
    echo ❌ Erro ao copiar executável!
    pause
    exit /b 1
)

REM O CSV já está embutido no executável completo

if exist "icone.ico" (
    copy "icone.ico" "%INSTALL_DIR%\" >nul 2>&1
)

echo ✅ Arquivos copiados com sucesso!
echo.

REM ═══════════════════════════════════════════════════════════════════════════════
REM CRIAR ATALHO NA ÁREA DE TRABALHO
REM ═══════════════════════════════════════════════════════════════════════════════

echo 🔗 Criando atalho na área de trabalho...

REM Criar script VBS para atalho
echo Set oWS = WScript.CreateObject("WScript.Shell") > "%TEMP%\CreateShortcut.vbs"
echo sLinkFile = "%USERPROFILE%\Desktop\Folder Manager.lnk" >> "%TEMP%\CreateShortcut.vbs"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%TEMP%\CreateShortcut.vbs"
echo oLink.TargetPath = "%INSTALL_DIR%\FolderManager_ANFAVEA.exe" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.WorkingDirectory = "%INSTALL_DIR%" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.Description = "Gerador de Pastas de NF - Folder Manager" >> "%TEMP%\CreateShortcut.vbs"
if exist "%INSTALL_DIR%\icone.ico" (
    echo oLink.IconLocation = "%INSTALL_DIR%\icone.ico" >> "%TEMP%\CreateShortcut.vbs"
)
echo oLink.Save >> "%TEMP%\CreateShortcut.vbs"

cscript "%TEMP%\CreateShortcut.vbs" >nul 2>&1
del "%TEMP%\CreateShortcut.vbs" >nul 2>&1

echo ✅ Atalho criado na área de trabalho!
echo.

REM ═══════════════════════════════════════════════════════════════════════════════
REM TESTE DE FUNCIONAMENTO
REM ═══════════════════════════════════════════════════════════════════════════════

echo 🧪 Deseja testar o aplicativo agora? (S/N)
set /p "TEST="
if /i "%TEST%"=="S" (
    echo.
    echo 🚀 Iniciando Folder Manager...
    start "" "%INSTALL_DIR%\FolderManager_ANFAVEA.exe"
    timeout /t 3 /nobreak >nul
    echo ✅ Aplicativo iniciado! Verifique se abriu corretamente.
)

echo.
echo ╔══════════════════════════════════════════════════════════════════════════════╗
echo ║                           ✅ INSTALAÇÃO CONCLUÍDA! ✅                        ║
echo ╚══════════════════════════════════════════════════════════════════════════════╝
echo.
echo 📁 Localização: %INSTALL_DIR%
echo 🔗 Atalho: Área de trabalho
echo 🚀 Para usar: Clique duas vezes no atalho "Folder Manager"
echo.
echo 💡 DICAS:
echo    • O aplicativo funciona offline
echo    • Não precisa de internet
echo    • Pode ser copiado para outros PCs Windows %ARCH%
echo.

if "%COMPATIBLE%"=="LIMITADA" (
    echo ⚠️  AVISO: Se houver problemas, atualize o Windows ou instale:
    echo    📦 Visual C++ Redistributable para Visual Studio 2015-2022
    echo.
)

echo Pressione qualquer tecla para finalizar...
pause >nul

REM Limpar variáveis
set "ARCH="
set "ARCH_TYPE="
set "WIN_VER="
set "COMPATIBLE="
set "INSTALL_DIR="
set "CONTINUE="
set "TEST="

exit /b 0