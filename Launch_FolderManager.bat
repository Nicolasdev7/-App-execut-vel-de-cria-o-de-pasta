@echo off
echo.
echo ========================================
echo    🗂️  FOLDER MANAGER - INICIANDO...
echo ========================================
echo.

REM Verifica se o executável existe
if not exist "dist\FolderManager.exe" (
    echo ❌ Executável não encontrado!
    echo 📁 Procurando por: dist\FolderManager.exe
    echo.
    echo 🛠️  Para gerar o executável, execute:
    echo    pyinstaller --onefile --windowed --icon=icone.ico --name="FolderManager" apk.py
    echo.
    pause
    exit /b 1
)

echo ✅ Executável encontrado!
echo 🚀 Iniciando Folder Manager...
echo.

REM Executa o aplicativo
start "" "dist\FolderManager.exe"

if errorlevel 1 (
    echo.
    echo ❌ Erro ao executar o aplicativo!
    echo 💡 Verifique se o arquivo FolderManager.exe não está corrompido.
    pause
    exit /b 1
)

echo ✅ Aplicativo iniciado com sucesso!
echo 💡 O Folder Manager está rodando em segundo plano.
echo.
timeout /t 3 /nobreak >nul