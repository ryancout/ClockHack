@echo off
setlocal
cd /d "%~dp0.."

echo [1/3] Limpando build anterior...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [2/3] Gerando executavel...
pyinstaller --noconfirm main.spec
if errorlevel 1 (
    echo Falha ao gerar o executavel.
    pause
    exit /b 1
)

echo [3/3] Executavel gerado em dist\ProcessadorPlanilhasFAS.exe
pause
