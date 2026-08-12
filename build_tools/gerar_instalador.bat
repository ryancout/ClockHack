@echo off
setlocal
cd /d "%~dp0"

if not exist "..\dist\FASJornada.exe" (
    echo O executavel nao foi encontrado em ..\dist\FASJornada.exe
    echo Gere o EXE antes de criar o instalador.
    pause
    exit /b 1
)

echo Gerando instalador...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "FASJornada.iss"
if errorlevel 1 (
    echo Falha ao gerar o instalador.
    pause
    exit /b 1
)

echo Instalador gerado em ..\dist
pause
