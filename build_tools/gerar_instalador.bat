@echo off
setlocal
cd /d %~dp0

if not exist "..\dist\ProcessadorPlanilhasFAS.exe" (
    echo O executavel nao foi encontrado em dist\ProcessadorPlanilhasFAS.exe
    echo Gere o EXE antes de criar o instalador.
    pause
    exit /b 1
)

set "ISCC_PATH=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist "%ISCC_PATH%" set "ISCC_PATH=C:\Program Files\Inno Setup 6\ISCC.exe"

if not exist "%ISCC_PATH%" (
    echo Inno Setup 6 nao encontrado.
    echo Instale o Inno Setup 6 e tente novamente.
    pause
    exit /b 1
)

"%ISCC_PATH%" ProcessadorPlanilhasFAS.iss
pause
