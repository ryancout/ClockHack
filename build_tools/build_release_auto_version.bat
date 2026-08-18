@echo off
setlocal
cd /d "%~dp0\.."

echo A versao e controlada por app\core\version.py.
echo Este processo valida metadados, testa e gera artefatos sem executar git add, commit ou push.

powershell -NoProfile -ExecutionPolicy Bypass -File "build_tools\build_release.ps1"
if errorlevel 1 (
    echo Falha ao gerar a release.
    pause
    exit /b 1
)

echo Release local concluida em dist\
pause
endlocal
