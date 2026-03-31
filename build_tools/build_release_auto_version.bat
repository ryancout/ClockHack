@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

cd /d "%~dp0\.."

set APP_NAME=ProcessadorPlanilhasFAS

echo.
set /p VERSION=Digite a versao (ex: 6.2.1): 

if "%VERSION%"=="" (
    echo Versao nao informada.
    pause
    exit /b 1
)

echo.
echo ================================
echo   BUILD AUTOMATIZADO - %APP_NAME%
echo   VERSAO %VERSION%
echo ================================
echo.

set ZIP_NAME=%APP_NAME%_v%VERSION%.zip
set COMMIT_MSG=Release v%VERSION% - build automatizado

echo [1/8] Atualizando app/core/version.py...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"(Get-Content 'app/core/version.py') -replace 'APP_VERSION\s*=\s*\"[^\"]+\"', 'APP_VERSION = \"%VERSION%\"' | Set-Content 'app/core/version.py'"

if errorlevel 1 (
    echo ERRO ao atualizar app/core/version.py
    pause
    exit /b 1
)

echo [2/8] Atualizando version_info.txt...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"$v='%VERSION%'.Split('.');" ^
"$maj=[int]$v[0]; $min=[int]$v[1]; $pat=[int]$v[2];" ^
"$content = @' 
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({0},{1},{2},0),
    prodvers=({0},{1},{2},0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '040904B0',
        [
          StringStruct('CompanyName', 'FAS'),
          StringStruct('FileDescription', 'Processador de Planilhas de Horas'),
          StringStruct('FileVersion', '{3}'),
          StringStruct('InternalName', 'ProcessadorPlanilhasFAS'),
          StringStruct('OriginalFilename', 'ProcessadorPlanilhasFAS.exe'),
          StringStruct('ProductName', 'Processador Planilhas FAS'),
          StringStruct('ProductVersion', '{3}')
        ]
      )
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
'@ -f $maj,$min,$pat,'%VERSION%';" ^
"Set-Content 'version_info.txt' $content -Encoding UTF8"

if errorlevel 1 (
    echo ERRO ao atualizar version_info.txt
    pause
    exit /b 1
)

echo [3/8] Limpando build antigo...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "%ZIP_NAME%" del /f /q "%ZIP_NAME%"

echo [4/8] Gerando EXE...
pyinstaller main.spec
if errorlevel 1 (
    echo ERRO ao gerar o EXE.
    pause
    exit /b 1
)

echo [5/8] Verificando EXE...
if not exist "dist\%APP_NAME%.exe" (
    echo EXE nao encontrado em dist\%APP_NAME%.exe
    pause
    exit /b 1
)

echo [6/8] Criando ZIP...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'dist\%APP_NAME%.exe' -DestinationPath '%ZIP_NAME%' -Force"
if errorlevel 1 (
    echo ERRO ao criar o ZIP.
    pause
    exit /b 1
)

echo [7/8] Git add / commit / push...
git add .
git commit -m "%COMMIT_MSG%"
git push origin main
if errorlevel 1 (
    echo ERRO no push para o GitHub.
    pause
    exit /b 1
)

echo [8/8] Finalizado com sucesso.
echo.
echo Versao aplicada: %VERSION%
echo EXE: dist\%APP_NAME%.exe
echo ZIP: %ZIP_NAME%
echo.

pause
endlocal