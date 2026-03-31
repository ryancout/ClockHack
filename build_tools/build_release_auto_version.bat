@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0\.."

set APP_NAME=ProcessadorPlanilhasFAS

echo.
set /p VERSION=Digite a versao (ex: 6.2.1): 

if "%VERSION%"=="" (
    echo Versao nao informada.
    pause
    exit /b 1
)

set ZIP_NAME=%APP_NAME%_v%VERSION%.zip
set COMMIT_MSG=Release v%VERSION% - build automatizado

echo.
echo ================================
echo   BUILD - %APP_NAME% v%VERSION%
echo ================================
echo.

echo [1/7] Atualizando version.py...
powershell -Command "(Get-Content 'app/core/version.py') -replace 'APP_VERSION = \".*\"', 'APP_VERSION = \"%VERSION%\"' | Set-Content 'app/core/version.py'"

echo [2/7] Atualizando version_info.txt...
powershell -Command ^
"$v='%VERSION%'.Split('.');" ^
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
'@ -f $v[0],$v[1],$v[2],'%VERSION%';" ^
"Set-Content 'version_info.txt' $content"

echo [3/7] Limpando build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [4/7] Gerando EXE...
pyinstaller main.spec

echo [5/7] Criando ZIP...
powershell -Command "Compress-Archive -Path 'dist\%APP_NAME%.exe' -DestinationPath '%ZIP_NAME%' -Force"

echo [6/7] Git commit/push...
git add .
git commit -m "%COMMIT_MSG%"
git push origin main

echo.
echo Finalizado!
echo ZIP: %ZIP_NAME%
echo.

pause