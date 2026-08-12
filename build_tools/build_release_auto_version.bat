@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0\.."

set APP_NAME=FASJornada

echo.
set /p VERSION=Digite a versao (ex: 6.2.1): 

if "%VERSION%"=="" (
    echo Versao nao informada.
    pause
    exit /b 1
)

set RELEASE_DIR=releases
set ZIP_NAME=%RELEASE_DIR%\%APP_NAME%_v%VERSION%.zip
set COMMIT_MSG=Release v%VERSION% - build automatizado

echo.
echo ================================
echo   BUILD - %APP_NAME% v%VERSION%
echo ================================
echo.

echo [1/7] Atualizando version.py...
echo APP_NAME = "FAS Jornada" > app\core\version.py
echo APP_VERSION = "%VERSION%" >> app\core\version.py

echo [2/7] Atualizando version_info.txt...
(
echo VSVersionInfo(
echo   ffi=FixedFileInfo(
echo     filevers=(%VERSION:.=,%,0),
echo     prodvers=(%VERSION:.=,%,0),
echo     mask=0x3f,
echo     flags=0x0,
echo     OS=0x40004,
echo     fileType=0x1,
echo     subtype=0x0,
echo     date=(0, 0)
echo   ),
echo   kids=[
echo     StringFileInfo([
echo       StringTable(
echo         '040904B0',
echo         [
echo           StringStruct('CompanyName', 'FAS'),
echo           StringStruct('FileDescription', 'Relatório e Análise de Jornada'),
echo           StringStruct('FileVersion', '%VERSION%'),
echo           StringStruct('InternalName', 'FASJornada'),
echo           StringStruct('OriginalFilename', 'FASJornada.exe'),
echo           StringStruct('ProductName', 'FAS Jornada'),
echo           StringStruct('ProductVersion', '%VERSION%')
echo         ]
echo       )
echo     ]),
echo     VarFileInfo([VarStruct('Translation', [1033, 1200])])
echo   ]
echo )
) > version_info.txt

echo [3/7] Preparando pasta releases...
if not exist %RELEASE_DIR% mkdir %RELEASE_DIR%
del /f /q %RELEASE_DIR%\%APP_NAME%_v*.zip 2>nul

echo [4/7] Limpando build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [5/7] Gerando EXE...
pyinstaller main.spec

if errorlevel 1 (
    echo ERRO ao gerar EXE.
    pause
    exit /b 1
)

if not exist "dist\%APP_NAME%.exe" (
    echo ERRO: EXE nao encontrado.
    pause
    exit /b 1
)

echo [6/7] Criando ZIP...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'dist\%APP_NAME%.exe' -DestinationPath '%ZIP_NAME%' -Force"

if errorlevel 1 (
    echo ERRO ao criar ZIP.
    pause
    exit /b 1
)

echo [7/7] Git commit/push...
git add .
git commit -m "%COMMIT_MSG%"

if errorlevel 1 (
    echo AVISO: Nenhuma alteracao para commit ou erro no commit.
)

git push origin main

if errorlevel 1 (
    echo ERRO no git push.
    pause
    exit /b 1
)

echo.
echo ================================
echo FINALIZADO COM SUCESSO
echo ================================
echo.
echo ZIP GERADO:
echo %ZIP_NAME%
echo.

pause
endlocal
