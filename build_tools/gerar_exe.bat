@echo off
cd /d %~dp0..
pyinstaller --noconfirm --clean "Processador de Planilhas FAS V6.spec"
pause
