@echo off
cd /d %~dp0..
<<<<<<< HEAD
pyinstaller --noconfirm --clean main.spec
=======
pyinstaller --noconfirm --clean "Processador de Planilhas FAS V6.spec"
>>>>>>> b23b4f0fc185652037e9b32f404393c6f1acc595
pause
