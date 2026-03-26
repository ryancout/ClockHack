@echo off
cd /d %~dp0..
pyinstaller --noconfirm --clean main.spec
pause
