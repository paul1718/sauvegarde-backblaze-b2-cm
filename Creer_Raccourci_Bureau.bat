@echo off
chcp 65001 >nul
cd /d "%~dp0"
python create_shortcut.py
echo.
echo Le raccourci "Sauvegarde Backblaze B2" avec son logo a été placé sur le Bureau !
pause
