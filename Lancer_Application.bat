@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist "dist\BackblazeB2Uploader\BackblazeB2Uploader.exe" (
    start "" "dist\BackblazeB2Uploader\BackblazeB2Uploader.exe"
) else (
    start "" pythonw app.py
)
