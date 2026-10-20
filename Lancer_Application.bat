@echo off
chcp 65001 >nul
cd /d "%~dp0"
where pythonw >nul 2>nul
if %errorlevel% equ 0 (
    start "" pythonw app.py
) else (
    if exist "dist\BackblazeB2Uploader\BackblazeB2Uploader.exe" (
        start "" "dist\BackblazeB2Uploader\BackblazeB2Uploader.exe"
    ) else (
        start "" python app.py
    )
)
