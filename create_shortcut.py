import subprocess
from pathlib import Path
from generate_icon import create_app_icon


def create_shortcuts() -> None:
    app_dir = Path(__file__).resolve().parent
    ico_path = create_app_icon()

    exe_path = app_dir / "dist" / "BackblazeB2Uploader" / "BackblazeB2Uploader.exe"
    bat_path = app_dir / "Lancer_Application.bat"

    target_path = exe_path if exe_path.exists() else bat_path

    # Script PowerShell utilisant WScript.Shell pour créer le raccourci sur le Bureau et dans le dossier projet
    ps_script = f"""
$WshShell = New-Object -ComObject WScript.Shell
$DesktopDir = [Environment]::GetFolderPath("Desktop")
$ShortcutPaths = @(
    (Join-Path $DesktopDir "Sauvegarde Backblaze B2.lnk"),
    (Join-Path "{app_dir}" "Sauvegarde Backblaze B2.lnk")
)

foreach ($lnk in $ShortcutPaths) {{
    $Shortcut = $WshShell.CreateShortcut($lnk)
    $Shortcut.TargetPath = "{target_path}"
    $Shortcut.WorkingDirectory = "{app_dir}"
    $Shortcut.IconLocation = "{ico_path},0"
    $Shortcut.Description = "Application de Sauvegarde Automatisée Backblaze B2"
    $Shortcut.Save()
    Write-Output "Raccourci créé : $lnk"
}}
"""
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
        check=True,
    )


if __name__ == "__main__":
    create_shortcuts()
