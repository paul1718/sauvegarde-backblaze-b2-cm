import subprocess
from pathlib import Path
from generate_icon import create_app_icon


def create_shortcuts() -> None:
    app_dir = Path(__file__).resolve().parent
    ico_path = create_app_icon()

    bat_path = app_dir / "Lancer_Application.bat"
    exe_path = app_dir / "dist" / "BackblazeB2Uploader" / "BackblazeB2Uploader.exe"

    # Privilégier Lancer_Application.bat qui lance instantanément la dernière version (avec icône officielle)
    target_path = bat_path if bat_path.exists() else exe_path

    ps_script = f"""
$WshShell = New-Object -ComObject WScript.Shell
$DesktopDir = [Environment]::GetFolderPath("Desktop")
$ShortcutPaths = @(
    (Join-Path $DesktopDir "Sauvegarde Backblaze B2.lnk"),
    (Join-Path "{app_dir}" "Sauvegarde Backblaze B2.lnk")
)

foreach ($lnk in $ShortcutPaths) {{
    if (Test-Path $lnk) {{ Remove-Item $lnk -Force -ErrorAction SilentlyContinue }}
    $Shortcut = $WshShell.CreateShortcut($lnk)
    $Shortcut.TargetPath = "{target_path}"
    $Shortcut.WorkingDirectory = "{app_dir}"
    $Shortcut.IconLocation = "{ico_path},0"
    $Shortcut.WindowStyle = 7
    $Shortcut.Description = "Sauvegarde Automatisée Backblaze B2 — Vos fichiers, sauvegardés automatiquement."
    $Shortcut.Save()
    Write-Output "Raccourci mis à jour avec la nouvelle charte graphique : $lnk"
}}
ie4uinit.exe -show
"""
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
        check=True,
    )


if __name__ == "__main__":
    create_shortcuts()
