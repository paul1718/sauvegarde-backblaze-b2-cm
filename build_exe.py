import subprocess
import sys
import time
from pathlib import Path
import customtkinter
from generate_icon import create_app_icon


def build() -> None:
    app_dir = Path(__file__).resolve().parent
    ctk_path = Path(customtkinter.__file__).resolve().parent
    ico_path = create_app_icon()
    assets_dir = app_dir / "assets"

    # Fermer proprement toute instance ouverte de BackblazeB2Uploader.exe pour éviter WinError 5 (Accès refusé)
    subprocess.run(
        ["taskkill", "/F", "/IM", "BackblazeB2Uploader.exe"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )
    time.sleep(0.5)

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--onedir",
        "--windowed",
        "--name",
        "BackblazeB2Uploader",
        "--icon",
        str(ico_path),
        "--add-data",
        f"{ctk_path};customtkinter/",
        "--add-data",
        f"{assets_dir};assets/",
        "--hidden-import",
        "openpyxl",
        "--hidden-import",
        "pandas",
        "--hidden-import",
        "b2sdk",
        "--hidden-import",
        "PIL.ImageTk",
        str(app_dir / "app.py"),
    ]

    print("Compilation de l'exécutable Windows avec PyInstaller et son icône officielle...")
    subprocess.run(cmd, check=True, cwd=str(app_dir))
    print("\nCompilation terminée avec succès !")
    print(f"Exécutable disponible dans : {app_dir / 'dist' / 'BackblazeB2Uploader' / 'BackblazeB2Uploader.exe'}")


if __name__ == "__main__":
    build()
