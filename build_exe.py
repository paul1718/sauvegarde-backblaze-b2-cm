import subprocess
import sys
from pathlib import Path
import customtkinter


def build() -> None:
    app_dir = Path(__file__).resolve().parent
    ctk_path = Path(customtkinter.__file__).resolve().parent

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--onedir",
        "--windowed",
        "--name",
        "BackblazeB2Uploader",
        "--add-data",
        f"{ctk_path};customtkinter/",
        "--hidden-import",
        "openpyxl",
        "--hidden-import",
        "pandas",
        "--hidden-import",
        "b2sdk",
        str(app_dir / "app.py"),
    ]

    print("Compilation de l'exécutable Windows avec PyInstaller...")
    subprocess.run(cmd, check=True, cwd=str(app_dir))
    print("\nCompilation terminée avec succès !")
    print(f"Exécutable disponible dans : {app_dir / 'dist' / 'BackblazeB2Uploader' / 'BackblazeB2Uploader.exe'}")


if __name__ == "__main__":
    build()
