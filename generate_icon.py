import shutil
from pathlib import Path
from PIL import Image


def create_app_icon() -> Path:
    """
    Utilise les visuels officiels du dossier `pack_identite_sauvegarde_backblaze_b2`
    pour générer les icônes `.ico` multi-résolutions Windows et copier les logos dans `assets/`.
    """
    app_dir = Path(__file__).resolve().parent
    pack_dir = app_dir / "pack_identite_sauvegarde_backblaze_b2"
    assets_dir = app_dir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)

    # Copie des ressources officielles du pack dans assets/
    files_to_copy = [
        "logo_principal.png",
        "logo_principal.svg",
        "icone_application.svg",
        "icone_1024.png",
        "icone_512.png",
        "icone_256.png",
        "icone_128.png",
        "icone_64.png",
        "icone_32.png",
    ]
    for fname in files_to_copy:
        src = pack_dir / fname
        if src.exists():
            shutil.copy2(src, assets_dir / fname)

    # Construction de l'icône Windows .ico à partir de l'icône officielle haute définition
    src_icon_png = assets_dir / "icone_1024.png"
    if not src_icon_png.exists():
        src_icon_png = assets_dir / "icone_256.png"

    ico_path = assets_dir / "app_icon.ico"
    ico_b2_path = assets_dir / "icone_application_b2.ico"

    if src_icon_png.exists():
        img = Image.open(src_icon_png).convert("RGBA")
        # Sauvegarde aussi sous app_icon.png
        img.resize((256, 256), Image.Resampling.LANCZOS).save(assets_dir / "app_icon.png", format="PNG")

        for target_ico in (ico_path, ico_b2_path):
            img.save(
                target_ico,
                format="ICO",
                sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)],
            )
        print(f"Icône officielle générée : {ico_b2_path}")

    return ico_b2_path if ico_b2_path.exists() else ico_path


if __name__ == "__main__":
    create_app_icon()
