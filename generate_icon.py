import shutil
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont


def create_app_icon() -> Path:
    """
    Utilise les visuels officiels du dossier `pack_identite_sauvegarde_backblaze_b2`
    pour générer les icônes `.ico` multi-résolutions Windows, copier les logos dans `assets/`
    et préparer une image standard par défaut pour les dossiers TAV vides si aucune n'a encore été fournie.
    """
    app_dir = Path(__file__).resolve().parent
    pack_dir = app_dir / "pack_identite_sauvegarde_backblaze_b2"
    assets_dir = app_dir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)

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

    # Création d'une image standard TAV par défaut si l'utilisateur n'en a pas encore sélectionné une
    std_tav_img_path = assets_dir / "image_standard_tav.png"
    if not std_tav_img_path.exists():
        std_img = Image.new("RGB", (800, 600), (242, 246, 252))
        draw = ImageDraw.Draw(std_img)
        draw.rounded_rectangle([40, 40, 760, 560], radius=24, fill=(255, 255, 255), outline=(8, 120, 249), width=4)
        try:
            f_title = ImageFont.truetype("arialbd.ttf", 36)
            f_sub = ImageFont.truetype("arial.ttf", 22)
        except Exception:
            f_title = ImageFont.load_default()
            f_sub = ImageFont.load_default()
        draw.text((220, 240), "DOSSIER TAV STANDARD", fill=(16, 42, 94), font=f_title)
        draw.text((195, 310), "Aucune image spécifique disponible dans ce dossier.", fill=(82, 96, 120), font=f_sub)
        std_img.save(std_tav_img_path, format="PNG")

    src_icon_png = assets_dir / "icone_1024.png"
    if not src_icon_png.exists():
        src_icon_png = assets_dir / "icone_256.png"

    ico_path = assets_dir / "app_icon.ico"
    ico_b2_path = assets_dir / "icone_application_b2.ico"

    if src_icon_png.exists():
        img = Image.open(src_icon_png).convert("RGBA")
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
